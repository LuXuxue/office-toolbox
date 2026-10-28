import asyncio
import logging
from typing import Any, Awaitable, Callable, Dict, List, Optional

import numpy as np

from oxr.config.settings import settings
from oxr.model.client import model_client
from oxr.pipeline.concurrency import gather_cancel_on_error
from oxr.pipeline.errors import PipelineError, format_error_message, get_error_status_code
from oxr.pipeline.models import Block, BlockType
from oxr.pipeline.postprocess.chemical_structure_postprocess import (
    postprocess_chemical_structure,
)
from oxr.pipeline.postprocess.code_postprocess import postprocess_code
from oxr.pipeline.postprocess.formula_postprocess import postprocess_formula
from oxr.pipeline.postprocess.picture_postprocess import postprocess_picture
from oxr.pipeline.postprocess.table_postprocess import postprocess_table
from oxr.pipeline.postprocess.table_images import (
    group_table_images, mark_table_images, prepare_table_images,
    restore_table_images, table_image_fallback, unchanged_block,
)
from oxr.pipeline.postprocess.text_postprocess import postprocess_text
from oxr.utils.image import crop_image, resize_image

logger = logging.getLogger(__name__)

LENGTH_RECOVERY_REPETITION_PENALTY = 1.1
TABLE_LENGTH_RECOVERY_SCALE = 0.75


def _finish_reason(output: str) -> str:
    return str(getattr(output, "finish_reason", "unknown"))


def _length_recovery_options(
    model_options: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    recovered = dict(model_options or {})
    current_penalty = float(recovered.get("repetition_penalty", 1.0))
    recovered["repetition_penalty"] = max(
        current_penalty,
        LENGTH_RECOVERY_REPETITION_PENALTY,
    )
    return recovered


async def _request_with_retries(
    request: Callable[[], Awaitable[str]],
    *,
    block: Block,
    attempts: int,
    timeout: int,
) -> str:
    for attempt in range(attempts):
        try:
            return await request()
        except Exception as error:
            if attempt == attempts - 1:
                message = format_error_message(error, f"Timeout after {timeout}s")
                raise PipelineError(
                    f"Recognition failed for block {block.idx} after {attempts} attempts: {message}",
                    status_code=get_error_status_code(error, default=503),
                ) from error
            await asyncio.sleep(1)

    raise AssertionError("Recognition retry loop exited unexpectedly.")


async def _request_block(
    block: Block,
    crop_img: np.ndarray,
    sem: asyncio.Semaphore,
    model_options: Optional[Dict[str, Any]],
) -> str:
    async with sem:
        if block.type == BlockType.TABLE:
            return await model_client.recognize_table(crop_img, model_options)
        if block.type == BlockType.FORMULA:
            return await model_client.recognize_formula(crop_img, model_options)
        if block.type == BlockType.CODE_BLOCK:
            return await model_client.recognize_code(crop_img, model_options)
        if block.type == BlockType.CHEMICAL_STRUCTURE:
            return await model_client.recognize_chemical_structure(
                crop_img,
                model_options,
            )
        return await model_client.recognize_text(crop_img, model_options)


async def recognize_block(
    block: Block,
    page_image: np.ndarray,
    sem: asyncio.Semaphore,
    file_name: str,
    page_num: int,
    model_options: Optional[Dict[str, Any]] = None,
    element_formats: Optional[Dict[str, str]] = None,
    table_images: Optional[List[Block]] = None,
) -> Block:
    crop_img = crop_image(page_image, block.bbox)
    element_formats = element_formats or {}

    if block.type == BlockType.PICTURE:
        # No OCR for picture, just save it
        return await asyncio.to_thread(postprocess_picture, block, crop_img, file_name, page_num)
    if (
        block.type == BlockType.CHEMICAL_STRUCTURE
        and element_formats.get("cs") == "image"
    ):
        return await asyncio.to_thread(postprocess_picture, block, crop_img, file_name, page_num)

    sources: Dict[int, str] = {}
    if table_images:
        try:
            crop_img, sources = await asyncio.to_thread(prepare_table_images, block, table_images, page_image)
        except ValueError as error:
            logger.warning("Table %s image placeholders skipped: %s; preserving original table image.", block.idx, error)
            mark_table_images(table_images)
            return await asyncio.to_thread(table_image_fallback, block, page_image)

    config = settings.oxr_model
    try:
        raw_out = await _request_with_retries(
            lambda: _request_block(block, crop_img, sem, model_options),
            block=block,
            attempts=config.retry_times,
            timeout=config.timeout,
        )
    except PipelineError:
        if not table_images:
            raise
        logger.warning("Table %s placeholder recognition failed; preserving original table image.", block.idx)
        mark_table_images(table_images)
        return await asyncio.to_thread(table_image_fallback, block, page_image)

    if _finish_reason(raw_out) == "length":
        recovery_image: Optional[np.ndarray] = None
        recovery_options: Optional[Dict[str, Any]] = None
        recovery_method: Optional[str] = None
        if block.type == BlockType.TABLE:
            recovery_image = resize_image(crop_img, TABLE_LENGTH_RECOVERY_SCALE)
            recovery_options = model_options
            recovery_method = f"input scale {TABLE_LENGTH_RECOVERY_SCALE}"
        elif block.type in {
            BlockType.TEXT,
            BlockType.TITLE,
            BlockType.CAPTION,
            BlockType.FOOTNOTE,
            BlockType.PAGE_HEADER,
            BlockType.PAGE_FOOTER,
            BlockType.STAMP,
        }:
            recovery_image = crop_img
            recovery_options = _length_recovery_options(model_options)
            recovery_method = (
                "repetition penalty "
                f"{recovery_options['repetition_penalty']}"
            )
        else:
            logger.warning(
                "Recognition output reached the token limit for block %s (%s); "
                "no recovery is defined, preserving the truncated model output.",
                block.idx,
                block.type.value,
            )

        if recovery_method is not None and recovery_image is not None:
            logger.warning(
                "Recognition output reached the token limit for block %s (%s); "
                "retrying with %s.",
                block.idx,
                block.type.value,
                recovery_method,
            )
            try:
                raw_out = await _request_with_retries(
                    lambda: _request_block(block, recovery_image, sem, recovery_options),
                    block=block,
                    attempts=config.retry_times,
                    timeout=config.timeout,
                )
            except PipelineError:
                if not table_images:
                    raise
                logger.warning("Table %s placeholder recovery failed; preserving original table image.", block.idx)
                mark_table_images(table_images)
                return await asyncio.to_thread(table_image_fallback, block, page_image)
            if _finish_reason(raw_out) == "length":
                logger.warning(
                    "Recognition output remained truncated for block %s (%s) "
                    "after length recovery with %s; preserving the truncated "
                    "model output.",
                    block.idx,
                    block.type.value,
                    recovery_method,
                )

    if block.type == BlockType.TABLE:
        if table_images:
            restored = restore_table_images(raw_out, sources)
            mark_table_images(table_images)
            if restored is None or _finish_reason(raw_out) == "length":
                logger.warning("Table %s has incomplete image IDs/output; preserving original table image.", block.idx)
                return await asyncio.to_thread(table_image_fallback, block, page_image)
            raw_out = restored
        return postprocess_table(
            block,
            raw_out,
            element_formats.get("table", "html"),
        )
    if block.type == BlockType.FORMULA:
        return postprocess_formula(block, raw_out)
    if block.type == BlockType.CODE_BLOCK:
        return postprocess_code(block, raw_out)
    if block.type == BlockType.CHEMICAL_STRUCTURE:
        return postprocess_chemical_structure(block, raw_out)
    return postprocess_text(block, raw_out)


async def recognize_pages(
    pages_blocks: List[List[Block]],
    pages_images: List[np.ndarray],
    file_name: str,
    semaphore: asyncio.Semaphore,
    model_options: Optional[Dict[str, Any]] = None,
    element_formats: Optional[Dict[str, str]] = None,
) -> List[List[Block]]:
    """Run recognition for all blocks across all pages concurrently."""
    if len(pages_blocks) != len(pages_images):
        raise ValueError("Page block and image counts must match")

    tasks = []
    # keep track of which task belongs to which page and block index
    # tasks structure: (page_idx, block_idx, task)
    for page_idx, (blocks, image) in enumerate(zip(pages_blocks, pages_images)):
        groups = (
            group_table_images(
                blocks, settings.pipeline.table_image_min_score,
                keep_outer_tables=settings.pipeline.full_overlap_dedupe,
            )
            if settings.pipeline.table_image_placeholders else {}
        )
        nested = {id(picture) for pictures in groups.values() for picture in pictures}
        for block_idx, block in enumerate(blocks):
            if id(block) in nested:
                task = unchanged_block(block)
            else:
                kwargs = {"table_images": groups[block_idx]} if block_idx in groups else {}
                task = recognize_block(
                    block, image, semaphore, file_name, page_idx,
                    model_options, element_formats, **kwargs,
                )
            tasks.append((page_idx, block_idx, task))

    # Run all tasks
    gathered_results = await gather_cancel_on_error(*(t[2] for t in tasks))

    # Reconstruct pages_blocks
    final_pages_blocks: List[List[Optional[Block]]] = [
        [None] * len(page_blocks) for page_blocks in pages_blocks
    ]
    for (page_idx, block_idx, _), block_result in zip(tasks, gathered_results):
        final_pages_blocks[page_idx][block_idx] = block_result

    return [[block for block in page if block is not None] for page in final_pages_blocks]
