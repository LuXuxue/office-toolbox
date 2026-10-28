from copy import deepcopy
import asyncio
import os
import time
import cv2
import numpy as np
from pathlib import Path
from typing import Any, Dict, Optional

from oxr.config.settings import settings
from oxr.model.client import model_client
from oxr.utils.pdf import pdf_to_images
from oxr.pipeline.layout import analyze_layouts
from oxr.pipeline.concurrency import gather_cancel_on_error
from oxr.pipeline.errors import PipelineError
from oxr.pipeline.reading_order import resolve_reading_orders
from oxr.pipeline.recognition import recognize_pages
from oxr.pipeline.assembler import assemble_markdown
from oxr.pipeline.models import Block, BlockType
from oxr.pipeline.image_assets import ImageAssets, image_assets
from oxr.pipeline.postprocess.assembly_postprocess import blocks_to_markdown_with_ranges
from oxr.pipeline.postprocess.formula_postprocess import (
    FORMULA_NUMBERING_ROLE, expand_formula_numbering_blocks, space_markdown_math,
)
from oxr.pipeline.postprocess.json_postprocess import serialize_final_blocks
from oxr.pipeline.postprocess.cross_category_dedupe import (
    dedupe_systematic_cross_category_blocks,
)
from oxr.pipeline.postprocess.overlap_dedupe import (
    dedupe_visual_parents,
    dedupe_complete_text,
)
from oxr.pipeline.postprocess.exact_text_dedupe import dedupe_exact_textual_blocks
from oxr.pipeline.postprocess.full_overlap_dedupe import dedupe_full_overlaps
from oxr.pipeline.postprocess.table_images import TABLE_IMAGE_ROLE


INVALID_DOCUMENT_MESSAGE = "Unable to decode document into valid page images."

# Internal processing policy; not configurable through YAML or environment.
_SYSTEMATIC_CROSS_CATEGORY_DEDUPE = True
_VISUAL_PARENT_DEDUPE = True
_COMPLETE_TEXT_DEDUPE = True


def _dedupe_full_page(raw: list[Block] | None, blocks: list[Block], image: np.ndarray) -> list[Block]:
    """Keep final content witnesses for geometry and visual-coverage cleanup."""
    if not _VISUAL_PARENT_DEDUPE or raw is None:
        return dedupe_full_overlaps(blocks, image)
    visual = dedupe_visual_parents(raw, blocks, image, allow_mixed_children=True)
    final = dedupe_full_overlaps(visual, image)
    # A visual parent's coverage proof is valid only while all witnesses survive
    # subsequent cleanup. Recheck it and restore the original parent if needed.
    visual_ids = {id(block) for block in visual}
    final_ids = {id(block) for block in final}
    recheck = [block for block in blocks if id(block) not in visual_ids or id(block) in final_ids]
    return dedupe_visual_parents(raw, recheck, image, allow_mixed_children=True)


def _decode_document_images(file_bytes: bytes, file_name: str) -> list[np.ndarray]:
    ext = os.path.splitext(file_name)[1].lower()
    try:
        if ext == ".pdf":
            images = pdf_to_images(file_bytes)
        else:
            np_arr = np.frombuffer(file_bytes, np.uint8)
            images = [cv2.imdecode(np_arr, cv2.IMREAD_COLOR)]
    except Exception as error:
        raise PipelineError(
            INVALID_DOCUMENT_MESSAGE, public_message=INVALID_DOCUMENT_MESSAGE,
        ) from error

    if not images or any(
        not isinstance(image, np.ndarray)
        or image.ndim not in {2, 3}
        or image.size == 0
        or image.shape[0] == 0
        or image.shape[1] == 0
        for image in images
    ):
        raise PipelineError(INVALID_DOCUMENT_MESSAGE, public_message=INVALID_DOCUMENT_MESSAGE)
    return images


async def process_document(
    file_bytes: bytes,
    file_name: str,
    output_formats: list[str],
    keep_header_footer: bool = False,
    model_options: Optional[Dict[str, Any]] = None,
    element_formats: Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    start_time = time.perf_counter()
    images = await asyncio.to_thread(_decode_document_images, file_bytes, file_name)
    return await process_decoded_document(
        images, file_name, output_formats, keep_header_footer,
        model_options, element_formats, started_at=start_time,
    )


async def process_decoded_document(
    images: list[np.ndarray],
    file_name: str,
    output_formats: list[str],
    keep_header_footer: bool = False,
    model_options: Optional[Dict[str, Any]] = None,
    element_formats: Optional[Dict[str, str]] = None,
    *,
    demo: bool = False,
    asset_dir: Optional[Path] = None,
    asset_url_prefix: str = "",
    started_at: Optional[float] = None,
) -> Dict[str, Any]:
    """Recognize already-decoded pages, optionally collecting demo source maps."""
    assets = ImageAssets(Path(asset_dir), asset_url_prefix) if asset_dir is not None else None
    token = image_assets.set(assets)
    try:
        return await _process_decoded_document(
            images, file_name, output_formats, keep_header_footer,
            model_options, element_formats, demo, started_at,
        )
    finally:
        image_assets.reset(token)


async def _process_decoded_document(
    images, file_name, output_formats, keep_header_footer,
    model_options, element_formats, demo, started_at,
):
    start_time = started_at if started_at is not None else time.perf_counter()
    full_overlap_dedupe = settings.pipeline.full_overlap_dedupe
    ext = os.path.splitext(file_name)[1].lower()
    page_num = len(images)
    decode_end_time = time.perf_counter()

    # 2. Phase 1: Layout Analysis
    pages_blocks = await analyze_layouts(images, model_options)
    layout_end_time = time.perf_counter()
    raw_pages_blocks = (
        deepcopy(pages_blocks) if _VISUAL_PARENT_DEDUPE else None
    )

    filtered_types = {BlockType.PAGE_HEADER, BlockType.PAGE_FOOTER}
    # Demo margins stay outside body reading order and cleanup. Recognition is
    # optional; disabled margins retain the original overlay-only behavior.
    demo_page_margins = [
        [deepcopy(block) for block in blocks if block.type in filtered_types]
        for blocks in pages_blocks
    ] if demo else None
    if demo or not keep_header_footer:
        pages_blocks = [
            [block for block in blocks if block.type not in filtered_types]
            for blocks in pages_blocks
        ]
    
    # 3. Phase 2: Reading order and recognition share the OXR service budget.
    semaphore = model_client.get_request_semaphore()
    recognition_args = (
        pages_blocks,
        images,
        file_name,
        semaphore,
        model_options,
        element_formats,
    )
    async def recognize_body():
        if settings.layout_analyze_model.provides_reading_order:
            return await recognize_pages(*recognition_args)
        reading_order_task = asyncio.create_task(
            resolve_reading_orders(pages_blocks, images, semaphore)
        )
        recognition_task = asyncio.create_task(recognize_pages(*recognition_args))
        reading_orders, recognized_pages = await gather_cancel_on_error(
            reading_order_task,
            recognition_task,
        )
        return [
            [blocks[local_id] for local_id in order]
            for blocks, order in zip(recognized_pages, reading_orders)
        ]

    if demo and keep_header_footer:
        pages_blocks, demo_page_margins = await gather_cancel_on_error(
            recognize_body(),
            recognize_pages(
                demo_page_margins, images, file_name, semaphore,
                model_options, element_formats,
            ),
        )
    else:
        pages_blocks = await recognize_body()
    recognition_end_time = time.perf_counter()

    # Table pictures have already been restored (or retained in the raster
    # fallback). Remove their top-level duplicates only after applying RO.
    pages_blocks = [
        [block for block in blocks if block.generated_role != TABLE_IMAGE_ROLE]
        for blocks in pages_blocks
    ]

    if _SYSTEMATIC_CROSS_CATEGORY_DEDUPE and not full_overlap_dedupe:
        pages_blocks = [
            dedupe_systematic_cross_category_blocks(blocks)
            for blocks in pages_blocks
        ]
    pages_blocks = [expand_formula_numbering_blocks(blocks) for blocks in pages_blocks]
    
    # Cleanup only final blocks; recognition crops and RO inputs stay intact.
    if full_overlap_dedupe:
        pages_blocks = [
            _dedupe_full_page(raw, blocks, image)
            for raw, blocks, image in zip(
                raw_pages_blocks or [None] * len(pages_blocks), pages_blocks, images,
            )
        ]
    else:
        if _VISUAL_PARENT_DEDUPE:
            pages_blocks = [
                dedupe_visual_parents(raw, blocks, image)
                for raw, blocks, image in zip(raw_pages_blocks, pages_blocks, images)
            ]
        if _COMPLETE_TEXT_DEDUPE:
            pages_blocks = [dedupe_complete_text(blocks) for blocks in pages_blocks]
        pages_blocks = [dedupe_exact_textual_blocks(blocks) for blocks in pages_blocks]

    if demo and keep_header_footer:
        for page_index, (blocks, margins) in enumerate(zip(pages_blocks, demo_page_margins)):
            def margin_order(block):
                return block.bbox[1], block.bbox[0], block.idx

            headers = sorted(
                (block for block in margins if block.type == BlockType.PAGE_HEADER),
                key=margin_order,
            )
            footers = sorted(
                (block for block in margins if block.type == BlockType.PAGE_FOOTER),
                key=margin_order,
            )
            # Preserve gaps left by cleanup: formula numbering uses adjacent
            # indexes to associate each surviving witness with its formula.
            for block in blocks:
                block.idx += len(headers)
            footer_start = max((block.idx for block in blocks), default=len(headers) - 1) + 1
            for idx, block in enumerate(headers):
                block.idx = idx
            for offset, block in enumerate(footers):
                block.idx = footer_start + offset
            pages_blocks[page_index] = headers + blocks + footers

    # 4. Phase 3: Assembly & Output Formatting
    outputs = {}
    
    if "json" in output_formats:
        json_pages = []
        for i, blocks in enumerate(pages_blocks):
            h, w = images[i].shape[:2]
            serialized_blocks, numbering_notes = serialize_final_blocks(blocks)
            page_data = {
                "page_num": i,
                "page_size": {"w": w, "h": h},
                "merge_content_from_pre_page": False, # Basic placeholder
                "blocks": serialized_blocks,
            }
            if numbering_notes:
                page_data["unattached_formula_numbering"] = numbering_notes
            json_pages.append(page_data)
        outputs["json"] = {"pages": json_pages}
        
    if "markdown" in output_formats:
        outputs["markdown"] = assemble_markdown(pages_blocks)
        
    end_time = time.perf_counter()

    metadata = {
        "page_num": page_num,
        "file_type": ext,
        "processing_time_ms": int((end_time - start_time) * 1000)
    }

    if settings.debug:
        block_counts: Dict[str, int] = {}
        blocks_per_page = []
        for blocks in pages_blocks:
            regions = [block for block in blocks if block.generated_role != FORMULA_NUMBERING_ROLE]
            blocks_per_page.append(len(regions))
            for block in regions:
                block_counts[block.type.value] = block_counts.get(block.type.value, 0) + 1

        metadata["debug_timing"] = {
            "decode_time_ms": int((decode_end_time - start_time) * 1000),
            "layout_time_ms": int((layout_end_time - decode_end_time) * 1000),
            "recognition_time_ms": int((recognition_end_time - layout_end_time) * 1000),
            "assembly_time_ms": int((end_time - recognition_end_time) * 1000),
        }
        metadata["debug_stats"] = {
            "total_blocks": sum(blocks_per_page),
            "blocks_per_page": blocks_per_page,
            "block_type_counts": dict(sorted(block_counts.items())),
        }
    
    result = {
        "file_name": file_name,
        "outputs": outputs,
        "metadata": metadata
    }

    if demo:
        metadata["keep_header_footer"] = keep_header_footer
        demo_pages = []
        for i, blocks in enumerate(pages_blocks):
            markdown, ranges = blocks_to_markdown_with_ranges(blocks)
            # The math postprocessor adds spaces only; source line ranges remain valid.
            markdown = space_markdown_math(markdown)
            serialized_blocks, numbering_notes = serialize_final_blocks(blocks, ranges)
            h, w = images[i].shape[:2]
            # Formula expansion reindexes content blocks. Give overlay-only boxes
            # separate IDs after those final indices without changing source maps.
            margin_start_idx = max((b.idx for b in blocks), default=-1) + 1
            margin_blocks = [
                {
                    "idx": margin_start_idx + offset,
                    "type": b.type.value, "bbox": b.bbox,
                    "content": "", "format": b.format.value,
                    "start_line": None, "end_line": None,
                }
                for offset, b in enumerate(demo_page_margins[i] if not keep_header_footer else [])
            ]
            demo_page = {
                "page_num": i,
                "page_size": {"w": w, "h": h},
                "markdown": markdown,
                "blocks": sorted(serialized_blocks, key=lambda b: b["idx"]) + margin_blocks,
            }
            if numbering_notes:
                demo_page["unattached_formula_numbering"] = numbering_notes
            demo_pages.append(demo_page)
        result["demo_pages"] = demo_pages
    return result
