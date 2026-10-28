import asyncio
import hashlib
import logging
import re
from collections import Counter
from typing import Sequence

import numpy as np

from oxr.config.settings import OXRModelConfig, settings
from oxr.model.client import OXRModelClient, model_client
from oxr.model.prompts import build_reading_order_prompt
from oxr.pipeline.concurrency import gather_cancel_on_error
from oxr.pipeline.errors import PipelineError, format_error_message, get_error_status_code
from oxr.pipeline.models import Block
from oxr.pipeline.xycut_order import xycut_plus_order


logger = logging.getLogger(__name__)

_READING_ORDER_PATTERN = re.compile(r"\s*(?:[0-9]+(?:\s+[0-9]+)*)?\s*")


def _reading_order_diagnostics(raw_text: object, block_count: int) -> str:
    """Describe invalid model output without logging document-derived text."""
    if not isinstance(raw_text, str):
        return f"output_type={type(raw_text).__name__}"

    digest = hashlib.sha256(raw_text.encode("utf-8")).hexdigest()[:16]
    if _READING_ORDER_PATTERN.fullmatch(raw_text) is None:
        return f"syntax=invalid char_count={len(raw_text)} raw_sha256={digest}"

    order = [int(value) for value in raw_text.split()]
    counts = Counter(order)
    expected = set(range(block_count))
    actual = set(order)
    missing_ids = sorted(expected - actual)
    duplicate_ids = sorted(value for value, count in counts.items() if count > 1)
    out_of_range_ids = sorted(actual - expected)
    return (
        f"syntax=valid token_count={len(order)} missing_ids={missing_ids} "
        f"duplicate_ids={duplicate_ids} out_of_range_ids={out_of_range_ids} "
        f"raw_sha256={digest}"
    )


def parse_reading_order(raw_text: str, block_count: int) -> list[int]:
    """Parse a strict whitespace-only integer permutation from OXR output."""
    if not isinstance(raw_text, str):
        raise ValueError("Invalid reading order: model output must be text")
    if (
        isinstance(block_count, bool)
        or not isinstance(block_count, int)
        or block_count < 0
    ):
        raise ValueError(
            "Invalid reading order: block count must be a non-negative integer"
        )
    if _READING_ORDER_PATTERN.fullmatch(raw_text) is None:
        raise ValueError(
            "Invalid reading order: expected whitespace-separated integer IDs only"
        )

    order = [int(value) for value in raw_text.split()]
    expected = list(range(block_count))
    if len(order) != block_count or sorted(order) != expected:
        raise ValueError(
            "Invalid reading order: expected an exact permutation of IDs "
            f"0 through {block_count - 1}"
        )
    return order


def _insert_missing_from_reference(
    order: Sequence[int],
    missing_ids: set[int],
    reference_order: Sequence[int],
) -> list[int]:
    """Insert missing IDs while preserving every existing pairwise order.

    XY-Cut++ supplies only the reference ranks for insertion. For each missing
    block, choose the position with the fewest pairwise disagreements against
    that reference. Existing OXR-RO IDs are never reordered.
    """
    reference_rank = {block_id: rank for rank, block_id in enumerate(reference_order)}
    repaired_order = list(order)
    for missing_id in sorted(missing_ids, key=reference_rank.__getitem__):
        missing_rank = reference_rank[missing_id]

        def insertion_cost(index: int) -> tuple[int, int]:
            left_inversions = sum(
                reference_rank[block_id] > missing_rank
                for block_id in repaired_order[:index]
            )
            right_inversions = sum(
                reference_rank[block_id] < missing_rank
                for block_id in repaired_order[index:]
            )
            return left_inversions + right_inversions, index

        insertion_index = min(
            range(len(repaired_order) + 1),
            key=insertion_cost,
        )
        repaired_order.insert(insertion_index, missing_id)
    return repaired_order


def _partial_reading_order(
    raw_text: object,
    block_count: int,
) -> tuple[list[int], set[int], bool] | None:
    """Parse an in-range partial order and retain each ID's first occurrence."""
    if (
        not isinstance(raw_text, str)
        or _READING_ORDER_PATTERN.fullmatch(raw_text) is None
    ):
        return None

    raw_order = [int(value) for value in raw_text.split()]
    expected = set(range(block_count))
    if not set(raw_order).issubset(expected):
        return None

    unique_order = list(dict.fromkeys(raw_order))
    missing_ids = expected - set(unique_order)
    has_duplicates = len(unique_order) != len(raw_order)
    return unique_order, missing_ids, has_duplicates


def repair_incomplete_reading_order(
    raw_text: object,
    blocks: Sequence[Block],
) -> list[int] | None:
    """Repair a nearly complete, valid, unique partial permutation.

    Pages below 20 blocks may be missing one ID; larger pages may be missing two.
    The model's relative order is preserved. XY-Cut++ is used only to choose the
    omitted blocks' insertion points; malformed, duplicate, out-of-range, or more
    incomplete output remains an error.
    """
    partial = _partial_reading_order(raw_text, len(blocks))
    if partial is None:
        return None

    order, missing_ids, has_duplicates = partial
    block_count = len(blocks)
    missing_count = len(missing_ids)
    allowed_missing = min(2, max(1, block_count // 10))
    if missing_count < 1 or missing_count > allowed_missing or has_duplicates:
        return None

    return _insert_missing_from_reference(
        order,
        missing_ids,
        xycut_plus_order(blocks),
    )


def supplement_incomplete_reading_order(
    raw_text: object,
    blocks: Sequence[Block],
) -> list[int] | None:
    """Use XY-Cut++ to insert omissions after OXR-RO exhausts its retries.

    At least half of all unique IDs must come from OXR-RO so this remains a
    supplementation step rather than a full-page replacement. Duplicate IDs
    are collapsed to their first occurrence; malformed or out-of-range output
    remains an error.
    """
    partial = _partial_reading_order(raw_text, len(blocks))
    if partial is None:
        return None

    order, missing_ids, has_duplicates = partial
    block_count = len(blocks)
    if not missing_ids and not has_duplicates:
        return None
    if block_count and len(order) * 2 < block_count:
        return None

    supplemented_order = _insert_missing_from_reference(
        order,
        missing_ids,
        xycut_plus_order(blocks),
    )
    if sorted(supplemented_order) != list(range(block_count)):
        raise ValueError("XY-Cut++ supplementation did not produce an exact permutation")
    return supplemented_order


async def _resolve_page_reading_order(
    blocks: Sequence[Block],
    image: np.ndarray,
    semaphore: asyncio.Semaphore,
    page_index: int,
    *,
    client: OXRModelClient,
    config: OXRModelConfig,
) -> list[int]:
    block_count = len(blocks)
    if block_count <= 1:
        return list(range(block_count))

    if image is None or image.ndim < 2:
        raise ValueError(
            f"Page {page_index} image must have height and width dimensions"
        )
    image_size = (int(image.shape[1]), int(image.shape[0]))
    objects = {
        "ref": [block.type.value for block in blocks],
        "bbox": [block.bbox for block in blocks],
        "bbox_type": "real",
    }
    prompt = build_reading_order_prompt(objects, image_size=image_size)
    max_tokens = min(config.max_tokens, max(16, 4 * block_count))

    for attempt in range(config.retry_times):
        response_received = False
        try:
            async with semaphore:
                raw_text = await client.request_reading_order(prompt, image, max_tokens)
            response_received = True
            try:
                return parse_reading_order(raw_text, block_count)
            except ValueError:
                repaired_order = repair_incomplete_reading_order(raw_text, blocks)
                if repaired_order is not None:
                    logger.warning(
                        "reading_order_repaired page=%d blocks=%d attempt=%d/%d %s",
                        page_index,
                        block_count,
                        attempt + 1,
                        config.retry_times,
                        _reading_order_diagnostics(raw_text, block_count),
                    )
                    return repaired_order
                if attempt == config.retry_times - 1:
                    supplemented_order = supplement_incomplete_reading_order(
                        raw_text,
                        blocks,
                    )
                    if supplemented_order is not None:
                        logger.warning(
                            "reading_order_supplemented page=%d blocks=%d attempt=%d/%d %s",
                            page_index,
                            block_count,
                            attempt + 1,
                            config.retry_times,
                            _reading_order_diagnostics(raw_text, block_count),
                        )
                        return supplemented_order
                    fallback_order = xycut_plus_order(blocks)
                    logger.warning(
                        "reading_order_xycut_fallback page=%d blocks=%d attempt=%d/%d %s",
                        page_index,
                        block_count,
                        attempt + 1,
                        config.retry_times,
                        _reading_order_diagnostics(raw_text, block_count),
                    )
                    return fallback_order
                raise
        except Exception as error:
            if response_received:
                logger.warning(
                    "reading_order_invalid page=%d blocks=%d attempt=%d/%d %s",
                    page_index,
                    block_count,
                    attempt + 1,
                    config.retry_times,
                    _reading_order_diagnostics(raw_text, block_count),
                )
            if attempt == config.retry_times - 1:
                message = format_error_message(error, f"Timeout after {config.timeout}s")
                raise PipelineError(
                    f"Reading order failed for page {page_index} after "
                    f"{config.retry_times} attempts: {message}",
                    status_code=get_error_status_code(error, default=503),
                ) from error
            await asyncio.sleep(1)

    raise AssertionError("unreachable")


async def resolve_reading_orders(
    pages_blocks: Sequence[Sequence[Block]],
    pages_images: Sequence[np.ndarray],
    semaphore: asyncio.Semaphore,
    *,
    client: OXRModelClient | None = None,
    config: OXRModelConfig | None = None,
) -> list[list[int]]:
    """Resolve each page's local reading-order permutation using original images."""
    if len(pages_blocks) != len(pages_images):
        raise ValueError("Page block and image counts must match")

    resolved_config = config or (
        client.config if client is not None else settings.oxr_model
    )
    owns_client = client is None and config is not None
    resolved_client = client or (
        OXRModelClient(resolved_config) if owns_client else model_client
    )
    try:
        return await gather_cancel_on_error(
            *(
                _resolve_page_reading_order(
                    blocks,
                    image,
                    semaphore,
                    page_index,
                    client=resolved_client,
                    config=resolved_config,
                )
                for page_index, (blocks, image) in enumerate(
                    zip(pages_blocks, pages_images)
                )
            )
        )
    finally:
        if owns_client:
            await resolved_client.aclose()
