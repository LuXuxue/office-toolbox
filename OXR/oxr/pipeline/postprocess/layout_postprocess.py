import math
from typing import Any, Sequence

from oxr.pipeline.models import Block, BlockType, ContentFormat


LA_CLASS_ID_TO_BLOCK_TYPE: tuple[BlockType, ...] = (
    BlockType.CAPTION,
    BlockType.FOOTNOTE,
    BlockType.FORMULA,
    BlockType.TEXT,
    BlockType.PAGE_FOOTER,
    BlockType.PAGE_HEADER,
    BlockType.PICTURE,
    BlockType.TITLE,
    BlockType.TABLE,
    BlockType.CODE_BLOCK,
    BlockType.STAMP,
    BlockType.CHEMICAL_STRUCTURE,
)


def _response_error(message: str) -> ValueError:
    return ValueError(f"Invalid layout response: {message}")


def _require_list(value: Any, field: str) -> list[Any]:
    if not isinstance(value, list):
        raise _response_error(f"{field} must be a list")
    return value


def _require_number(value: Any, field: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise _response_error(f"{field} must be numeric")
    number = float(value)
    if not math.isfinite(number):
        raise _response_error(f"{field} must be finite")
    return number


def _require_image_size(value: Any, page_index: int) -> tuple[int, int]:
    if not isinstance(value, tuple) or len(value) != 2:
        raise _response_error(f"page {page_index} image dimensions must be a width-height tuple")
    width, height = value
    if (
        isinstance(width, bool)
        or isinstance(height, bool)
        or not isinstance(width, int)
        or not isinstance(height, int)
        or width <= 0
        or height <= 0
    ):
        raise _response_error(f"page {page_index} image dimensions must be positive integers")
    return width, height


def _intersection_over_union(left: list[float], right: list[float]) -> float:
    intersection_width = max(0.0, min(left[2], right[2]) - max(left[0], right[0]))
    intersection_height = max(0.0, min(left[3], right[3]) - max(left[1], right[1]))
    intersection = intersection_width * intersection_height
    if intersection == 0.0:
        return 0.0
    left_area = (left[2] - left[0]) * (left[3] - left[1])
    right_area = (right[2] - right[0]) * (right[3] - right[1])
    return intersection / (left_area + right_area - intersection)


def _classwise_nms(
    candidates: list[tuple[int, BlockType, list[float], float]],
    iou_threshold: float | None,
) -> list[tuple[int, BlockType, list[float], float]]:
    if iou_threshold is None:
        return candidates

    kept: list[tuple[int, BlockType, list[float], float]] = []
    for candidate in sorted(candidates, key=lambda item: (-item[3], item[0])):
        _, block_type, bbox, _ = candidate
        if any(
            kept_type == block_type
            and _intersection_over_union(bbox, kept_bbox) > iou_threshold
            for _, kept_type, kept_bbox, _ in kept
        ):
            continue
        kept.append(candidate)
    kept_positions = {position for position, _, _, _ in kept}
    return [candidate for candidate in candidates if candidate[0] in kept_positions]


def parse_layout_batch(
    payload: Any,
    image_sizes: Sequence[tuple[int, int]],
    score_threshold: float,
    nms_iou_threshold: float | None = 0.6,
) -> list[list[Block]]:
    """Validate a native LA batch response and convert retained candidates to blocks."""
    if not isinstance(payload, dict):
        raise _response_error("response must be an object")
    if payload.get("status") != "success":
        raise _response_error("status must equal 'success'")
    if not isinstance(payload.get("params"), dict):
        raise _response_error("params must be an object")

    data = payload.get("data")
    if not isinstance(data, dict):
        raise _response_error("data must be an object")
    labels = _require_list(data.get("labels"), "data.labels")
    boxes = _require_list(data.get("boxes"), "data.boxes")
    scores = _require_list(data.get("scores"), "data.scores")
    threshold = _require_number(score_threshold, "score threshold")
    if not 0.0 <= threshold <= 1.0:
        raise _response_error("score threshold must be between 0 and 1")
    nms_threshold = None
    if nms_iou_threshold is not None:
        nms_threshold = _require_number(nms_iou_threshold, "NMS IoU threshold")
        if not 0.0 <= nms_threshold <= 1.0:
            raise _response_error("NMS IoU threshold must be between 0 and 1")
    if not (len(labels) == len(boxes) == len(scores) == len(image_sizes)):
        raise _response_error("batch count must match image count")

    pages: list[list[Block]] = []
    for page_index, image_size in enumerate(image_sizes):
        width, height = _require_image_size(image_size, page_index)
        page_labels = _require_list(labels[page_index], f"page {page_index} labels")
        page_boxes = _require_list(boxes[page_index], f"page {page_index} boxes")
        page_scores = _require_list(scores[page_index], f"page {page_index} scores")
        if not (len(page_labels) == len(page_boxes) == len(page_scores)):
            raise _response_error(f"page {page_index} candidate count mismatch")

        candidates: list[tuple[int, BlockType, list[float], float]] = []
        for candidate_index, (label, raw_box, raw_score) in enumerate(
            zip(page_labels, page_boxes, page_scores)
        ):
            field = f"page {page_index} candidate {candidate_index}"
            if isinstance(label, bool) or not isinstance(label, int):
                raise _response_error(f"{field} label must be an integer")
            if not 0 <= label < len(LA_CLASS_ID_TO_BLOCK_TYPE):
                raise _response_error(f"{field} label must be between 0 and 11")
            score = _require_number(raw_score, f"{field} score")
            if not 0.0 <= score <= 1.0:
                raise _response_error(f"{field} score must be between 0 and 1")
            box = _require_list(raw_box, f"{field} bbox")
            if len(box) != 4:
                raise _response_error(f"{field} bbox must contain four coordinates")
            x1, y1, x2, y2 = [
                _require_number(value, f"{field} bbox[{coordinate_index}]")
                for coordinate_index, value in enumerate(box)
            ]
            if score < threshold:
                continue
            if x1 >= x2 or y1 >= y2:
                raise _response_error(f"{field} bbox must be nondegenerate")
            x1 = min(max(x1, 0.0), float(width))
            y1 = min(max(y1, 0.0), float(height))
            x2 = min(max(x2, 0.0), float(width))
            y2 = min(max(y2, 0.0), float(height))
            if x1 >= x2 or y1 >= y2:
                continue
            candidates.append(
                (
                    candidate_index,
                    LA_CLASS_ID_TO_BLOCK_TYPE[label],
                    [x1, y1, x2, y2],
                    score,
                )
            )

        page_blocks: list[Block] = []
        for _, block_type, bbox, score in _classwise_nms(candidates, nms_threshold):
            page_blocks.append(
                Block(
                    idx=len(page_blocks),
                    type=block_type,
                    bbox=bbox,
                    content="",
                    format=ContentFormat.TEXT,
                    layout_score=score,
                )
            )
        pages.append(page_blocks)
    return pages
