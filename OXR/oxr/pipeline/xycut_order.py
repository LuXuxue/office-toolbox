"""XY-Cut++ adapter that returns a validated order without dropping blocks."""

from __future__ import annotations

import math
from typing import Sequence

from oxr.pipeline._vendor import xycut_plus_plus
from oxr.pipeline.models import Block, BlockType


_LABEL_MAP = {
    BlockType.TEXT: "text",
    BlockType.TITLE: "paragraph_title",
    BlockType.CAPTION: "figure_title",
    BlockType.FOOTNOTE: "footnote",
    BlockType.FORMULA: "formula",
    BlockType.TABLE: "table",
    BlockType.PICTURE: "image",
    BlockType.PAGE_HEADER: "text",
    BlockType.PAGE_FOOTER: "text",
    BlockType.STAMP: "seal",
    BlockType.CHEMICAL_STRUCTURE: "image",
    BlockType.CODE_BLOCK: "algorithm",
}

_NO_MASK_LABELS = [
    "text",
    "formula",
    "algorithm",
    "reference",
    "content",
    "abstract",
]


def _integer_bbox(block: Block) -> list[int]:
    if len(block.bbox) != 4 or not all(math.isfinite(value) for value in block.bbox):
        raise ValueError(f"invalid bbox for XY-Cut++ block {block.idx}: {block.bbox!r}")

    x1, y1, x2, y2 = block.bbox
    bbox = [math.floor(x1), math.floor(y1), math.ceil(x2), math.ceil(y2)]
    if bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
        raise ValueError(f"degenerate bbox for XY-Cut++ block {block.idx}: {block.bbox!r}")
    return bbox


def xycut_plus_order(blocks: Sequence[Block]) -> list[int]:
    """Return the pinned XY-Cut++ reference order for every input block.

    The adapter is order-only: overlap deletion is disabled, and the returned
    sequence must be an exact permutation before it can be used for repair.
    """
    if len(blocks) <= 1:
        return list(range(len(blocks)))

    entries = [
        {
            "block_bbox": _integer_bbox(block),
            "block_label": _LABEL_MAP[block.type],
            "block_content": str(index),
            "seg_start_flag": False,
            "seg_end_flag": False,
        }
        for index, block in enumerate(blocks)
    ]
    result = xycut_plus_plus.get_layout_ordering(
        entries,
        no_mask_labels=_NO_MASK_LABELS,
        remove_overlaps=False,
    )
    order = [int(entry["block_content"]) for entry in result]
    expected = list(range(len(blocks)))
    if len(order) != len(expected) or sorted(order) != expected:
        raise ValueError("XY-Cut++ did not return a complete unique permutation")
    return order
