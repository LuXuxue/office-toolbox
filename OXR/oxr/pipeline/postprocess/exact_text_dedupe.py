"""Remove exact cross-category duplicates of recognized textual blocks."""

import math
from collections import defaultdict

from oxr.pipeline.models import Block, BlockType


_TEXTUAL_TYPES = frozenset({BlockType.TEXT, BlockType.CAPTION, BlockType.TITLE})


def dedupe_exact_textual_blocks(blocks: list[Block]) -> list[Block]:
    """Keep the highest-confidence block without changing its fields or order.

    Require identical geometry, original text and assembly fields, one block
    per category, and finite confidence for every member of the group.
    """
    groups: dict[tuple[object, ...], list[int]] = defaultdict(list)
    for position, block in enumerate(blocks):
        if block.type not in _TEXTUAL_TYPES:
            continue
        if not isinstance(block.content, str) or not block.content.strip():
            continue
        box = block.bbox
        if not isinstance(box, (list, tuple)) or len(box) != 4:
            continue
        if any(not isinstance(value, (int, float)) or not math.isfinite(value) for value in box):
            continue
        if box[2] <= box[0] or box[3] <= box[1]:
            continue
        key = (
            tuple(box), block.content, block.format,
            block.text_before, block.text_after, block.generated_role,
        )
        groups[key].append(position)

    removed: set[int] = set()
    for positions in groups.values():
        categories = {blocks[position].type for position in positions}
        if len(categories) < 2 or len(categories) != len(positions):
            continue
        scores = [blocks[position].layout_score for position in positions]
        if any(
            not isinstance(score, (int, float))
            or isinstance(score, bool)
            or not math.isfinite(score)
            for score in scores
        ):
            continue
        winner = max(
            positions,
            key=lambda position: (blocks[position].layout_score, -position),
        )
        removed.update(position for position in positions if position != winner)

    if not removed:
        return blocks
    return [block for position, block in enumerate(blocks) if position not in removed]
