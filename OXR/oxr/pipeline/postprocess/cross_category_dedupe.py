"""Remove systematic exact Text/Formula duplicates from one page."""

import re
from collections import defaultdict

from oxr.pipeline.models import Block, BlockType


_CROSS_CATEGORY_TYPES = frozenset({BlockType.TEXT, BlockType.FORMULA})
_OUTER_MATH_WRAPPERS = (("$$", "$$"), ("$", "$"), (r"\[", r"\]"), (r"\(", r"\)"))
_MIN_DUPLICATE_GROUPS = 2
_MIN_DUPLICATE_PERCENT = 10


def normalize_cross_category_content(content: str) -> str:
    """Normalize one outer math wrapper and whitespace."""
    normalized = content.strip()
    for opening, closing in _OUTER_MATH_WRAPPERS:
        if (
            normalized.startswith(opening)
            and normalized.endswith(closing)
            and len(normalized) >= len(opening) + len(closing)
        ):
            normalized = normalized[len(opening) : -len(closing)].strip()
            break
    return re.sub(r"\s+", "", normalized)


def _layout_score(block: Block) -> float:
    return block.layout_score if block.layout_score is not None else 0.0


def dedupe_systematic_cross_category_blocks(blocks: list[Block]) -> list[Block]:
    """Remove dense exact Text/Formula duplicates while preserving page order."""
    groups: dict[tuple[float, ...], list[int]] = defaultdict(list)
    for position, block in enumerate(blocks):
        if block.type in _CROSS_CATEGORY_TYPES:
            groups[tuple(block.bbox)].append(position)

    remove: set[int] = set()
    duplicate_group_count = 0
    for positions in groups.values():
        if {blocks[position].type for position in positions} != _CROSS_CATEGORY_TYPES:
            continue

        by_content: dict[str, list[int]] = defaultdict(list)
        for position in positions:
            normalized = normalize_cross_category_content(blocks[position].content)
            if normalized:
                by_content[normalized].append(position)

        for matching_positions in by_content.values():
            if len(matching_positions) < 2:
                continue
            if {
                blocks[position].type for position in matching_positions
            } != _CROSS_CATEGORY_TYPES:
                continue

            kept_position = max(
                matching_positions,
                key=lambda position: (_layout_score(blocks[position]), -position),
            )
            remove.update(
                position for position in matching_positions if position != kept_position
            )
            duplicate_group_count += 1

    content_count = sum(block.type in _CROSS_CATEGORY_TYPES for block in blocks)
    systematic = (
        duplicate_group_count >= _MIN_DUPLICATE_GROUPS
        and len(remove) * 100 >= content_count * _MIN_DUPLICATE_PERCENT
    )
    if not remove or not systematic:
        return blocks
    return [block for position, block in enumerate(blocks) if position not in remove]
