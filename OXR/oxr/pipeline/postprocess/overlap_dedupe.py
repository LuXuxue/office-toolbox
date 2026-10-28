"""Page-level overlap cleanup after recognition and formula-number expansion.

Uses original layout evidence for Picture/Chemical parents and final reading
order/content for Text. Retained boxes and IDs are never resized or renumbered.
"""
from collections import Counter
from dataclasses import replace
import re

import cv2
import numpy as np

from oxr.pipeline.models import Block, BlockType


def _area(box: list[float]) -> float:
    return max(0.0, box[2] - box[0]) * max(0.0, box[3] - box[1])


def _intersection(a: list[float], b: list[float]) -> float:
    return _area([max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])])


def _inside(child: Block, parent: Block) -> bool:
    return _intersection(child.bbox, parent.bbox) / max(_area(child.bbox), 1e-9) >= 0.95


def _key(block: Block) -> tuple[BlockType, tuple[float, ...]]:
    return block.type, tuple(block.bbox)


def _normalized(content: str) -> str:
    return re.sub(r"\s+", "", content)


def dedupe_visual_parents(
    raw_blocks: list[Block], blocks: list[Block], image: np.ndarray,
    *, allow_mixed_children: bool = False,
) -> list[Block]:
    """Remove a uniquely mapped parent only when its stronger children cover ink.

    Captions count toward ink coverage and must still exist in the final page.
    Missing confidence, missing/ambiguous boxes, or nested deletion skip cleanup.
    Mixed visual groups also preserve titles and require at least two stronger
    children, while counting every retained child toward complete ink coverage.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image
    decisions = []
    for parent in raw_blocks:
        if parent.type not in {BlockType.PICTURE, BlockType.CHEMICAL_STRUCTURE} or parent.layout_score is None:
            continue
        children = [
            child for child in raw_blocks
            if child.idx != parent.idx
            and (child.type == parent.type or allow_mixed_children
                 and child.type in {BlockType.PICTURE, BlockType.CHEMICAL_STRUCTURE})
            and child.layout_score is not None
            and (allow_mixed_children or child.layout_score >= parent.layout_score + 0.1)
            and _area(child.bbox) < 0.65 * _area(parent.bbox)
            and _inside(child, parent)
        ]
        if len(children) < 2:
            continue
        if allow_mixed_children and sum(
            child.layout_score >= parent.layout_score + 0.1 for child in children
        ) < 2:
            continue
        if any(
            _intersection(a.bbox, b.bbox) / max(_area(a.bbox) + _area(b.bbox) - _intersection(a.bbox, b.bbox), 1e-9) > 0.1
            for i, a in enumerate(children) for b in children[i + 1:]
        ):
            continue
        x, y, u, v = map(round, parent.bbox)
        height, width = gray.shape
        x, y, u, v = max(0, x), max(0, y), min(width, u), min(height, v)
        if x >= u or y >= v:
            continue
        ink = gray[y:v, x:u] < 200
        covered = np.zeros(ink.shape, dtype=bool)
        label_types = {BlockType.CAPTION, BlockType.TITLE} if allow_mixed_children else {BlockType.CAPTION}
        captions = [b for b in raw_blocks if b.type in label_types and _inside(b, parent)]
        for block in children + captions:
            a, c, e, f = map(round, block.bbox)
            covered[max(0, c - y):min(v - y, f - y), max(0, a - x):min(u - x, e - x)] = True
        if not ink.any() or float((ink & covered).sum() / ink.sum()) < 0.999:
            continue
        decisions.append((parent, children, captions))

    raw_ids = Counter(block.idx for block in raw_blocks)
    final_keys = Counter(_key(block) for block in blocks)
    planned = {parent.idx for parent, _, _ in decisions}
    removed = set()
    for parent, children, captions in decisions:
        if any(raw_ids[b.idx] != 1 for b in [parent, *children]):
            continue
        if any(final_keys[_key(b)] != 1 for b in [parent, *children, *captions]):
            continue
        if any(child.idx in planned for child in children):
            continue
        removed.add(_key(parent))
    return [block for block in blocks if _key(block) not in removed]


def dedupe_complete_text(blocks: list[Block]) -> list[Block]:
    """Keep a large Text only if ordered small Texts reproduce all its characters.

    Positions, rather than native IDs, establish final reading order. Parent
    separators preserve sentence boundaries; existing line breaks become blank
    lines between child pieces only in plain text. Structured content retains
    its original separators; child-internal spacing never overwrites the parent.
    All other fields, including generated roles,
    survive unchanged.
    """
    text_positions = [i for i, b in enumerate(blocks) if b.type == BlockType.TEXT]
    removed = set()
    groups = []
    # Match the frozen selector: provisional removals reserve children even if
    # the later complete-cover guard rejects that parent.
    for parent_pos in sorted(text_positions, key=lambda i: (-_area(blocks[i].bbox), i)):
        if parent_pos in removed:
            continue
        parent = blocks[parent_pos]
        parent_text = _normalized(parent.content)
        eligible = []
        for child_pos in text_positions:
            if child_pos == parent_pos or child_pos in removed:
                continue
            child = blocks[child_pos]
            if _area(child.bbox) > 0.8 * _area(parent.bbox) or not _inside(child, parent):
                continue
            child_text = _normalized(child.content)
            if len(child_text) < 5 or child_text not in parent_text:
                continue
            eligible.append(child_pos)
        if not eligible:
            continue
        positions = sorted([parent_pos, *eligible])
        if len(positions) != positions[-1] - positions[0] + 1:
            continue
        cursor = 0
        for child_pos in eligible:
            child_text = _normalized(blocks[child_pos].content)
            position = parent_text.find(child_text, cursor)
            if position < 0:
                break
            cursor = position + len(child_text)
        else:
            removed.update(eligible)
            groups.append((parent_pos, eligible))

    actual_removed = set()
    replacements = {}
    for parent_pos, children in groups:
        raw = blocks[parent_pos].content
        if _normalized(raw) != "".join(_normalized(blocks[i].content) for i in children):
            continue
        char_positions = [i for i, char in enumerate(raw) if not char.isspace()]
        cursor = 0
        pieces = []
        # Conservatively leave structured text byte-for-byte intact. A child
        # boundary alone cannot prove that a newline is outside math, code,
        # links, emphasis, HTML, a list or an indented block.
        structured = bool(
            re.search(r"[$`\\<>\[\]*_~|]", raw)
            or re.search(r" {2,}\r?\n|(?m:^[ ]{0,3}(?:=+|-+)[ \t]*$)", raw)
            or re.search(r"(?m)^(?: {4}|\t| {0,3}(?:#{1,6}\s|(?:[-+*]|\d+[.)])\s))", raw)
        )
        for i, child_pos in enumerate(children):
            if i:
                separator = raw[char_positions[cursor - 1] + 1:char_positions[cursor]]
                if not structured and ("\n" in separator or "\r" in separator):
                    separator = "\n\n"
                pieces.append(separator)
            length = len(_normalized(blocks[child_pos].content))
            # The child proves coverage, but the selected parent's text is the
            # source of truth, including internal word and formula spacing.
            pieces.append(raw[char_positions[cursor]:char_positions[cursor + length - 1] + 1])
            cursor += length
        content = raw[:char_positions[0]] + "".join(pieces) + raw[char_positions[-1] + 1:]
        assert _normalized(content) == _normalized(raw)
        replacements[parent_pos] = replace(blocks[parent_pos], content=content)
        actual_removed.update(children)
    return [replacements.get(i, block) for i, block in enumerate(blocks) if i not in actual_removed]
