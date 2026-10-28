"""Conservative containment and duplicate-overlap cleanup on one page."""

import math
import re

import cv2
import numpy as np

from oxr.pipeline.models import Block, BlockType
from oxr.pipeline.postprocess.cross_category_dedupe import normalize_cross_category_content
from oxr.pipeline.postprocess.formula_postprocess import FORMULA_NUMBERING_ROLE


_MIN_IOU = 0.80
_MIN_SMALLER_COVERAGE = 0.95
_TEXT_HEIGHT_TOLERANCE = 0.20
_MIN_TEXT_CHILD_COVERAGE = 0.80
_TEXTUAL_TYPES = frozenset(BlockType) - {BlockType.PICTURE, BlockType.TABLE}
_PICTURE_TEXT_TYPES = _TEXTUAL_TYPES - {BlockType.CHEMICAL_STRUCTURE, BlockType.STAMP}


def _unprotected_textual_pair(a: Block, b: Block) -> bool:
    return (
        a.type in _TEXTUAL_TYPES and b.type in _TEXTUAL_TYPES
        and not {BlockType.CODE_BLOCK, BlockType.CHEMICAL_STRUCTURE} & {a.type, b.type}
        and FORMULA_NUMBERING_ROLE not in {a.generated_role, b.generated_role}
    )


def _formula_pair(a: Block, b: Block) -> bool:
    return _unprotected_textual_pair(a, b) and BlockType.FORMULA in {a.type, b.type}


def _valid_box(box: list[float]) -> bool:
    if not isinstance(box, (list, tuple)) or len(box) != 4:
        return False
    if any(
        not isinstance(value, (int, float))
        or isinstance(value, bool)
        or not math.isfinite(value)
        for value in box
    ):
        return False
    return box[0] < box[2] and box[1] < box[3]


def _contains(outer: list[float], inner: list[float]) -> bool:
    return (
        tuple(outer) != tuple(inner)
        and outer[0] <= inner[0] and outer[1] <= inner[1]
        and outer[2] >= inner[2] and outer[3] >= inner[3]
    )


def _overlaps(a: list[float], b: list[float]) -> bool:
    intersection = max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(
        0.0, min(a[3], b[3]) - max(a[1], b[1])
    )
    area_a = (a[2] - a[0]) * (a[3] - a[1])
    area_b = (b[2] - b[0]) * (b[3] - b[1])
    return (
        intersection / (area_a + area_b - intersection) >= _MIN_IOU
        and intersection / min(area_a, area_b) >= _MIN_SMALLER_COVERAGE
    )


def _content(block: Block) -> str:
    text = block.text_before + block.content + block.text_after
    if block.type == BlockType.CODE_BLOCK:
        return text
    text = "".join(text.split())
    if block.type not in {BlockType.FORMULA, BlockType.TEXT}:
        return text
    normalized = normalize_cross_category_content(text)
    # Only canonicalize the restricted atom/count/arrow syntax. Unknown LaTeX
    # stays literal, so formatting differences cannot erase distinct formulas.
    equation = re.fullmatch(r"\\ce\{([A-Za-z0-9()+\-= >]+)\}", normalized)
    if equation:
        chemical = equation[1]
    else:
        chemical = re.sub(r"\\mathrm\{([A-Za-z]+)\}", r"\1", normalized)
        chemical = re.sub(r"_\{([0-9]+)\}", r"\1", chemical)
        chemical = re.sub(r"\\(?:to|rightarrow)(?![A-Za-z])", "->", chemical)
    if "->" in chemical and re.fullmatch(r"[A-Za-z0-9()+\-= >]+", chemical):
        return chemical
    return normalized


def _preserves_content(winner: Block, loser: Block, image: np.ndarray | None) -> bool:
    if loser.type == BlockType.TABLE:
        # Geometry and a caption cannot prove that table cells/structure survive.
        return (
            winner.type == BlockType.TABLE
            and bool(loser.content.strip())
            and (winner.text_before, winner.content, winner.text_after)
            == (loser.text_before, loser.content, loser.text_after)
        )
    if loser.type == BlockType.PICTURE:
        if winner.type != BlockType.PICTURE:
            return False
        a, b = winner.bbox, loser.bbox
        if a[0] <= b[0] and a[1] <= b[1] and a[2] >= b[2] and a[3] >= b[3]:
            return True
        # A shifted crop can lose an edge even when its area overlaps almost all
        # of the other crop. Require source pixels to verify actual ink coverage.
        if image is None:
            return False
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image
        height, width = gray.shape
        x, y, u, v = map(round, b)
        x, y, u, v = max(0, x), max(0, y), min(width, u), min(height, v)
        if x >= u or y >= v:
            return False
        ink = gray[y:v, x:u] < 200
        covered = np.zeros(ink.shape, dtype=bool)
        left, top, right, bottom = map(round, a)
        left, top = max(x, left), max(y, top)
        right, bottom = min(u, right), min(v, bottom)
        if left < right and top < bottom:
            covered[top - y:bottom - y, left - x:right - x] = True
        return bool(ink.any()) and float((ink & covered).sum() / ink.sum()) >= 0.999
    if winner.type not in _TEXTUAL_TYPES:
        return False
    if loser.type == BlockType.CODE_BLOCK:
        return (
            winner.type == BlockType.CODE_BLOCK
            and bool(_content(loser)) and _content(loser) in _content(winner)
        )
    content = _content(loser)
    retained = _content(winner)
    return bool(content) and content in retained


def _absorbs(parent: Block, child: Block, image: np.ndarray | None) -> bool:
    if _contains(parent.bbox, child.bbox):
        return _unprotected_textual_pair(parent, child) or _preserves_content(parent, child, image)
    # Near-containment scales with the child text height. Ordinary text must
    # prove content inclusion; formula pairs retain the formula OCR exemption.
    if not _unprotected_textual_pair(parent, child):
        return False
    a, b = parent.bbox, child.bbox
    tolerance = (b[3] - b[1]) * _TEXT_HEIGHT_TOLERANCE
    intersection = max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(
        0.0, min(a[3], b[3]) - max(a[1], b[1])
    )
    child_area = (b[2] - b[0]) * (b[3] - b[1])
    return (
        (a[2] - a[0]) * (a[3] - a[1]) > child_area
        and intersection / child_area >= _MIN_TEXT_CHILD_COVERAGE
        and a[0] - tolerance <= b[0] and a[1] - tolerance <= b[1]
        and a[2] + tolerance >= b[2] and a[3] + tolerance >= b[3]
        and (_formula_pair(parent, child) or _preserves_content(parent, child, image))
    )


def _valid_score(score: float | None) -> bool:
    return (
        isinstance(score, (int, float))
        and not isinstance(score, bool)
        and math.isfinite(score)
    )


def dedupe_full_overlaps(blocks: list[Block], image: np.ndarray | None = None) -> list[Block]:
    """Keep containing boxes and suppress only near-identical intersections.

    Require full containment, or IoU >= 0.80 and smaller-box coverage >= 0.95.
    Complete textual containment skips content matching, except for chemical structures, code and
    generated equation numbers. Near-containment requires at least 80% child
    coverage and edge protrusion <= 20% of the child height. Ordinary text must
    preserve content; formula pairs skip content matching and use layout
    confidence for duplicate intersections. Tables require identical content;
    Pictures sharing exactly the same box with surviving recognized text yield
    to that text regardless of confidence. Other pictures require a containing
    picture crop or verified source ink coverage.
    Reconsider children whenever their
    proposed parent loses score suppression; only final survivors absorb them.
    Unknown scores do not participate in score suppression. Retained fields,
    object references and reading-order positions stay unchanged.
    """
    valid = [position for position, block in enumerate(blocks) if _valid_box(block.bbox)]
    remaining = set(valid)
    while remaining:
        children = {
            child for child in remaining
            if any(
                _absorbs(blocks[parent], blocks[child], image)
                for parent in remaining
            )
        }
        roots = remaining - children
        scored = sorted(
            (position for position in roots if _valid_score(blocks[position].layout_score)),
            key=lambda position: (-blocks[position].layout_score, position),
        )
        kept: list[int] = []
        losers = set()
        for position in scored:
            if any(
                _overlaps(blocks[winner].bbox, blocks[position].bbox)
                and (
                    _formula_pair(blocks[winner], blocks[position])
                    or _preserves_content(blocks[winner], blocks[position], image)
                )
                for winner in kept
            ):
                losers.add(position)
            else:
                kept.append(position)
        if not losers:
            remaining = roots
            break
        remaining -= losers

    # Use final survivors as witnesses so a suppressed text box cannot erase
    # its picture alias. Exact geometry is required; edge intersections stay.
    remaining -= {
        position for position in remaining
        if blocks[position].type == BlockType.PICTURE
        and any(
            blocks[text].type in _PICTURE_TEXT_TYPES
            and blocks[text].generated_role != FORMULA_NUMBERING_ROLE
            and bool(blocks[text].content.strip())
            and tuple(blocks[text].bbox) == tuple(blocks[position].bbox)
            for text in remaining
        )
    }
    removed = set(valid) - remaining

    if not removed:
        return blocks
    return [block for position, block in enumerate(blocks) if position not in removed]
