"""Serialize detected regions without turning formula numbers into regions."""

import re

from oxr.pipeline.models import Block, BlockType
from oxr.pipeline.postprocess.formula_postprocess import FORMULA_NUMBERING_ROLE


def _numbered_content(content: str, before: str, after: str) -> str:
    if before:
        content += r"\leqno " + before
    if after:
        parenthesized = re.fullmatch(r"\((.*)\)", after, re.DOTALL)
        if parenthesized:
            content += r"\tag{" + parenthesized.group(1) + "}"
        else:
            content += r"\tag*{" + after + "}"
    return content


def serialize_final_blocks(
    blocks: list[Block], source_ranges: dict[int, tuple[int, int]] | None = None,
) -> tuple[list[dict], list[dict]]:
    """Fold surviving number witnesses into their formula; retain orphan text as notes.

    Number witnesses and their source have adjacent expansion indexes and the
    same box. Keep original indexes so footnote references and ordering remain
    stable. Optional demo ranges include attached numbering lines. This function
    never mutates the Markdown assembly input or its source ranges.
    """
    formulas = {block.idx: block for block in blocks if block.type == BlockType.FORMULA}
    numbering: dict[int, dict[str, str]] = {}
    numbering_indexes: dict[int, list[int]] = {}
    notes = []
    for block in blocks:
        if block.generated_role != FORMULA_NUMBERING_ROLE:
            continue
        owner = None
        for delta, field in ((1, "text_before"), (-1, "text_after")):
            formula = formulas.get(block.idx + delta)
            if (formula is not None and formula.bbox == block.bbox
                    and getattr(formula, field).strip() == block.content.strip()):
                owner = formula
                numbering.setdefault(owner.idx, {})[field] = block.content
                numbering_indexes.setdefault(owner.idx, []).append(block.idx)
                break
        if owner is None:
            # A formula removed by cleanup can leave protected numbering behind.
            # Preserve that text without inventing another detected region.
            notes.append({"idx": block.idx, "content": block.content})

    serialized = []
    for block in blocks:
        if block.generated_role == FORMULA_NUMBERING_ROLE:
            continue
        numbers = numbering.get(block.idx, {})
        content = _numbered_content(
            block.content, numbers.get("text_before", ""), numbers.get("text_after", ""),
        ) if block.type == BlockType.FORMULA else block.content
        region = {
            "idx": block.idx, "type": "Chemical Equation" if block.type == BlockType.FORMULA and r"\ce" in block.content else block.type.value, "bbox": list(block.bbox),
            "content": content, "format": block.format.value,
        }
        if source_ranges is not None:
            spans = [source_ranges[idx] for idx in (
                block.idx, *numbering_indexes.get(block.idx, []),
            ) if idx in source_ranges]
            region["start_line"] = min(span[0] for span in spans) if spans else None
            region["end_line"] = max(span[1] for span in spans) if spans else None
        serialized.append(region)
    return serialized, notes
