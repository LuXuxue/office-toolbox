from typing import List
from oxr.pipeline.models import Block, BlockType, ContentFormat
from oxr.pipeline.postprocess.formula_postprocess import clean_formula_markdown_delimiters

FORMULA_NUMBERING_ROLE = "formula_numbering"


def _is_formula_numbering(block: Block) -> bool:
    return getattr(block, "generated_role", "") == FORMULA_NUMBERING_ROLE


def _use_tight_separator(previous: Block, current: Block) -> bool:
    return (
        previous.type == BlockType.FORMULA and _is_formula_numbering(current)
    ) or (
        _is_formula_numbering(previous) and current.type == BlockType.FORMULA
    )


def _block_to_markdown_part(block: Block, content: str) -> str:
    if block.type == BlockType.TITLE:
        return f"# {content}\n"
    if block.type == BlockType.FORMULA:
        return f"$$\n{clean_formula_markdown_delimiters(content)}\n$$\n"
    if block.type in (
        BlockType.TABLE,
        BlockType.PICTURE,
        BlockType.CODE_BLOCK,
        BlockType.CHEMICAL_STRUCTURE,
    ):
        return f"{content}\n"
    if block.type in (BlockType.PAGE_HEADER, BlockType.PAGE_FOOTER):
        return f"*{content}*\n"
    if block.type == BlockType.CAPTION:
        return f"*{content}*\n"
    if block.type == BlockType.FOOTNOTE:
        return f"[^{block.idx}]: {content}\n"
    return f"{content}\n"


def blocks_to_markdown(blocks: List[Block]) -> str:
    """Convert a list of blocks to a single markdown string."""
    return blocks_to_markdown_with_ranges(blocks)[0]


def blocks_to_markdown_with_ranges(blocks: List[Block]) -> tuple[str, dict[int, tuple[int, int]]]:
    """Assemble once, retaining zero-based, end-exclusive source line ranges."""
    md_parts = []
    emitted_blocks = []
    ranges = {}
    line = 0
    
    # Sort blocks by index just in case
    sorted_blocks = sorted(blocks, key=lambda b: b.idx)
    
    for block in sorted_blocks:
        content = block.content.strip()
        if not content:
            continue
            
        part = _block_to_markdown_part(block, content)
        if md_parts and not _use_tight_separator(emitted_blocks[-1], block):
            md_parts.append("\n")
            line += 1
        md_parts.append(part)
        ranges[block.idx] = (line, line + part.count("\n"))
        line += part.count("\n")
        emitted_blocks.append(block)
            
    return "".join(md_parts), ranges
