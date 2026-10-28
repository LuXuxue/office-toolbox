from typing import List
from oxr.pipeline.models import Block
from oxr.pipeline.postprocess.assembly_postprocess import blocks_to_markdown
from oxr.pipeline.postprocess.formula_postprocess import space_markdown_math

def assemble_markdown(pages_blocks: List[List[Block]]) -> str:
    """Assemble all recognized blocks across all pages into a final markdown document."""
    full_md = []
    
    for page_num, blocks in enumerate(pages_blocks):
        # We can add page separators or just concatenate.
        # For now, just concatenate
        page_md = blocks_to_markdown(blocks)
        full_md.append(page_md)
        
    return space_markdown_math("\n\n".join(full_md))
