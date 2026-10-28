import re
from oxr.pipeline.models import Block, ContentFormat
from oxr.pipeline.postprocess.formula_postprocess import move_formula_numbering_outside_math

def postprocess_text(block: Block, raw_text: str) -> Block:
    """Clean up text block output."""
    text = raw_text.strip()
    # Remove some excessive whitespaces
    text = re.sub(r'\n+', '\n', text)
    text = move_formula_numbering_outside_math(text)
    block.content = text
    block.format = ContentFormat.TEXT
    return block
