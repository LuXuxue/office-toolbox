import re

from oxr.pipeline.models import Block, ContentFormat


def postprocess_code(block: Block, raw_text: str) -> Block:
    """Clean code recognition output while preserving markdown code formatting."""
    text = raw_text.strip()
    match = re.search(r"```(?:[^\n`]*)\n(.*?)\n```", text, re.DOTALL)
    if match:
        text = match.group(1).strip()

    block.content = f"```\n{text}\n```" if text else ""
    block.format = ContentFormat.MARKDOWN
    return block
