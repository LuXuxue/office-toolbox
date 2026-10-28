import re

from oxr.pipeline.models import Block, ContentFormat


def _strip_wrappers(text: str) -> str:
    text = text.strip()
    match = re.search(r"```(?:[^\n`]*)\n(.*?)\n```", text, re.DOTALL)
    if match:
        text = match.group(1).strip()

    math_wrappers = [
        (r"^\$\$(.*)\$\$$", re.DOTALL),
        (r"^\$(.*)\$$", re.DOTALL),
        (r"^\\\[(.*)\\\]$", re.DOTALL),
        (r"^\\\((.*)\\\)$", re.DOTALL),
    ]
    for pattern, flags in math_wrappers:
        match = re.match(pattern, text, flags)
        if match:
            text = match.group(1).strip()

    match = re.match(r"^\\smiles\{(.*)\}$", text, re.DOTALL)
    if match:
        text = match.group(1).strip()

    return text


def postprocess_chemical_structure(block: Block, raw_text: str) -> Block:
    """Convert chemical structure recognition output to SoMarkDown SMILES syntax."""
    smiles = _strip_wrappers(raw_text)
    block.content = f"$\\smiles{{{smiles}}}$" if smiles else ""
    block.format = ContentFormat.MARKDOWN
    return block
