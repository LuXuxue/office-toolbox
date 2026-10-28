import re

from oxr.pipeline.models import Block, BlockType, ContentFormat

FORMULA_NUMBERING_ROLE = "formula_numbering"
_TEXT_COMMANDS = {
    "text",
    "textrm",
    "textnormal",
    "textbf",
    "textit",
    "mbox",
    "hbox",
    "verb",
}
_LATEX_COMMAND = re.compile(r"\\(?:[A-Za-z]+|.)", re.DOTALL)
_MARKDOWN_CODE = re.compile(r"(?P<ticks>`+)[\s\S]*?(?P=ticks)|~~~[\s\S]*?~~~")
_HTML_COMMENT = re.compile(r"<!--[\s\S]*?(?:-->|\Z)")
_HTML_TAG = re.compile(
    r"</?[A-Za-z][A-Za-z0-9-]*(?:\s[^>]*)?/?>"
    r"|<![A-Z][^>]*(?:>|\Z)|<\?[^>]*(?:>|\Z)"
)
_HTML_ENTITY_TAG = re.compile(
    r"&(?:lt|#0*60|#x0*3c);/?[A-Za-z][\s\S]*?"
    r"&(?:gt|#0*62|#x0*3e);",
    re.IGNORECASE,
)


def _protected_markup(text: str, index: int) -> re.Match[str] | None:
    return (
        _HTML_COMMENT.match(text, index)
        or _HTML_TAG.match(text, index)
        or _HTML_ENTITY_TAG.match(text, index)
    )


def _valid_math_open(text: str, index: int, width: int) -> bool:
    if width == 2:
        return not text.startswith("$$$", index)
    end = index + 1
    return end < len(text) and not text[end].isspace() and text[end] != "$"


def _valid_math_close(text: str, index: int, width: int) -> bool:
    if width == 2:
        return index > 0 and not text.startswith("$$$", index)
    return index > 0 and not text[index - 1].isspace() and not text.startswith("$$", index)


def space_latex_row_breaks(latex: str) -> str:
    """Separate LaTeX row breaks from an adjacent parenthesis or group brace."""
    result: list[str] = []
    index = 0
    while index < len(latex):
        if latex[index] == "%":
            end = latex.find("\n", index)
            end = len(latex) if end < 0 else end
            result.append(latex[index:end])
            index = end
            continue

        match = _LATEX_COMMAND.match(latex, index)
        if not match:
            result.append(latex[index])
            index += 1
            continue

        command = match.group()
        end = match.end()
        if command[1:] in _TEXT_COMMANDS:
            cursor = end
            while cursor < len(latex) and latex[cursor].isspace():
                cursor += 1
            if cursor < len(latex) and latex[cursor] == "{":
                depth = 1
                cursor += 1
                while cursor < len(latex) and depth:
                    escaped = _LATEX_COMMAND.match(latex, cursor)
                    if escaped:
                        cursor = escaped.end()
                        continue
                    if latex[cursor] == "{":
                        depth += 1
                    elif latex[cursor] == "}":
                        depth -= 1
                    cursor += 1
                if depth:
                    return latex
                end = cursor

        result.append(latex[index:end])
        if command == r"\\" and end < len(latex) and latex[end] in "(){}":
            result.append(" ")
        index = end

    return "".join(result)


def space_markdown_math(markdown: str) -> str:
    """Apply row-break spacing only inside complete dollar-delimited math."""
    result: list[str] = []
    index = 0
    while index < len(markdown):
        code = _MARKDOWN_CODE.match(markdown, index)
        if code:
            result.append(code.group())
            index = code.end()
            continue
        html = _protected_markup(markdown, index)
        if html:
            result.append(html.group())
            index = html.end()
            continue
        if markdown[index] != "$" or _is_escaped(markdown, index):
            result.append(markdown[index])
            index += 1
            continue

        width = 2 if markdown.startswith("$$", index) else 1
        if not _valid_math_open(markdown, index, width):
            result.append(markdown[index : index + width])
            index += width
            continue

        cursor = index + width
        end = None
        while cursor < len(markdown):
            if _MARKDOWN_CODE.match(markdown, cursor) or _protected_markup(markdown, cursor):
                break
            if markdown[cursor] == "$" and not _is_escaped(markdown, cursor):
                found_width = 2 if markdown.startswith("$$", cursor) else 1
                if (
                    found_width == width
                    and _valid_math_close(markdown, cursor, width)
                ):
                    end = cursor
                break
            cursor += 1

        if end is None:
            result.append(markdown[index : index + width])
            index += width
            continue

        delimiter = "$" * width
        result.append(
            delimiter
            + space_latex_row_breaks(markdown[index + width : end])
            + delimiter
        )
        index = end + width

    return "".join(result)


def _parse_braced_group(text: str, open_index: int) -> tuple[str, int] | None:
    if open_index >= len(text) or text[open_index] != "{":
        return None

    depth = 1
    index = open_index + 1
    while index < len(text):
        char = text[index]
        if char == "\\":
            index += 2
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[open_index + 1:index], index + 1
        index += 1

    return None


def _has_command_boundary(text: str, index: int) -> bool:
    return index >= len(text) or not text[index].isalpha()


def _numbering_command_at(text: str, index: int) -> str | None:
    if text.startswith(r"\tag*", index) and _has_command_boundary(text, index + len(r"\tag*")):
        return "tag*"
    if text.startswith(r"\tag", index) and _has_command_boundary(text, index + len(r"\tag")):
        return "tag"
    if text.startswith(r"\leqno", index) and _has_command_boundary(text, index + len(r"\leqno")):
        return "leqno"
    if text.startswith(r"\eqno", index) and _has_command_boundary(text, index + len(r"\eqno")):
        return "eqno"
    return None


def _find_next_numbering_command(text: str, start: int) -> int:
    index = start
    while index < len(text):
        if _numbering_command_at(text, index):
            return index
        index += 1
    return len(text)


def _append_numbering(target: list[str], value: str) -> None:
    value = value.strip()
    if value:
        target.append(value)


def _extract_formula_numbering(text: str) -> tuple[str, str, str]:
    formula_parts: list[str] = []
    before_numbers: list[str] = []
    after_numbers: list[str] = []
    index = 0

    while index < len(text):
        command = _numbering_command_at(text, index)
        if command in {"tag", "tag*"}:
            command_end = index + len(r"\tag*") if command == "tag*" else index + len(r"\tag")
            group_start = command_end
            while group_start < len(text) and text[group_start].isspace():
                group_start += 1

            parsed = _parse_braced_group(text, group_start)
            if parsed is None:
                formula_parts.append(text[index])
                index += 1
                continue

            number_text, index = parsed
            number_text = number_text.strip()
            if command == "tag":
                _append_numbering(after_numbers, f"({number_text})")
            else:
                _append_numbering(after_numbers, number_text)
            continue

        if command in {"eqno", "leqno"}:
            command_end = index + len(r"\leqno") if command == "leqno" else index + len(r"\eqno")
            number_start = command_end
            while number_start < len(text) and text[number_start].isspace():
                number_start += 1
            number_end = _find_next_numbering_command(text, number_start)
            number_text = text[number_start:number_end].strip()
            if command == "leqno":
                _append_numbering(before_numbers, number_text)
            else:
                _append_numbering(after_numbers, number_text)
            index = number_end
            continue

        formula_parts.append(text[index])
        index += 1

    return "".join(formula_parts).strip(), " ".join(before_numbers), " ".join(after_numbers)


def _is_escaped(text: str, index: int) -> bool:
    slash_count = 0
    cursor = index - 1
    while cursor >= 0 and text[cursor] == "\\":
        slash_count += 1
        cursor -= 1
    return slash_count % 2 == 1


def clean_formula_markdown_delimiters(latex: str) -> str:
    """Normalize clear wrapper/seam dollars in a Markdown-only formula copy.

    Leave mixed prose, incomplete math, text commands, escaped currency, and
    comments intact. Never modify recognized blocks or infer formula numbers.
    """
    dollars = []
    visible = []
    depth = 0
    index = 0
    while index < len(latex):
        if latex[index] == "%":
            end = latex.find("\n", index)
            index = len(latex) if end < 0 else end
            continue
        command = _LATEX_COMMAND.match(latex, index)
        if command:
            index = command.end()
            if command.group()[1:] in _TEXT_COMMANDS:
                if command.group()[1:] == "verb":
                    return latex
                cursor = index
                while cursor < len(latex) and latex[cursor].isspace():
                    cursor += 1
                group = _parse_braced_group(latex, cursor)
                if group is None:
                    return latex
                index = group[1]
            visible.append(" ")
            continue
        char = latex[index]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth < 0:
                return latex
        elif char == "$":
            if depth or latex.startswith("$$", index) or (index and latex[index - 1] == "$"):
                return latex
            dollars.append(index)
        visible.append(char)
        index += 1

    if depth or not dollars:
        return latex
    prose = "".join(visible)
    if "`" in prose or re.search(r"[\u3400-\u9fff]|\b[A-Za-z]{2,}\s+[A-Za-z]{2,}\b", prose):
        return latex
    wrapper = len(dollars) == 2 and dollars == [0, len(latex) - 1]
    seams = len(dollars) % 2 == 0 and all(
        latex[start + 1:end].isspace()
        for start, end in zip(dollars[::2], dollars[1::2])
    )
    if not wrapper and not seams:
        return latex
    if wrapper:
        return latex[1:-1]
    operators = ("=", "+", "-", "*", "/", "<", ">")
    # Only join explicit continuations; preserve independent or unclear rows.
    if not all(
        latex[:start].rstrip().endswith(operators)
        or latex[end + 1:].lstrip().startswith(operators)
        for start, end in zip(dollars[::2], dollars[1::2])
    ):
        return latex
    parts = []
    cursor = 0
    for start, end in zip(dollars[::2], dollars[1::2]):
        parts.extend((latex[cursor:start], latex[start + 1:end]))
        cursor = end + 1
    parts.append(latex[cursor:])
    return "".join(parts)


def _find_unescaped(text: str, marker: str, start: int) -> int:
    index = start
    while True:
        index = text.find(marker, index)
        if index == -1:
            return -1
        if not _is_escaped(text, index):
            return index
        index += len(marker)


def _format_numbered_math(
    open_marker: str,
    formula: str,
    close_marker: str,
    before: str,
    after: str,
) -> str:
    parts = []
    if before:
        parts.append(before)
    parts.append(f"{open_marker}{formula}{close_marker}")
    if after:
        parts.append(after)
    return " ".join(parts)


def _rewrite_numbered_math_range(
    open_marker: str,
    formula: str,
    close_marker: str,
) -> str:
    cleaned_formula, before, after = _extract_formula_numbering(formula)
    if not before and not after:
        return f"{open_marker}{formula}{close_marker}"
    return _format_numbered_math(open_marker, cleaned_formula, close_marker, before, after)


def move_formula_numbering_outside_math(text: str) -> str:
    result = []
    index = 0

    while index < len(text):
        if text.startswith(r"\[", index):
            close_index = text.find(r"\]", index + 2)
            if close_index != -1:
                formula = text[index + 2 : close_index]
                result.append(_rewrite_numbered_math_range(r"\[", formula, r"\]"))
                index = close_index + 2
                continue

        if text.startswith(r"\(", index):
            close_index = text.find(r"\)", index + 2)
            if close_index != -1:
                formula = text[index + 2 : close_index]
                result.append(_rewrite_numbered_math_range(r"\(", formula, r"\)"))
                index = close_index + 2
                continue

        if text.startswith("$$", index) and not _is_escaped(text, index):
            close_index = _find_unescaped(text, "$$", index + 2)
            if close_index != -1:
                formula = text[index + 2 : close_index]
                result.append(_rewrite_numbered_math_range("$$", formula, "$$"))
                index = close_index + 2
                continue

        if text[index] == "$" and not _is_escaped(text, index):
            close_index = _find_unescaped(text, "$", index + 1)
            if close_index != -1:
                formula = text[index + 1 : close_index]
                result.append(_rewrite_numbered_math_range("$", formula, "$"))
                index = close_index + 1
                continue

        result.append(text[index])
        index += 1

    return "".join(result)


def _numbering_text_block(source: Block, content: str) -> Block:
    return Block(
        idx=source.idx,
        type=BlockType.TEXT,
        bbox=list(source.bbox),
        content=content,
        format=ContentFormat.TEXT,
        generated_role=FORMULA_NUMBERING_ROLE,
    )


def expand_formula_numbering_blocks(blocks: list[Block]) -> list[Block]:
    expanded: list[Block] = []
    for block in blocks:
        if block.type != BlockType.FORMULA:
            expanded.append(block)
            continue

        text_before = getattr(block, "text_before", "").strip()
        text_after = getattr(block, "text_after", "").strip()
        if text_before:
            expanded.append(_numbering_text_block(block, text_before))
        expanded.append(block)
        if text_after:
            expanded.append(_numbering_text_block(block, text_after))

    for idx, block in enumerate(expanded):
        block.idx = idx
    return expanded


def postprocess_formula(block: Block, raw_text: str) -> Block:
    """Clean up formula block output to valid LaTeX."""
    text = raw_text.strip()
    match = re.search(r"```latex\n(.*?)\n```", text, re.DOTALL)
    if match:
        text = match.group(1)
    elif "```" in text:
        match = re.search(r"```(?:.*?)\n(.*?)\n```", text, re.DOTALL)
        if match:
            text = match.group(1)

    # Formula blocks store bare KaTeX/LaTeX; assembly owns display-math delimiters.
    wrappers = [
        r"^\$\$(.*?)\$\$$",
        r"^\$(.*?)\$$",
        r"^\\\[(.*?)\\\]$",
        r"^\\\((.*?)\\\)$",
    ]
    for pattern in wrappers:
        match = re.match(pattern, text, re.DOTALL)
        if match:
            text = match.group(1).strip()
            break

    text, text_before, text_after = _extract_formula_numbering(text)
    text = space_latex_row_breaks(text)

    block.content = text
    block.format = ContentFormat.LATEX
    block.text_before = text_before
    block.text_after = text_after
    block.generated_role = ""
    return block
