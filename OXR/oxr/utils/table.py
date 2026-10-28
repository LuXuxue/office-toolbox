import re
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import List, Optional


@dataclass
class _TableCell:
    text: str
    rowspan: int = 1
    colspan: int = 1


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows: List[List[_TableCell]] = []
        self._table_depth = 0
        self._current_row: Optional[List[_TableCell]] = None
        self._current_cell: Optional[dict] = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        tag = tag.lower()
        if tag == "table":
            self._table_depth += 1
            return
        if self._table_depth == 0:
            return
        if tag == "tr":
            self._current_row = []
            return
        if tag in {"td", "th"} and self._current_row is not None and self._current_cell is None:
            attr_map = {key.lower(): value for key, value in attrs}
            self._current_cell = {
                "text": [],
                "rowspan": _positive_int(attr_map.get("rowspan")),
                "colspan": _positive_int(attr_map.get("colspan")),
            }
            return
        if tag == "br" and self._current_cell is not None:
            self._current_cell["text"].append(" ")
        if tag == "img" and self._current_cell is not None:
            src = dict(attrs).get("src")
            if src:
                self._current_cell["text"].append(f"![]({src})")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"} and self._current_cell is not None and self._current_row is not None:
            self._current_row.append(
                _TableCell(
                    text=_normalize_cell_text("".join(self._current_cell["text"])),
                    rowspan=self._current_cell["rowspan"],
                    colspan=self._current_cell["colspan"],
                )
            )
            self._current_cell = None
            return
        if tag == "tr" and self._current_row is not None:
            if self._current_row:
                self.rows.append(self._current_row)
            self._current_row = None
            return
        if tag == "table" and self._table_depth:
            self._table_depth -= 1

    def handle_data(self, data: str) -> None:
        if self._current_cell is not None:
            self._current_cell["text"].append(data)


def _positive_int(value: Optional[str]) -> int:
    try:
        parsed = int(value or "1")
    except ValueError:
        return 1
    return max(parsed, 1)


def _normalize_cell_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def strip_table_code_fence(text: str) -> str:
    text = text.strip()
    match = re.search(r"```html\n(.*?)\n```", text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    if "```" in text:
        match = re.search(r"```(?:.*?)\n(.*?)\n```", text, re.DOTALL)
        if match:
            return match.group(1).strip()
    return text


def _expand_spans(rows: List[List[_TableCell]]) -> List[List[str]]:
    expanded: List[List[str]] = []
    pending: dict[int, tuple[str, int]] = {}

    for source_row in rows:
        row: List[str] = []
        col_idx = 0

        def fill_pending_cell() -> bool:
            nonlocal col_idx
            if col_idx not in pending:
                return False
            text, remaining = pending[col_idx]
            row.append(text)
            if remaining > 1:
                pending[col_idx] = (text, remaining - 1)
            else:
                del pending[col_idx]
            col_idx += 1
            return True

        for cell in source_row:
            while fill_pending_cell():
                pass
            for _ in range(cell.colspan):
                row.append(cell.text)
                if cell.rowspan > 1:
                    pending[col_idx] = (cell.text, cell.rowspan - 1)
                col_idx += 1

        while pending and col_idx <= max(pending):
            if not fill_pending_cell():
                row.append("")
                col_idx += 1

        expanded.append(row)

    return expanded


def _escape_markdown_cell(text: str) -> str:
    return text.replace("\\", "\\\\").replace("|", "\\|")


def html_table_to_markdown(html: str) -> str:
    """Convert an HTML table to Markdown, expanding rowspans and colspans."""
    cleaned = strip_table_code_fence(html)
    parser = _TableParser()
    parser.feed(cleaned)

    rows = _expand_spans(parser.rows)
    if not rows:
        return _normalize_cell_text(re.sub(r"<[^>]+>", " ", cleaned))

    width = max(len(row) for row in rows)
    padded_rows = [row + [""] * (width - len(row)) for row in rows]
    escaped_rows = [
        [_escape_markdown_cell(_normalize_cell_text(cell)) for cell in row]
        for row in padded_rows
    ]

    header = escaped_rows[0]
    separator = ["---"] * width
    body = escaped_rows[1:]
    markdown_rows = [header, separator, *body]
    return "\n".join(f"| {' | '.join(row)} |" for row in markdown_rows)
