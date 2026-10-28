"""Associate table pictures, mask numbered TIT tokens and restore inline images."""

import base64
import math
import re
from html.parser import HTMLParser

import cv2
import numpy as np

from oxr.pipeline.models import Block, BlockType, ContentFormat
from oxr.utils.image import crop_image
from oxr.utils.table import strip_table_code_fence
from oxr.pipeline.image_assets import image_assets

TABLE_IMAGE_ROLE = "table_embedded_image"
_TOKEN = re.compile(r"(?:<tit>|&lt;tit(?:>|&gt;))\s*([0-9]*)", re.IGNORECASE)
_HTML_TAG = re.compile(
    r"</?[A-Za-z][A-Za-z0-9:_-]*(?:\s+(?:\"[^\"]*\"|'[^']*'|[^'\"<>])*)?\s*/?>|<!--.*?-->",
    re.DOTALL,
)
_PADDING = 4


def _valid_box(box: list[float]) -> bool:
    return (
        len(box) == 4
        and all(
            isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)
            for v in box
        )
        and box[0] < box[2]
        and box[1] < box[3]
    )


def group_table_images(
    blocks: list[Block],
    min_score: float,
    *,
    keep_outer_tables: bool = False,
) -> dict[int, list[Block]]:
    """Assign each confident Picture to a fully containing Table.

    Normally prefer the smallest table. Full overlap cleanup keeps outer tables,
    so in that mode integrate pictures in the outermost table instead of losing
    them when the inner table is suppressed.

    Keys are page-list positions, not block idx/read-order values. Existing block
    order and coordinates are never changed. No image classification is inferred
    from text or from page-specific coordinates.
    """
    tables = [
        (i, b)
        for i, b in enumerate(blocks)
        if b.type == BlockType.TABLE and _valid_box(b.bbox)
    ]
    groups: dict[int, list[Block]] = {}
    for picture in blocks:
        score = picture.layout_score
        if (
            picture.type != BlockType.PICTURE
            or not _valid_box(picture.bbox)
            or not isinstance(score, (int, float))
            or isinstance(score, bool)
            or not math.isfinite(score)
            or score < min_score
        ):
            continue
        x1, y1, x2, y2 = picture.bbox
        parents = [
            (i, table)
            for i, table in tables
            if table.bbox != picture.bbox
            and table.bbox[0] <= x1
            and table.bbox[1] <= y1
            and table.bbox[2] >= x2
            and table.bbox[3] >= y2
        ]
        if parents:
            i, _ = min(
                parents,
                key=lambda item: (
                    (-1 if keep_outer_tables else 1)
                    * (item[1].bbox[2] - item[1].bbox[0])
                    * (item[1].bbox[3] - item[1].bbox[1]),
                    item[0],
                ),
            )
            groups.setdefault(i, []).append(picture)
    for pictures in groups.values():
        pictures.sort(key=lambda b: (b.bbox[1], b.bbox[0], b.idx))
    return groups


def image_data_url(image: np.ndarray) -> str:
    assets = image_assets.get()
    if assets is not None:
        return assets.save(image)
    ok, encoded = cv2.imencode(".png", image)
    if not ok:
        raise ValueError("Unable to encode a table image.")
    return "data:image/png;base64," + base64.b64encode(encoded).decode("ascii")


def prepare_table_images(
    table: Block,
    pictures: list[Block],
    page_image: np.ndarray,
) -> tuple[np.ndarray, dict[int, str]]:
    """Keep original picture pixels; send blue/white numbered tokens to TBR."""
    masked = crop_image(page_image, table.bbox)
    h, w = masked.shape[:2]
    page_h, page_w = page_image.shape[:2]
    left = max(0, min(page_w, round(table.bbox[0])))
    top = max(0, min(page_h, round(table.bbox[1])))
    sources: dict[int, str] = {}
    regions: list[tuple[int, int, int, int]] = []
    for i, picture in enumerate(pictures, 1):
        a, b, c, d = (round(v) for v in picture.bbox)
        x1, y1 = max(0, a - left - _PADDING), max(0, b - top - _PADDING)
        x2, y2 = min(w, c - left + _PADDING), min(h, d - top + _PADDING)
        if x1 >= x2 or y1 >= y2:
            raise ValueError("Empty table picture crop.")
        if any(
            max(x1, a) < min(x2, c) and max(y1, b) < min(y2, d)
            for a, b, c, d in regions
        ):
            raise ValueError(
                "Overlapping table pictures cannot receive unambiguous tokens."
            )
        regions.append((x1, y1, x2, y2))
        sources[i] = image_data_url(
            page_image[top + y1 : top + y2, left + x1 : left + x2]
        )
        symbol = f"<tit>{i}"
        scale, thickness = 0.65, 1
        (tw, th), baseline = cv2.getTextSize(
            symbol, cv2.FONT_HERSHEY_SIMPLEX, scale, thickness
        )
        if tw + 4 > x2 - x1 or th + baseline + 4 > y2 - y1:
            raise ValueError(
                "Table picture is too small for a readable numbered token."
            )
        cv2.rectangle(masked, (x1, y1), (x2 - 1, y2 - 1), (255, 0, 0), -1)
        cv2.putText(
            masked,
            symbol,
            (x1 + (x2 - x1 - tw) // 2, y1 + (y2 - y1 + th) // 2),
            cv2.FONT_HERSHEY_SIMPLEX,
            scale,
            (255, 255, 255),
            thickness,
            cv2.LINE_AA,
        )
    return masked, sources


def _escape_stray_less_than(text: str) -> str:
    """Keep HTML tags intact while escaping OCR's bare math comparisons."""
    parts = []
    start = 0
    while (position := text.find("<", start)) != -1:
        parts.append(text[start:position])
        tag = _HTML_TAG.match(text, position)
        if tag:
            parts.append(tag.group())
            start = tag.end()
        else:
            parts.append("&lt;")
            start = position + 1
    parts.append(text[start:])
    return "".join(parts)


class _CellTokenValidator(HTMLParser):
    """Require every token to occur in cell text in a complete HTML table."""

    def __init__(self, markers: list[str]):
        super().__init__(convert_charrefs=True)
        self.markers = markers
        self.seen: list[str] = []
        self.table = False
        self.row = False
        self.cell = False
        self.complete = False
        self.valid = True

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            if self.table:
                self.valid = False
            self.table = True
        elif tag == "tr":
            if not self.table or self.row:
                self.valid = False
            self.row = True
        elif tag in {"td", "th"}:
            if not self.row or self.cell:
                self.valid = False
            self.cell = True

    def handle_endtag(self, tag):
        if tag in {"td", "th"}:
            if not self.cell:
                self.valid = False
            self.cell = False
        elif tag == "tr":
            if not self.row or self.cell:
                self.valid = False
            self.row = False
        elif tag == "table":
            if not self.table or self.row or self.cell:
                self.valid = False
            self.table = False
            self.complete = True

    def handle_data(self, data):
        for marker in self.markers:
            if marker in data:
                if not self.table or not self.cell:
                    self.valid = False
                self.seen.extend([marker] * data.count(marker))


def restore_table_images(raw_text: str, sources: dict[int, str]) -> str | None:
    """Restore only a complete unique ID set; never guess by token order."""
    text = strip_table_code_fence(raw_text)
    matches = list(_TOKEN.finditer(text))
    if len(matches) != len(sources) or not sources:
        return None
    numbers = [m.group(1) for m in matches]
    if any(not n or len(n) > len(str(len(sources))) for n in numbers):
        return None
    ids = [int(n) for n in numbers]
    if sorted(ids) != sorted(sources):
        return None
    prefix = "__OXR_TABLE_IMAGE_"
    while prefix in text:
        prefix = "_" + prefix
    markers = [f"{prefix}{i}__" for i in ids]
    marker_iter = iter(markers)
    probe = _escape_stray_less_than(_TOKEN.sub(lambda _: next(marker_iter), text))
    validator = _CellTokenValidator(markers)
    validator.feed(probe)
    validator.close()
    if not (
        validator.valid
        and validator.complete
        and not validator.table
        and not validator.row
        and not validator.cell
        and sorted(validator.seen) == sorted(markers)
    ):
        return None
    for marker, i in zip(markers, ids):
        probe = probe.replace(
            marker,
            f'<img data-table-image-id="{i}" alt="Table image {i}" src="{sources[i]}">',
        )
    return probe


def table_image_fallback(table: Block, page_image: np.ndarray) -> Block:
    """Retain the entire original table, including all pictures, on failure."""
    table.content = (
        f'<img alt="Table" src="{image_data_url(crop_image(page_image, table.bbox))}">'
    )
    table.format = ContentFormat.HTML
    return table


def mark_table_images(pictures: list[Block]) -> None:
    for picture in pictures:
        picture.generated_role = TABLE_IMAGE_ROLE


async def unchanged_block(block: Block) -> Block:
    """Keep RO's original index space while the parent recognizes its pictures."""
    return block
