import math
from collections.abc import Mapping, Sequence
from numbers import Real


PROMPT_TEXT_RECOGNITION = """Recognize the content of the image."""

PROMPT_FORMULA_RECOGNITION = """Recognize the formula of the image into KaTex format."""

PROMPT_TABLE_RECOGNITION = """Recognize the table of the image into html format."""

PROMPT_CODE_RECOGNITION = """Recognize the code of the image into markdown code format."""

PROMPT_LAYOUT_ANALYSIS = """Detect the layout of the input image, layout categories include [Caption,Footnote,Formula,Text,Page-footer,Page-header,Picture,Title,Table,Stamp,Chemical Structure,Code Block]. Box format is [x1,y1,x2,y2] and output box with reading order."""

PROMPT_WHOLE_PAGE_RECOGNITION = """Recognize the whole page content of the image."""

PROMPT_CHEMICAL_STRUCTURE_RECOGNITION = """Recognize the chemical structure of the image into SMILES."""

PROMPT_READING_ORDER = """
Given the detected layout boxes below, output the reading order as a whitespace-separated permutation of 0-based box IDs only.

Boxes:
{boxes}

Output only the box IDs."""


def _validate_coordinate(value: object, *, field: str) -> float:
    if isinstance(value, bool) or not isinstance(value, Real) or not math.isfinite(value):
        raise ValueError(f"{field} must contain only finite numbers")
    return float(value)


def _format_bbox(
    bbox: Sequence[object],
    *,
    bbox_type: str,
    image_size: tuple[int, int],
) -> str:
    if len(bbox) != 4:
        raise ValueError("each reading-order bbox must contain exactly four coordinates")

    width, height = image_size
    if width <= 0 or height <= 0:
        raise ValueError("image dimensions must be positive")

    if bbox_type == "real":
        scale_x, scale_y = 1000 / width, 1000 / height
    elif bbox_type == "norm1":
        scale_x = scale_y = 1000
    else:
        raise ValueError("objects.bbox_type must be 'real' or 'norm1'")

    coordinates = [
        int(round(_validate_coordinate(value, field="objects.bbox") * scale))
        for value, scale in zip(bbox, (scale_x, scale_y, scale_x, scale_y))
    ]
    return (
        f"<|box_start|>({coordinates[0]},{coordinates[1]}),"
        f"({coordinates[2]},{coordinates[3]})<|box_end|>"
    )


def build_reading_order_prompt(
    objects: Mapping[str, object],
    *,
    image_size: tuple[int, int],
) -> str:
    """Render ms-swift grounding objects into the reading-order prompt."""
    if not isinstance(objects, Mapping):
        raise ValueError("reading_order requires an objects JSON object")

    refs = objects.get("ref")
    bboxes = objects.get("bbox")
    if not isinstance(refs, Sequence) or isinstance(refs, (str, bytes)):
        raise ValueError("objects.ref must be a list")
    if not isinstance(bboxes, Sequence) or isinstance(bboxes, (str, bytes)):
        raise ValueError("objects.bbox must be a list")
    if len(refs) != len(bboxes):
        raise ValueError("objects.ref and objects.bbox must have the same length")

    bbox_type = objects.get("bbox_type") or "real"
    if not isinstance(bbox_type, str):
        raise ValueError("objects.bbox_type must be 'real' or 'norm1'")

    lines = []
    for index, (ref, bbox) in enumerate(zip(refs, bboxes)):
        if not isinstance(ref, str):
            raise ValueError("objects.ref must contain only strings")
        if not isinstance(bbox, Sequence) or isinstance(bbox, (str, bytes)):
            raise ValueError("objects.bbox must contain coordinate lists")
        lines.append(
            f"{index} | <|object_ref_start|>{ref}<|object_ref_end|> | "
            f"{_format_bbox(bbox, bbox_type=bbox_type, image_size=image_size)}"
        )

    return PROMPT_READING_ORDER.format(boxes="\n".join(lines))
