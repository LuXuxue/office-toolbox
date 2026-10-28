from enum import Enum
from dataclasses import dataclass

class BlockType(str, Enum):
    TEXT           = "Text"
    TITLE          = "Title"
    CAPTION        = "Caption"
    FOOTNOTE       = "Footnote"
    FORMULA        = "Formula"
    TABLE          = "Table"
    PICTURE        = "Picture"
    PAGE_HEADER    = "Page-header"
    PAGE_FOOTER    = "Page-footer"
    STAMP          = "Stamp"
    CHEMICAL_STRUCTURE = "Chemical Structure"
    CODE_BLOCK     = "Code Block"

class ContentFormat(str, Enum):
    TEXT     = "text"
    HTML     = "html"      # Table
    LATEX    = "latex"     # Formula
    MARKDOWN = "markdown"

@dataclass
class Block:
    idx:     int
    type:    BlockType
    bbox:    list[float]     # [x1, y1, x2, y2] pixel coordinates
    content: str             # recognition 结果
    format:  ContentFormat
    text_before: str = ""
    text_after: str = ""
    generated_role: str = ""
    layout_score: float | None = None
