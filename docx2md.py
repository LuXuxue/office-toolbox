#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx2md — 将 Word (.docx) 文件转换为 Markdown 的单文件独立脚本

用法:
    python docx2md.py path/to/file.docx        转换单个文件
    python docx2md.py -d path/to/dir           转换目录下所有 .docx

转换结果写在与 docx 同一目录、文件名相同的 .md 文件。

支持: 多级标题、Word 自动编号、列表、表格（含合并单元格与多级表头）、
      超链接、加粗/斜体/删除线、文档属性与域、修订、OMML 公式（转 LaTeX）

依赖:
    pip install python-docx

"""

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.styles import BabelFish
from docx.table import Table
from docx.text.paragraph import Paragraph


# ══════════════════════════════════════════════════════════════════
# Block 数据结构
# ══════════════════════════════════════════════════════════════════

@dataclass
class Block:
    """内容块，按文档顺序排列"""
    type: str  # "paragraph" / "table"
    text: str = ""
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)
    numbered: bool = False  # 段落本身带编号（标题/列表），合并时不可与邻段粘连


# ══════════════════════════════════════════════════════════════════
# 后处理
# ══════════════════════════════════════════════════════════════════

def remove_cjk_spaces(text: str) -> str:
    """去除中文字符之间的空格（半角与全角）"""
    text = re.sub(r'(?<=[\u4e00-\u9fff]) +(?=[\u4e00-\u9fff])', '', text)
    text = re.sub(r'(?<=[\u4e00-\u9fff])\u3000+(?=[\u4e00-\u9fff])', '', text)
    return text


def merge_broken_paragraphs(blocks) -> list:
    """过滤无用行 + 合并断裂短行 + 去除中文间多余空格

    只合并「上一行以「1.」这类序号结尾 + 本行很短」的情况，即 Word 里被换行
    拆开的序号与标题；docx 段落不会被分页截断，不做跨页续行合并。
    """
    result = []
    for b in blocks:
        if b.type == "paragraph":
            text = b.text.strip()
            if re.match(r'^\d+$', text):
                continue
            if re.match(r'^[IVXLCDM]{1,6}$', text):
                continue
            if re.match(r'^[\.\s]{3,}$', text):
                continue

            b.text = cleaned = remove_cjk_spaces(text)

            # 标题/列表项是独立的结构行，不能与相邻段落粘连
            prev = result[-1] if result else None
            if (prev is not None and prev.type == "paragraph"
                    and not b.numbered and not prev.numbered
                    and len(cleaned) <= 6
                    and re.search(r'\d+\.\s*$', prev.text.strip())):
                prev.text += cleaned
                continue

            result.append(b)
        else:
            result.append(b)
    return result


_HEADING_NUM_RE = re.compile(r'^(\d+(?:\.\d+)+)\s+(\S.*?)\s*$')
_HEADING_SWAP_RE = re.compile(r'^(.+?)[ 　]*(\d+\.\d+(?:\.\d+)*)$')
_HEADING_BAD_TAIL_RE = re.compile(r'[。？！?!.;；]\s*$')
_HEADING_BAD_INNER_RE = re.compile(r'。')
_HEADING_MAX_TITLE_LEN = 60


def _heading_level(num: str) -> int:
    return num.count('.') + 1


def _looks_like_heading_title(title: str) -> bool:
    if not title:
        return False
    if len(title) > _HEADING_MAX_TITLE_LEN:
        return False
    if _HEADING_BAD_TAIL_RE.search(title):
        return False
    if _HEADING_BAD_INNER_RE.search(title):
        return False
    return True


def fix_headings(blocks) -> list:
    """修复章节标题格式并添加 Markdown 标题级别前缀"""
    for b in blocks:
        if b.type != "paragraph":
            continue
        text = b.text.strip()
        if not text:
            continue
        # 自动编号的列表项还原后也是「1. 标题」的样子，与手写章节标题无法区分；
        # 真正的 Word 标题已带 # 前缀，两条正则都匹配不到，无需再处理
        if b.numbered or text.startswith('#'):
            continue

        m = _HEADING_SWAP_RE.match(text)
        if m:
            title = m.group(1).strip()
            num = m.group(2).strip()
            text = f"{num} {title}"
            b.text = text

        m = _HEADING_NUM_RE.match(text)
        if m:
            num = m.group(1)
            title = m.group(2).strip()
            level = _heading_level(num)
            if 2 <= level <= 6 and _looks_like_heading_title(title):
                b.text = f"{'#' * level} {num} {title}"

    return blocks


def headers_match(a, b) -> bool:
    a_clean = [h.strip().lower() for h in a if h.strip()]
    b_clean = [h.strip().lower() for h in b if h.strip()]
    return a_clean == b_clean


def _header_similarity(a, b) -> float:
    def chars(headers):
        s = set()
        for h in headers:
            for c in h.strip().lower():
                if not c.isspace():
                    s.add(c)
        return s

    ca, cb = chars(a), chars(b)
    if not ca and not cb:
        return 1.0
    if not ca or not cb:
        return 0.0
    return len(ca & cb) / len(ca | cb)


def _looks_like_data_row(block: Block) -> bool:
    if not block.rows:
        return False
    first = block.rows[0]
    non_empty = [c.strip() for c in first if c.strip()]
    if not non_empty:
        return False
    for t in non_empty:
        if _CODE_PATTERN.match(t):
            return True
    return False


def merge_consecutive_tables(blocks) -> list:
    """合并紧挨着的两张表（Word 里被拆成两个 w:tbl 的同一张表）"""
    if len(blocks) < 2:
        return blocks

    result = []
    i = 0
    while i < len(blocks):
        current = blocks[i]
        if current.type != "table" or i + 1 >= len(blocks):
            result.append(current)
            i += 1
            continue

        nxt = blocks[i + 1]
        if nxt.type != "table":
            result.append(current)
            i += 1
            continue

        should_merge = False
        merged_rows = nxt.rows
        nxt_first_row = list(nxt.rows[0]) if nxt.rows else []

        if headers_match(current.headers, nxt.headers):
            should_merge = True
        elif _looks_like_data_row(nxt):
            should_merge = True
        elif _header_similarity(current.headers, nxt.headers) >= 0.8:
            if _looks_like_data_row(nxt) or not headers_match(nxt.headers,
                                                             nxt_first_row):
                should_merge = True
            if should_merge and nxt.rows and not _looks_like_data_row(nxt):
                if _header_similarity(nxt.headers, nxt_first_row) >= 0.8:
                    # 后一张表的首行其实是表头，合并时丢掉
                    merged_rows = nxt.rows[1:]

        if should_merge:
            merged = Block(
                type="table",
                headers=current.headers,
                rows=current.rows + merged_rows,
            )
            i += 2
            while (i < len(blocks) and blocks[i].type == "table"
                   and headers_match(current.headers, blocks[i].headers)):
                merged.rows += blocks[i].rows
                i += 1
            result.append(merged)
        else:
            result.append(current)
            i += 1
    return result


# ══════════════════════════════════════════════════════════════════
# 表格数据处理
# ══════════════════════════════════════════════════════════════════

_CODE_PATTERN = re.compile(r'[A-Za-z]{2}\d{3,}')
# 单元格内容形如 11km、100%、≥80、2.5：出现这类值就说明该行是数据行而不是表头
_VALUE_CELL_RE = re.compile(r'^[≈≥≤<>＜＞约近]?\s*\d[\d,. ]*\s*[^\d\s]{0,4}$')


def _cell_texts(row) -> list:
    """单元格文本规范化（换行折成空格、去空白）"""
    return [str(c or "").replace("\n", " ").strip() for c in row]


def _is_header_row(row) -> bool:
    """按内容猜测某行是不是表头（短文本、无数值）"""
    non_empty = [t for t in _cell_texts(row) if t]
    if not non_empty:
        return True
    for t in non_empty:
        if _CODE_PATTERN.match(t):
            return False
    if any(_VALUE_CELL_RE.match(t) for t in non_empty):
        return False
    if len(non_empty) >= 3:
        avg = sum(len(t) for t in non_empty) / len(non_empty)
        return avg <= 5
    return True


def _first_row_is_header(data) -> bool:
    """整张表都没识别出表头时的兜底判定

    表头文字可以很长（如「计算平均日污水量(m³/d)」），靠长度猜会漏判，
    这里改看结构：首行不含纯数字而下面含；若下面全是长文本，则比较长度。
    """
    if len(data) < 2:
        return False
    first = [t for t in _cell_texts(data[0]) if t]
    if not first or any(t.isdigit() for t in first):
        return False
    rest = [t for row in data[1:] for t in _cell_texts(row) if t]
    if any(t.isdigit() for t in rest):
        return True
    if not rest:
        return False
    avg_first = sum(len(t) for t in first) / len(first)
    avg_rest = sum(len(t) for t in rest) / len(rest)
    return avg_first * 1.5 < avg_rest


def _detect_header_count(data, header_flags=None) -> int:
    """推断表头行数

    优先用 Word 自己记录的「重复标题行」；没有标记才按内容猜，都判不出时兜底。
    """
    if header_flags and len(header_flags) == len(data):
        marked = 0
        for flag in header_flags:
            if not flag:
                break
            marked += 1
        if 0 < marked < len(data):
            return marked

    count = 0
    for row in data:
        if not _is_header_row(row):
            break
        count += 1
    if count:
        return count
    return 1 if _first_row_is_header(data) else 0


def process_table_data(data, header_flags=None):
    """处理表格数据：扁平化多级表头、压缩空列、清理换行"""
    if not data:
        return [], []

    data_start = _detect_header_count(data, header_flags)

    header_rows = data[:data_start]
    data_rows = data[data_start:]

    if not data_rows:
        col_count = max(len(r) for r in header_rows) if header_rows else 0
        col_has = [False] * col_count
        for row in header_rows:
            for ci in range(min(len(row), col_count)):
                if row[ci] and str(row[ci]).strip():
                    col_has[ci] = True
        keep = [i for i, h in enumerate(col_has) if h]
        trimmed = []
        for row in header_rows:
            trimmed.append([str(row[i] or "") if i < len(row) else "" for i in keep])

        def clean(s: str) -> str:
            return s.replace("\n", "")

        flat = []
        for ci in range(len(trimmed[0])):
            parts = [clean(trimmed[ri][ci]).strip() for ri in range(len(trimmed))]
            flat.append(" ".join(p for p in parts if p))
        return flat, []

    col_count = max(
        max((len(r) for r in data_rows), default=0),
        max((len(r) for r in header_rows), default=0),
    )
    col_has = [False] * col_count
    for row in list(header_rows) + list(data_rows):
        for ci in range(min(len(row), col_count)):
            if row[ci] and str(row[ci]).strip():
                col_has[ci] = True
    keep = [i for i in range(col_count) if col_has[i]]
    if not keep:
        return [], []

    trimmed_rows = []
    for row in data_rows:
        trimmed_rows.append([str(row[i] or "") if i < len(row) else "" for i in keep])

    trimmed_headers = []
    for row in header_rows:
        trimmed_headers.append([str(row[i] or "") if i < len(row) else "" for i in keep])

    def clean_cell(s: str) -> str:
        return remove_cjk_spaces(s.replace("\n", " "))

    flat_headers = []
    for ci in range(len(keep)):
        parts = []
        for ri in range(len(trimmed_headers)):
            val = clean_cell(trimmed_headers[ri][ci]).strip()
            # 纵向合并的单元格会在多行表头里重复同一文字，拼接时只保留一次
            if val and (not parts or parts[-1] != val):
                parts.append(val)
        flat_headers.append(" ".join(parts))

    rows = [[clean_cell(c) for c in row] for row in trimmed_rows]

    flat_headers = _align_headers_to_data(flat_headers, rows)

    final_headers = []
    final_rows = [[] for _ in range(len(rows))]
    for ci in range(len(flat_headers)):
        h = flat_headers[ci]
        if h.strip():
            final_headers.append(h)
            for ri in range(len(rows)):
                final_rows[ri].append(rows[ri][ci])
        else:
            has_data = any(rows[ri][ci].strip() for ri in range(len(rows)))
            if has_data:
                final_headers.append(h)
                for ri in range(len(rows)):
                    final_rows[ri].append(rows[ri][ci])

    return final_headers, final_rows


def _align_headers_to_data(headers, rows):
    """修正合并单元格造成的表头与数据列错位：有表头无数据的列并到相邻列"""
    if not headers or not rows:
        return headers
    n = len(headers)
    new_headers = list(headers)
    has_header = [bool(h.strip()) for h in headers]
    has_data = [
        any(ci < len(row) and str(row[ci]).strip() for row in rows)
        for ci in range(n)
    ]

    for ci in range(n):
        if has_header[ci] and not has_data[ci]:
            target = None
            if ci - 1 >= 0 and has_data[ci - 1] and not has_header[ci - 1]:
                target = ci - 1
            elif ci + 1 < n and has_data[ci + 1] and not has_header[ci + 1]:
                target = ci + 1
            if target is not None:
                moved = new_headers[ci].strip()
                existing = new_headers[target].strip()
                new_headers[target] = f"{moved} {existing}".strip() if existing else moved
                new_headers[ci] = ""
                has_header[ci] = False
                has_header[target] = True
    return new_headers


# ══════════════════════════════════════════════════════════════════
# Markdown 渲染
# ══════════════════════════════════════════════════════════════════

_BLANK_LINES_RE = re.compile(r"\n{3,}")
_DISPLAY_MATH_RE = re.compile(r"^\$\$[\s\S]+\$\$$")
_HEADING_LINE_RE = re.compile(r"^#{1,6}\s")


def render(blocks) -> str:
    """将 Block 对象列表渲染为 Markdown 字符串"""
    lines = []

    prev_was_table = False
    for block in blocks:
        if block.type == "paragraph":
            text = block.text.strip()
            if not text:
                continue
            if prev_was_table:
                lines.append("")
            if lines and lines[-1] != "" and _HEADING_LINE_RE.match(text):
                # 标题前留一个空行，正文段落不会被误读成标题的续行
                lines.append("")
            lines.append(text)
            if _DISPLAY_MATH_RE.match(text):
                # 块级公式后留空行，连续 $$ 块不会被合并成一个
                lines.append("")
            prev_was_table = False
        else:
            if lines and lines[-1] != "":
                lines.append("")
            lines.extend(_render_table(block))
            prev_was_table = True

    # 折叠多余的空行，保证段落之间最多一个空行
    return _BLANK_LINES_RE.sub("\n\n", "\n".join(lines).strip())


def _render_table(block: Block):
    """将 Table Block 渲染为 GFM 表格语法"""
    if not block.headers and not block.rows:
        return []

    def escape(s: str) -> str:
        return s.replace("|", "\\|").replace("\n", "")

    result = []

    header_line = "| " + " | ".join(escape(h) for h in block.headers) + " |"
    result.append(header_line)

    separator = "| " + " | ".join("---" for _ in block.headers) + " |"
    result.append(separator)

    for row in block.rows:
        padded = list(row) + [""] * (len(block.headers) - len(row))
        data_line = "| " + " | ".join(escape(c) for c in padded) + " |"
        result.append(data_line)

    return result


# ══════════════════════════════════════════════════════════════════
# 自动编号
#
# Word 的自动编号不写在正文里，而是由 numbering.xml 定义、渲染时才显示，
# 直接抽 w:t 只能拿到「概述」而丢掉序号「1」。这里重建编号的计算逻辑。
# ══════════════════════════════════════════════════════════════════

_ROMAN_UNITS = ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'),
                (90, 'xc'), (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'),
                (5, 'v'), (4, 'iv'), (1, 'i'))
_CN_DIGITS = '零一二三四五六七八九'
_CIRCLED_DIGITS = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
_LEGAL_DIGITS = '零壹贰叁肆伍陆柒捌玖'

_BULLET_FORMATS = ('bullet', 'none')


def _to_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    for value, sym in _ROMAN_UNITS:
        while n >= value:
            out.append(sym)
            n -= value
    return "".join(out)


def _to_letter(n: int) -> str:
    """序号转字母：1→A，26→Z，27→AA"""
    if n <= 0:
        return str(n)
    out = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord('A') + rem))
    return "".join(reversed(out))


def _to_chinese(n: int) -> str:
    """阿拉伯数字转中文数字（chineseCounting 规则：10 写作「十」，11 写作「十一」）"""
    if n <= 0:
        return str(n)
    if n < 10:
        return _CN_DIGITS[n]
    if n < 20:
        return '十' + (_CN_DIGITS[n - 10] if n > 10 else '')
    if n < 100:
        rem = n % 10
        return _CN_DIGITS[n // 10] + '十' + (_CN_DIGITS[rem] if rem else '')
    if n < 1000:
        return _CN_DIGITS[n // 100] + '百' + _cn_tail(n % 100, _to_chinese, '一')
    if n < 10000:
        return _CN_DIGITS[n // 1000] + '千' + _cn_tail(n % 1000, _to_chinese, '一', True)
    return str(n)


def _cn_tail(rem: int, render, lead: str = '', zero_lead: bool = False) -> str:
    """千位/百位之后的部分：被跳过的位补「零」，小写 10~19 补前导「一」（大写「壹拾」自带）"""
    if rem == 0:
        return ''
    if rem < 10:
        return '零' + render(rem)
    zero = '零' if zero_lead and rem < 100 else ''
    text = render(rem)
    if lead and 10 <= rem < 20:
        text = lead + text
    return zero + text


def _to_chinese_legal(n: int) -> str:
    """大写中文数字（chineseLegalSimplified）"""
    if n <= 0:
        return str(n)
    if n < 10:
        return _LEGAL_DIGITS[n]
    if n < 100:
        rem = n % 10
        return _LEGAL_DIGITS[n // 10] + '拾' + (_LEGAL_DIGITS[rem] if rem else '')
    if n < 1000:
        return _LEGAL_DIGITS[n // 100] + '佰' + _cn_tail(n % 100, _to_chinese_legal)
    if n < 10000:
        return _LEGAL_DIGITS[n // 1000] + '仟' + _cn_tail(n % 1000, _to_chinese_legal, '', True)
    return str(n)


def _format_num_value(value: int, fmt: str) -> str:
    """按 w:numFmt 渲染单个级别的编号值"""
    fmt = (fmt or 'decimal').strip()
    if fmt == 'decimal':
        return str(value)
    if fmt == 'decimalZero':
        return str(value).zfill(2)
    if fmt == 'upperRoman':
        return _to_roman(value).upper()
    if fmt == 'lowerRoman':
        return _to_roman(value).lower()
    if fmt == 'upperLetter':
        return _to_letter(value).upper()
    if fmt == 'lowerLetter':
        return _to_letter(value).lower()
    if fmt == 'decimalFullWidth':
        return "".join(chr(ord(c) + 0xFEE0) for c in str(value))
    if fmt == 'decimalHalfWidth':
        return str(value)
    if fmt == 'chineseCounting' or fmt == 'chineseCountingThousand':
        return _to_chinese(value)
    if fmt == 'chineseLegalSimplified':
        return _to_chinese_legal(value)
    if fmt == 'ideographDigital':
        return _to_chinese(value)
    if fmt == 'decimalEnclosedCircle':
        if 1 <= value <= len(_CIRCLED_DIGITS):
            return _CIRCLED_DIGITS[value - 1]
        return f"({value})"
    if fmt == 'decimalEnclosedParen':
        return f"({value})"
    if fmt == 'decimalEnclosedFullstop':
        return f"{value}."
    if fmt == 'hex':
        return f"{value:X}"
    return str(value)


class NumberingResolver:
    """按文档顺序重建 Word 自动编号（numPr + numbering.xml）"""

    def __init__(self, doc=None):
        self._doc = doc
        self._abstracts = {}       # abstractNumId -> {ilvl: lvl定义}
        self._abstract_links = {}  # abstractNumId -> {styleLink, numStyleLink}
        self._num_map = {}         # numId -> abstractNumId
        self._start_override = {}  # (numId, ilvl) -> 起始值
        self._lvl_override = {}    # (numId, ilvl) -> lvl定义（lvlOverride 内联定义）
        self._pstyle_map = {}      # 样式名/styleId -> (numId, ilvl)
        self._counters = {}        # numId -> {ilvl: 当前值}
        self._depth = {}           # numId -> 当前展开的层级深度
        self._last_level = {}      # numId -> 上一段落的 ilvl
        self._style_cache = {}     # styleId -> (numId, ilvl)
        self._levels_cache = {}    # numId -> 级别定义
        self._loaded = False

    # ── 加载 ────────────────────────────────────────────────────────
    def load(self) -> "NumberingResolver":
        if self._loaded:
            return self
        self._loaded = True
        if self._doc is None:
            return self
        numbering = None
        try:
            part = self._doc.part.numbering_part
            numbering = part.element if part is not None else None
        except Exception:
            numbering = None
        if numbering is None:
            return self

        for child in numbering:
            if child.tag == qn('w:abstractNum'):
                self._parse_abstract(child)
            elif child.tag == qn('w:num'):
                self._parse_num(child)

        # 「将级别链接到样式」：编号级别上的 w:pStyle
        num_ids = sorted(self._num_map,
                         key=lambda v: (_to_int(v) is None, _to_int(v) or 0))
        for num_id in num_ids:
            levels = self._levels(num_id)
            for ilvl in sorted(k for k in levels if isinstance(k, int)):
                style_ref = (levels[ilvl] or {}).get('pStyle')
                if not style_ref:
                    continue
                for key in self._pstyle_keys(style_ref):
                    self._pstyle_map.setdefault(key, (num_id, ilvl))
        return self

    def _pstyle_keys(self, style_ref):
        """w:pStyle 记录的是 styleId，这里同时收集样式名，便于两种方式匹配"""
        ref = (style_ref or '').strip()
        keys = {ref.lower()}
        st = self._style_elem(ref)
        if st is None:
            st = self._style_elem_by_name(ref)
        if st is not None:
            sid = st.get(qn('w:styleId'))
            if sid:
                keys.add(sid.lower())
            nm = st.find(qn('w:name'))
            if nm is not None and nm.get(qn('w:val')):
                keys.add(nm.get(qn('w:val')).strip().lower())
            try:
                keys.add(BabelFish.internal2ui(ref).strip().lower())
            except Exception:
                pass
        return {k for k in keys if k}

    def _parse_abstract(self, elem):
        aid = elem.get(qn('w:abstractNumId'))
        if aid is None:
            return
        levels = {}
        for lvl in elem.findall(qn('w:lvl')):
            ilvl = _to_int(lvl.get(qn('w:ilvl')))
            if ilvl is None:
                continue
            levels[ilvl] = _parse_lvl(lvl)
        self._abstracts[aid] = levels

        style_link = elem.find(qn('w:styleLink'))
        num_style_link = elem.find(qn('w:numStyleLink'))
        self._abstract_links[aid] = {
            'styleLink': style_link.get(qn('w:val')) if style_link is not None else None,
            'numStyleLink': (num_style_link.get(qn('w:val'))
                             if num_style_link is not None else None),
        }

    def _parse_num(self, elem):
        nid = elem.get(qn('w:numId'))
        aid = elem.find(qn('w:abstractNumId'))
        if nid is None or aid is None:
            return
        self._num_map[nid] = aid.get(qn('w:val'))
        for ov in elem.findall(qn('w:lvlOverride')):
            ilvl = _to_int(ov.get(qn('w:ilvl')))
            if ilvl is None:
                continue
            so = ov.find(qn('w:startOverride'))
            if so is not None:
                start = _to_int(so.get(qn('w:val')))
                if start is not None:
                    self._start_override[(nid, ilvl)] = start
            inline_lvl = ov.find(qn('w:lvl'))
            if inline_lvl is not None:
                self._lvl_override[(nid, ilvl)] = _parse_lvl(inline_lvl)

    # ── 样式 → 编号 ─────────────────────────────────────────────────
    def _styles_root(self):
        try:
            return self._doc.styles.element
        except Exception:
            return None

    def _style_elem_by_name(self, name):
        root = self._styles_root()
        if root is None or not name:
            return None
        target = (name or '').strip().lower()
        for st in root.findall(qn('w:style')):
            nm = st.find(qn('w:name'))
            if nm is None:
                continue
            if (nm.get(qn('w:val')) or '').strip().lower() == target:
                return st
        return None

    def _style_elem(self, style_id):
        root = self._styles_root()
        if root is None or not style_id:
            return None
        for st in root.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == style_id:
                return st
        return None

    def _style_numpr(self, style_id, depth=0):
        """沿 basedOn 链查找样式自身的编号设置"""
        if style_id is None or depth > 10:
            return None, None
        if style_id in self._style_cache:
            return self._style_cache[style_id]
        self._style_cache[style_id] = (None, None)  # 防止循环引用
        st = self._style_elem(style_id)
        num_id = ilvl = None
        if st is not None:
            ppr = st.find(qn('w:pPr'))
            numpr = ppr.find(qn('w:numPr')) if ppr is not None else None
            if numpr is not None:
                e = numpr.find(qn('w:numId'))
                if e is not None:
                    num_id = e.get(qn('w:val'))
                e = numpr.find(qn('w:ilvl'))
                if e is not None:
                    ilvl = e.get(qn('w:val'))
            if num_id is None:
                based = st.find(qn('w:basedOn'))
                if based is not None:
                    num_id, ilvl = self._style_numpr(based.get(qn('w:val')), depth + 1)
        result = (num_id, ilvl)
        self._style_cache[style_id] = result
        return result

    def _pstyle_num(self, style_id, style_name):
        """编号定义中通过 w:pStyle 关联到样式的级别（Word 的「将级别链接到样式」）"""
        for key in (style_id, style_name):
            if not key:
                continue
            found = self._pstyle_map.get(key.strip().lower())
            if found is not None:
                return found
        return None

    # ── 段落 → (numId, ilvl) ────────────────────────────────────────
    def _paragraph_numpr(self, para):
        num_id = ilvl = None
        ppr = para._element.find(qn('w:pPr'))
        numpr = ppr.find(qn('w:numPr')) if ppr is not None else None
        if numpr is not None:
            e = numpr.find(qn('w:numId'))
            if e is not None:
                num_id = e.get(qn('w:val'))
            e = numpr.find(qn('w:ilvl'))
            if e is not None:
                ilvl = e.get(qn('w:val'))

        style_name = None
        style_id = None
        try:
            style = para.style
        except Exception:
            style = None
        if style is not None:
            style_name = getattr(style, "name", None)
            style_id = getattr(style, "style_id", None)
            if style_id is None:
                el = getattr(style, "_element", None)
                style_id = el.get(qn('w:styleId')) if el is not None else None

        if num_id is None and style_id is not None:
            s_num, s_ilvl = self._style_numpr(style_id)
            num_id, ilvl = s_num, s_ilvl

        if num_id == '0':
            # 显式声明 numId=0 表示取消编号，不再回退到样式/pStyle 关联
            return None, None

        if not num_id:
            # 段落与样式都没直接声明时，尝试按 numbering 的 pStyle 关联反查
            linked = self._pstyle_num(style_id, style_name)
            if linked is None:
                return None, None
            num_id, ilvl = linked[0], linked[1]

        return num_id, (_to_int(ilvl) or 0 if ilvl is not None else 0)

    # ── 编号计算 ────────────────────────────────────────────────────
    def _levels(self, num_id, _seen=None):
        if num_id in self._levels_cache:
            return self._levels_cache[num_id]
        _seen = _seen if _seen is not None else set()
        aid = self._num_map.get(num_id)
        levels = self._abstracts.get(aid) if aid is not None else None
        links = self._abstract_links.get(aid, {}) if aid is not None else {}
        link_name = links.get('numStyleLink')
        if levels and link_name and link_name not in _seen:
            # 编号样式（w:numStyleLink）指向真正带 styleLink 的编号定义
            _seen.add(link_name)
            st = self._style_elem_by_name(link_name)
            if st is not None:
                link_num, _ = self._style_numpr(st.get(qn('w:styleId')))
                if link_num:
                    linked = self._levels(link_num, _seen)
                    if linked:
                        levels = linked
        self._levels_cache[num_id] = levels
        return levels

    def _lvl_def(self, num_id, levels, ilvl):
        override = self._lvl_override.get((num_id, ilvl))
        base = levels.get(ilvl) if levels else None
        if override and base:
            merged = dict(base)
            merged.update(override)
            return merged
        return override or base

    def _start_value(self, num_id, levels, ilvl):
        override = self._start_override.get((num_id, ilvl))
        if override is not None:
            return override
        lvl = self._lvl_def(num_id, levels, ilvl)
        if lvl:
            return lvl.get('start', 1) or 1
        return 1

    def _next_value(self, num_id, levels, ilvl):
        """推进计数器并返回该级别的当前值（同时保持各级别上下文）"""
        counters = self._counters.setdefault(num_id, {})
        depth = self._depth.get(num_id, 0)
        start = self._start_value(num_id, levels, ilvl)

        if ilvl > depth:
            for lv in range(depth, ilvl):
                if counters.get(lv) is None:
                    counters[lv] = self._start_value(num_id, levels, lv)
            depth = ilvl
        elif ilvl < depth:
            for lv in range(ilvl + 1, depth):
                counters.pop(lv, None)
            depth = ilvl

        lvl = self._lvl_def(num_id, levels, ilvl)
        restart = (lvl or {}).get('restart')
        last = self._last_level.get(num_id)
        if restart is not None and restart < ilvl and last is not None and last <= restart:
            counters.pop(ilvl, None)

        current = counters.get(ilvl)
        counters[ilvl] = start if current is None else current + 1
        self._depth[num_id] = ilvl + 1
        self._last_level[num_id] = ilvl
        return counters[ilvl]

    def _render_lvl_text(self, text, num_id, levels, ilvl):
        counters = self._counters.get(num_id, {})

        def repl(m):
            if m.group(0) == '%%':  # 转义后的字面量 %
                return '%'
            idx = int(m.group(1)) - 1
            if idx < 0:
                return m.group(0)
            value = counters.get(idx)
            if value is None:
                value = self._start_value(num_id, levels, idx)
            fmt = 'decimal'
            lvl = self._lvl_def(num_id, levels, idx)
            if lvl:
                fmt = lvl.get('fmt', 'decimal')
            return _format_num_value(value, fmt)

        return _LVL_PLACEHOLDER_RE.sub(repl, text)

    # ── 对外接口 ────────────────────────────────────────────────────
    def prefix_for(self, para) -> str:
        """返回可直接拼在段落文本前的编号前缀（含分隔符），无编号时为空串"""
        number, suff = self._resolve(para)
        if not number:
            return ""
        if suff == 'nothing' or number.endswith((' ', '\u3000')):
            return number
        # 中文标点后不再补空格，保持与 Word 中「1、标题」的排版一致
        if number[-1] in '、，。；：？！）】》':
            return number
        return f"{number} "

    def has_number(self, para) -> bool:
        """判断段落是否带可还原的自动编号（不推进计数器）"""
        num_id, ilvl = self._paragraph_numpr(para)
        if num_id is None:
            return False
        levels = self._levels(num_id)
        lvl = self._lvl_def(num_id, levels, ilvl) if levels else None
        if not lvl or not lvl.get('text'):
            return False
        return (lvl.get('fmt') or 'decimal') not in _BULLET_FORMATS

    def _resolve(self, para):
        """计算段落编号，返回 (编号文本, 分隔方式)"""
        num_id, ilvl = self._paragraph_numpr(para)
        if num_id is None:
            return "", 'tab'
        levels = self._levels(num_id)
        if not levels:
            return "", 'tab'
        lvl = self._lvl_def(num_id, levels, ilvl)
        if not lvl:
            return "", 'tab'
        if (lvl.get('fmt') or 'decimal') in _BULLET_FORMATS:
            # 项目符号类编号不还原为文本
            return "", lvl.get('suff', 'tab')
        self._next_value(num_id, levels, ilvl)
        text = lvl.get('text') or ''
        if not text:
            return "", lvl.get('suff', 'tab')
        return self._render_lvl_text(text, num_id, levels, ilvl), lvl.get('suff', 'tab')


_LVL_PLACEHOLDER_RE = re.compile(r'%%|%(\d)')


def _to_int(value):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _parse_lvl(lvl):
    """解析单个 w:lvl 定义"""

    def attr(tag, default=None):
        e = lvl.find(qn(tag))
        if e is None:
            return default
        val = e.get(qn('w:val'))
        return default if val is None else val

    return {
        'start': _to_int(attr('w:start', '1')) or 1,
        'fmt': attr('w:numFmt', 'decimal') or 'decimal',
        'text': attr('w:lvlText', '') or '',
        'restart': _to_int(attr('w:lvlRestart')),
        'pStyle': attr('w:pStyle'),
        'suff': attr('w:suff', 'tab') or 'tab',
    }


# ══════════════════════════════════════════════════════════════════
# 公式转换（OMML → LaTeX）
#
# Word 公式是 OMML 标记（m:oMath）而非文本，抽 w:t 只能拿到零散字符，
# 这里还原成 LaTeX：块级输出 $$…$$，行内输出 $…$。
# ══════════════════════════════════════════════════════════════════

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def _m(tag: str) -> str:
    """OMML 元素的 qname"""
    return f'{{{MATH_NS}}}{tag}'


_MATH_CHARS = {
    '∑': r'\sum', '∏': r'\prod', '∐': r'\coprod', '∫': r'\int',
    '∬': r'\iint', '∭': r'\iiint', '∮': r'\oint', '⋀': r'\bigwedge',
    '⋁': r'\bigvee', '⋂': r'\bigcap', '⋃': r'\bigcup', '⋮': r'\vdots',
    '⋱': r'\ddots', '⨀': r'\bigodot', '⨁': r'\bigoplus', '⨂': r'\bigotimes',
}
_MATH_ACCENTS = {
    '\u0302': r'\hat', '\u02c6': r'\hat', '\u0303': r'\tilde',
    '\u0304': r'\bar', '\u0305': r'\bar', '\u00af': r'\bar',
    '\u0306': r'\breve', '\u0307': r'\dot', '\u0308': r'\ddot',
    '\u030c': r'\check', '\u20d7': r'\vec', '\u2192': r'\vec',
    '\u2194': r'\overleftrightarrow',
}
_MATH_ESCAPES = {
    '\\': r'\backslash{}', '{': r'\{', '}': r'\}', '$': r'\$', '&': r'\&',
    '#': r'\#', '%': r'\%', '_': r'\_', '^': r'\^{}', '~': r'\sim',
}
_MATH_FUNC_RE = re.compile(r'^[A-Za-z][A-Za-z0-9]*$')
# 只认汉字/假名，全角标点（＝×≤…）另由 _MATH_SYMBOLS 处理
_CJK_RE = re.compile(r'[\u3005-\u30ff\u4e00-\u9fff]')
# 只承载内容的容器元素，渲染时直接递归子元素
_MATH_CONTAINERS = frozenset(
    _m(t) for t in ('e', 'num', 'den', 'sub', 'sup', 'lim', 'deg', 'fName', 'mr'))

# Word 里常见的全角/符号/希腊字母写法，转成 LaTeX 宏
_MATH_SYMBOLS = {
    '＝': '=', '＋': '+', '－': '-', '×': r'\times ', '÷': r'\div ',
    '＜': '<', '＞': '>', '≤': r'\leq ', '≥': r'\geq ',
    '≠': r'\neq ', '≈': r'\approx ', '（': '(', '）': ')', '％': r'\%',
    '√': r'\sqrt ', '∑': r'\sum ', '∫': r'\int ', 'π': r'\pi ',
    'μ': r'\mu ', 'μ': r'\mu ', 'α': r'\alpha ', 'β': r'\beta ',
    'γ': r'\gamma ', 'σ': r'\sigma ', 'θ': r'\theta ', 'λ': r'\lambda ',
    'Δ': r'\Delta ', 'ψ': r'\psi ', 'φ': r'\phi ', 'Φ': r'\Phi ',
    'ρ': r'\rho ', 'ε': r'\varepsilon ', 'ω': r'\omega ', 'η': r'\eta ',
    'ξ': r'\xi ', 'τ': r'\tau ', 'ν': r'\nu ', 'ζ': r'\zeta ',
    'κ': r'\kappa ', 'Ω': r'\Omega ', '°': r'^{\circ}',
}
# 全角 ASCII（！～）统一转半角
_FULLWIDTH_ASCII = {c: chr(c - 0xFEE0) for c in range(0xFF01, 0xFF5F)}
_PRIME_CHARS = ("'", '\u2032', '\u2019')


def _escape_tex_text(text: str) -> str:
    """文本模式（\\text{…}）下的转义：只处理 TeX 保留符号，不放宏"""
    return ''.join(_MATH_ESCAPES.get(c, c) for c in text)


def _escape_math(text: str) -> str:
    """数学模式下的转义（全角写法、希腊字母等转成 LaTeX 宏）"""
    text = text.translate(_FULLWIDTH_ASCII)
    out = []
    for c in text:
        if c in _MATH_SYMBOLS:
            out.append(_MATH_SYMBOLS[c])
        else:
            out.append(_MATH_ESCAPES.get(c, c))
    return ''.join(out)


def _math_text(text: str) -> str:
    """公式里的普通文本：含中文时用 \\text 包裹，避免被当成变量"""
    if not text:
        return ''
    text = text.translate(_FULLWIDTH_ASCII)
    if _CJK_RE.search(text):
        # \text 内只能放普通字符，放宏会让部分渲染器报错
        return r'\text{' + _escape_tex_text(text) + '}'
    return _escape_math(text)


def _as_script(value: str) -> str:
    """上下标内容：多字母英文（min/max 等）按正体，撇号用 \\prime"""
    if not value:
        return ''
    if value in _PRIME_CHARS:
        return r'\prime'
    if len(value) > 1 and _MATH_FUNC_RE.match(value):
        return r'\mathrm{' + value + '}'
    return value


def _math_children(elem, skip=()) -> str:
    parts = []
    for child in elem:
        if child.tag in skip:
            continue
        parts.append(_render_math(child))
    return ''.join(parts)


def _math_val(parent, tag, default=None):
    """读取 m:xxxPr 里的属性值（OMML 的属性在 m 命名空间，不是 w）"""
    pr = parent.find(_m(tag)) if parent is not None else None
    if pr is None:
        return default
    val = pr.get(_m('val'))
    if val is None:
        val = pr.get(qn('w:val'))
    return default if val is None else val


def _math_flag(parent, tag) -> bool:
    """读取 m:xxxHide 之类的开关（缺省即关闭）"""
    pr = parent.find(_m(tag)) if parent is not None else None
    if pr is None:
        return False
    val = pr.get(_m('val'))
    if val is None:
        val = pr.get(qn('w:val'))
    return val is None or str(val).lower() in ('1', 'on', 'true')


def _render_math_run(elem) -> str:
    """m:r → 公式里的文字"""
    texts = []
    for child in elem:
        if child.tag == _m('t'):
            texts.append(child.text or '')
        elif child.tag in (_m('br'), _m('tab'), qn('w:br'), qn('w:tab')):
            texts.append(' ')
    text = ''.join(texts)
    if not text:
        return ''

    rpr = elem.find(_m('rPr'))
    upright = False
    if rpr is not None:
        # sty="p"（普通体）/ scr="roman" 都是正体，nor 表示普通文本
        upright = (_math_val(rpr, 'sty') == 'p'
                   or _math_val(rpr, 'scr') == 'roman'
                   or rpr.find(_m('nor')) is not None)
    if upright and _MATH_FUNC_RE.match(text.strip()):
        return r'\mathrm{' + _escape_math(text.strip()) + '}'
    return _math_text(text)


def _render_math_delimiter(elem) -> str:
    """m:d → \\left( … \\right)"""
    pr = elem.find(_m('dPr'))
    beg, end, sep = '(', ')', '|'
    if pr is not None:
        beg = _math_val(pr, 'begChr', '(')
        end = _math_val(pr, 'endChr', ')')
        sep = _math_val(pr, 'sepChr', '|')
    parts = []
    for child in elem:
        if child.tag != _m('e'):
            continue
        parts.append(_render_math(child))
    inner = (f' \\mathrel{{{_escape_math(sep)}}} '.join(parts) if len(parts) > 1
             else ''.join(parts))
    left = r'\left' if beg else r'\left.'
    right = r'\right' if end else r'\right.'
    return f'{left}{_escape_math(beg)}{inner}{right}{_escape_math(end)}'


def _render_math_nary(elem) -> str:
    """m:nary → 求和/积分等 n 元运算符"""
    pr = elem.find(_m('naryPr'))
    chr_ = '∫'
    und_limits = False
    if pr is not None:
        chr_ = _math_val(pr, 'chr', '∫')
        und_limits = _math_val(pr, 'limLoc', 'subSup') == 'undOvr'
    op = _MATH_CHARS.get(chr_) or _escape_math(chr_)

    sub = sup = body = ''
    for child in elem:
        if child.tag == _m('sub'):
            sub = _render_math(child)
        elif child.tag == _m('sup'):
            sup = _render_math(child)
        elif child.tag == _m('e'):
            body = _render_math(child)
    if pr is not None:
        if _math_flag(pr, 'subHide'):
            sub = ''
        if _math_flag(pr, 'supHide'):
            sup = ''

    scripts = ''
    if sub:
        scripts += f'_{{{sub}}}'
    if sup:
        scripts += f'^{{{sup}}}'
    if und_limits:
        scripts = r'\limits' + scripts
    return f'{op}{scripts} {body}'


def _render_math_matrix(elem) -> str:
    """m:m（mr 里是 m:e 单元格）/ m:eqArr（每个 m:e 就是一行）"""
    rows = []
    aligned = False
    for child in elem:
        if child.tag == _m('mr'):
            rows.append(' & '.join(_render_math(c) for c in child
                                   if c.tag == _m('e')))
        elif child.tag == _m('e'):
            aligned = True
            rows.append(' & '.join(_render_math(c) for c in child))
    env = 'aligned' if aligned else 'matrix'
    body = ' \\\\ '.join(rows)
    return f'\\begin{{{env}}}{body}\\end{{{env}}}'


def _render_math(elem) -> str:
    """把一个 OMML 元素渲染为 LaTeX"""
    tag = elem.tag
    if tag == _m('t'):
        return _math_text(elem.text or '')
    if tag == _m('r'):
        return _render_math_run(elem)
    if tag == _m('f'):
        num = den = ''
        for child in elem:
            if child.tag == _m('num'):
                num = _render_math(child)
            elif child.tag == _m('den'):
                den = _render_math(child)
        pr = elem.find(_m('fPr'))
        if _math_val(pr, 'type') == 'noBar':
            return f'{num or "1"}/{den or "1"}'
        return f'\\frac{{{num or "1"}}}{{{den or "1"}}}'
    if tag in (_m('sSub'), _m('sSup'), _m('sSubSup'), _m('limLow'), _m('limUpp')):
        skip = {_m('sSubPr'), _m('sSupPr'), _m('sSubSupPr'),
                _m('limLowPr'), _m('limUppPr')}
        base = sub = sup = lim = ''
        for child in elem:
            if child.tag in skip:
                continue
            if child.tag == _m('e'):
                base = _render_math(child)
            elif child.tag == _m('sub'):
                sub = _render_math(child)
            elif child.tag == _m('sup'):
                sup = _render_math(child)
            elif child.tag == _m('lim'):
                lim = _render_math(child)
        low, high = _as_script(sub), _as_script(sup)
        if tag == _m('limLow'):
            low = _as_script(lim)
        elif tag == _m('limUpp'):
            high = _as_script(lim)
        out = base
        if low:
            out += f'_{{{low}}}'
        if high:
            out += f'^{{{high}}}'
        return out
    if tag == _m('rad'):
        pr = elem.find(_m('radPr'))
        deg = ''
        body = ''
        for child in elem:
            if child.tag == _m('deg'):
                deg = _render_math(child)
            elif child.tag == _m('e'):
                body = _render_math(child)
        if pr is not None and _math_flag(pr, 'degHide'):
            deg = ''
        return f'\\sqrt[{deg}]{{{body}}}' if deg else f'\\sqrt{{{body}}}'
    if tag == _m('d'):
        return _render_math_delimiter(elem)
    if tag == _m('nary'):
        return _render_math_nary(elem)
    if tag in (_m('m'), _m('eqArr')):
        return _render_math_matrix(elem)
    if tag == _m('func'):
        name = ''
        for child in elem:
            if child.tag == _m('fName'):
                raw = ''.join(t.text or '' for t in child.iter(_m('t'))).strip()
                name = (r'\mathrm{' + _escape_math(raw) + '}'
                        if _MATH_FUNC_RE.match(raw) else _math_text(raw))
        body = ''
        for child in elem:
            if child.tag == _m('e'):
                body = _render_math(child)
        return f'{name}{body}'
    if tag == _m('acc') or tag == _m('groupChr'):
        pr = elem.find(_m('accPr'))
        if pr is None:
            pr = elem.find(_m('groupChrPr'))
        chr_ = _math_val(pr, 'chr', '\u0302')
        if tag == _m('groupChr'):
            macro = {chr(0x23DF): r'\underbrace', chr(0x23DE): r'\overbrace'}.get(
                chr_, r'\underbrace')
        else:
            macro = _MATH_ACCENTS.get(chr_, r'\hat')
        body = low = high = ''
        for child in elem:
            if child.tag == _m('e'):
                body = _render_math(child)
            elif child.tag == _m('sub'):
                low = _render_math(child)
            elif child.tag == _m('sup'):
                high = _render_math(child)
        if tag == _m('groupChr') and low:
            return f'{macro}{{{body}}}_{{{low}}}'
        if high:
            return f'{macro}{{{body}}}^{{{high}}}'
        return f'{macro}{{{body}}}'
    if tag == _m('bar'):
        pr = elem.find(_m('barPr'))
        pos = _math_val(pr, 'pos', 'bot')
        body = ''
        for child in elem:
            if child.tag == _m('e'):
                body = _render_math(child)
        macro = r'\overline' if pos == 'top' else r'\underline'
        return f'{macro}{{{body}}}'
    if tag in (_m('box'), _m('border'), _m('phant')):
        # 边框/框注/幻影只起视觉作用，保留内容即可
        return _math_children(elem, skip=(tag + 'Pr',))
    if tag == _m('oMath'):
        return _math_children(elem)
    if tag == _m('oMathPara'):
        return _math_children(elem, skip=(_m('oMathParaPr'),))
    if tag in _MATH_CONTAINERS:
        # m:e / m:num / m:den / m:sub / m:sup / m:lim / m:deg 等纯容器
        return _math_children(elem)
    # m:ctrlPr / m:argPr / m:aln 等无内容元素
    return ''


def render_math(elem) -> str:
    """OMML 公式 → LaTeX 片段（$$…$$ 或 $…$）"""
    body = _render_math(elem).strip()
    if not body:
        return ''
    if elem.tag == _m('oMathPara'):
        return f'$${body}$$'
    return f'${body}$'


# ══════════════════════════════════════════════════════════════════
# Word 解析
# ══════════════════════════════════════════════════════════════════

HEADING_LEVEL_MAP = {
    "Heading 1": 1, "Heading 2": 2, "Heading 3": 3,
    "Heading 4": 4, "Heading 5": 5, "Heading 6": 6,
    "标题 1": 1, "标题 2": 2, "标题 3": 3,
    "标题 4": 4, "标题 5": 5, "标题 6": 6,
}

# Word 内置标题样式只到 9 级，Markdown 只支持 1-6 级，7 级以上统一压到 6 级
_HEADING_STYLE_RE = re.compile(r'^(?:heading|标题)\s*(\d{1,2})$', re.IGNORECASE)
_HEADING_MAX_LEVEL = 6


def heading_level_of(style_name: str) -> Optional[int]:
    """返回段落样式对应的 Markdown 标题级别（1-6），非标题返回 None"""
    if not style_name:
        return None
    level = HEADING_LEVEL_MAP.get(style_name)
    if level:
        return level
    m = _HEADING_STYLE_RE.match(style_name.strip())
    if m:
        return min(int(m.group(1)), _HEADING_MAX_LEVEL)
    return None


# 块级容器：Word 会把整段内容（封面、文档属性、目录等）包在 w:sdt 里
_BODY_CONTAINER_TAGS = frozenset((
    qn('w:sdt'), qn('w:customXml'), qn('w:ins'), qn('w:moveTo'),
))
_TOC_GALLERIES = ('table of contents', '目录')


def _is_toc_control(elm) -> bool:
    """内容控件是否为 Word 自动生成的目录（Table of Contents）"""
    if elm.tag != qn('w:sdt'):
        return False
    pr = elm.find(qn('w:sdtPr'))
    if pr is None:
        return False
    obj = pr.find(qn('w:docPartObj'))
    if obj is None:
        return False
    gallery = obj.find(qn('w:docPartGallery'))
    val = (gallery.get(qn('w:val')) or '') if gallery is not None else ''
    return val.strip().lower() in _TOC_GALLERIES


def _iter_body_blocks(container, host):
    """按文档顺序遍历正文块，穿透 w:sdt 等块级容器

    封面页、文档属性等内容常被包在 w:sdt 里，只遍历 w:body 的直接子元素会整块丢失。
    目录控件跳过：它是自动生成的索引，页码在 markdown 里没有意义。
    """
    for child in container:
        tag = child.tag
        if tag == qn('w:p'):
            yield Paragraph(child, host)
        elif tag == qn('w:tbl'):
            yield Table(child, host)
        elif tag in _BODY_CONTAINER_TAGS and not _is_toc_control(child):
            content = child.find(qn('w:sdtContent'))
            yield from _iter_body_blocks(content if content is not None else child, host)


def parse_docx(path: str) -> list:
    """解析 .docx 文件，返回按文档顺序排列的 Block 列表"""
    doc = Document(path)
    blocks = []

    numbering = NumberingResolver(doc).load()

    # Paragraph/Table 需要 body 容器对象来访问 part、样式等
    host = getattr(doc, '_body', None) or doc
    list_counter = 0

    for elm in _iter_body_blocks(doc.element.body, host):
        if isinstance(elm, Paragraph):
            para = elm
            text = _render_paragraph(para, list_counter, doc, numbering)
            if not text.strip():
                continue
            style_name = para.style.name if para.style else ""
            is_heading = heading_level_of(style_name) is not None
            is_numbered = _is_numbered_list(para)
            is_bullet = _is_bullet_list(para)

            if is_heading:
                # 标题不参与正文列表计数
                list_counter = 0
            elif is_numbered:
                list_counter += 1
            elif not is_bullet:
                list_counter = 0

            blocks.append(Block(
                type="paragraph", text=text,
                numbered=(is_heading or is_numbered or is_bullet
                          or bool(numbering and numbering.has_number(para))),
            ))
        else:
            data = _extract_table_grid(elm)
            headers, rows = process_table_data(data, _table_header_flags(elm))
            if headers and rows:
                blocks.append(Block(type="table", headers=headers, rows=rows))

    blocks = merge_consecutive_tables(blocks)
    blocks = merge_broken_paragraphs(blocks)
    blocks = fix_headings(blocks)
    return blocks


def _render_paragraph(para, list_counter: int, doc=None, numbering=None) -> str:
    """渲染段落为 Markdown 文本（含样式前缀、Word 自动编号与行内格式）"""
    style_name = para.style.name if para.style else ""

    inline = _render_inline(para, doc)
    # 自动编号（含「表%1.%2-%3」这类自定义序号）不在文本里，需按 numbering.xml 补回
    prefix = numbering.prefix_for(para) if numbering else ""

    level = heading_level_of(style_name)
    if level:
        # 标题不能跨行：分页符/分栏符/换行符一律压成空格
        title = _collapse_spaces(inline)
        if not title:
            return ""
        if prefix:
            title = prefix + title
        return f"{'#' * level} {title}"

    if _is_bullet_list(para):
        if prefix:
            return prefix + inline.lstrip()
        return f"- {inline}"
    if _is_numbered_list(para):
        if prefix:
            return prefix + inline.lstrip()
        return f"{list_counter + 1}. {inline}"

    if prefix:
        return prefix + inline.lstrip()
    return inline


def _collapse_spaces(text: str) -> str:
    """把换行、制表符与连续空白压成单个空格"""
    return re.sub(r'\s+', ' ', text).strip()


# w:sdt 是文档属性/封面域与修订插入所用的容器，w:fldSimple 是 { TITLE } 这类简单域，
# 漏掉它们整段文字都会丢失；w:del / w:moveFrom 故意排除，删掉的内容不该出现
_INLINE_CONTAINER_TAGS = frozenset((
    qn('w:sdt'), qn('w:fldSimple'), qn('w:ins'), qn('w:moveTo'),
    qn('w:smartTag'), qn('w:bdo'), qn('w:dir'),
))


def _render_inline(para, doc=None) -> str:
    """渲染段落的行内内容：run 的加粗/斜体/删除线 + 超链接 + 内容控件/域"""
    return _render_inline_children(para._element, doc)


def _render_inline_children(parent, doc=None) -> str:
    """递归渲染元素下的行内内容，容器元素（内容控件/域/修订插入）要展开"""
    parts = []
    for child in parent:
        tag = child.tag
        if tag == qn('w:r'):
            parts.append(_render_run(child))
        elif tag == qn('w:hyperlink'):
            parts.append(_render_hyperlink(child, doc))
        elif tag in _INLINE_CONTAINER_TAGS:
            content = child.find(qn('w:sdtContent'))
            parts.append(_render_inline_children(
                content if content is not None else child, doc))
        elif tag in (_m('oMath'), _m('oMathPara')):
            parts.append(render_math(child))
    return "".join(parts)


def _render_run(r_element) -> str:
    """渲染单个 run 元素，应用加粗/斜体/删除线"""
    texts = []
    for t in r_element.iter():
        if t.tag == qn('w:t'):
            texts.append(t.text or "")
        elif t.tag in (qn('w:br'), qn('w:cr')):
            texts.append(_render_break(t))
        elif t.tag == qn('w:tab'):
            texts.append("\t")
    text = "".join(texts)
    if not text:
        return ""

    rpr = r_element.find(qn('w:rPr'))
    bold = italic = strike = False
    if rpr is not None:
        if rpr.find(qn('w:b')) is not None:
            bold = True
        if rpr.find(qn('w:i')) is not None:
            italic = True
        if rpr.find(qn('w:strike')) is not None:
            strike = True

    if strike:
        text = f"~~{text}~~"
    if bold:
        text = f"**{text}**"
    if italic:
        text = f"*{text}*"
    return text


def _render_break(elem) -> str:
    """换行符的文本表示

    分页符与分栏符只是排版指令，转成换行会把标题从中间截断，故丢弃。
    """
    if elem.tag == qn('w:br') and elem.get(qn('w:type')) in ('page', 'column'):
        return ""
    return "\n"


def _render_hyperlink(hl_element, doc=None) -> str:
    """渲染超链接元素为 [text](url)"""
    texts = []
    for r in hl_element.iter(qn('w:r')):
        texts.append(_render_run(r))
    text = "".join(texts)

    rid = hl_element.get(qn('r:id'))
    url = ""
    if rid and doc is not None:
        try:
            rel = doc.part.rels.get(rid)
            if rel is not None:
                url = rel.target_ref or ""
        except Exception:
            url = ""
    if url:
        return f"[{text}]({url})"
    return text


def _is_bullet_list(para) -> bool:
    """判断是否为无序列表项"""
    style_name = para.style.name if para.style else ""
    if "List Bullet" in style_name or "项目符号" in style_name:
        return True
    ppr = para._element.find(qn('w:pPr'))
    if ppr is not None:
        numpr = ppr.find(qn('w:numPr'))
        if numpr is not None:
            numfmt = numpr.find(qn('w:numFmt'))
            return numfmt is not None and numfmt.get(qn('w:val')) == 'bullet'
    return False


def _is_numbered_list(para) -> bool:
    """判断是否为有序列表项"""
    style_name = para.style.name if para.style else ""
    if "List Number" in style_name or "编号" in style_name:
        return True
    ppr = para._element.find(qn('w:pPr'))
    if ppr is not None:
        numpr = ppr.find(qn('w:numPr'))
        if numpr is not None:
            numfmt = numpr.find(qn('w:numFmt'))
            if numfmt is None:
                return False
            return numfmt.get(qn('w:val')) in ('decimal', 'chineseCounting')
    return False


def _extract_table_grid(tbl) -> list:
    """提取表格为二维网格，正确处理合并单元格"""
    grid = []
    vmerge_pending = {}

    for tr in tbl._element.iter(qn('w:tr')):
        row = []
        col_idx = 0
        tc_list = list(tr.findall(qn('w:tc')))
        tc_iter = iter(tc_list)

        while col_idx < _table_grid_cols(tbl) or tc_iter:
            try:
                tc = next(tc_iter)
            except StopIteration:
                break

            tcpr = tc.find(qn('w:tcPr'))
            grid_span = 1
            vmerge_state = "none"
            if tcpr is not None:
                gs = tcpr.find(qn('w:gridSpan'))
                if gs is not None:
                    try:
                        grid_span = int(gs.get(qn('w:val'), '1'))
                    except ValueError:
                        grid_span = 1
                vm = tcpr.find(qn('w:vMerge'))
                if vm is not None:
                    val = vm.get(qn('w:val'))
                    vmerge_state = "restart" if val == "restart" else "continue"

            cell_text = _tc_text(tc)

            if vmerge_state == "restart":
                for i in range(grid_span):
                    vmerge_pending[col_idx + i] = cell_text if i == 0 else ""
                row.append(cell_text)
                for _ in range(grid_span - 1):
                    row.append("")
                col_idx += grid_span
            elif vmerge_state == "continue":
                for i in range(grid_span):
                    row.append(vmerge_pending.get(col_idx + i, ""))
                col_idx += grid_span
            else:
                for i in range(grid_span):
                    vmerge_pending.pop(col_idx + i, None)
                row.append(cell_text)
                for _ in range(grid_span - 1):
                    row.append("")
                col_idx += grid_span

        if row:
            grid.append(row)

    if not grid:
        return [[cell.text for cell in row.cells] for row in tbl.rows]
    return grid


def _table_grid_cols(tbl) -> int:
    """获取表格的网格列数（tblGrid 的 gridCol 数量）"""
    tbl_grid = tbl._element.find(qn('w:tblGrid'))
    if tbl_grid is not None:
        return len(list(tbl_grid.iter(qn('w:gridCol'))))
    first_row = tbl._element.find(qn('w:tr'))
    if first_row is not None:
        return len(list(first_row.iter(qn('w:tc'))))
    return 0


def _table_header_flags(tbl) -> list:
    """逐行读取 Word 的「重复标题行」标记

    比按文字猜可靠；行序与 _extract_table_grid 一致，长度对不上时调用方会退回猜测。
    """
    flags = []
    for tr in tbl._element.iter(qn('w:tr')):
        trpr = tr.find(qn('w:trPr'))
        flags.append(trpr is not None and trpr.find(qn('w:tblHeader')) is not None)
    return flags


def _tc_text(tc_element) -> str:
    """提取 tc 元素的纯文本（含公式的 LaTeX 形式）"""
    texts = []
    for t in tc_element.iter():
        if t.tag == qn('w:t'):
            texts.append(t.text or "")
        elif t.tag in (_m('oMath'), _m('oMathPara')):
            texts.append(render_math(t))
    return "".join(texts)


# ══════════════════════════════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════════════════════════════

def convert(docx_path: str) -> str:
    """将 docx 转换为 markdown 文本"""
    return render(parse_docx(docx_path))


def _convert_one(src: Path) -> bool:
    """将单个 docx 文件转换为同目录同名的 md 文件，成功返回 True"""
    try:
        content = convert(str(src))
    except Exception as e:
        print(f"转换失败: {src}: {e}", file=sys.stderr)
        return False

    out_path = src.with_suffix(".md")
    try:
        out_path.write_text(content, encoding="utf-8")
    except OSError as e:
        print(f"写入失败: {out_path}: {e}", file=sys.stderr)
        return False

    print(f"已生成: {out_path}")
    return True


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx2md",
        description="将 Word (.docx) 文件转换为 Markdown，输出同目录同名 .md 文件",
    )
    parser.add_argument("path", nargs="?", help="docx 文件路径（与 -d 二选一）")
    parser.add_argument("-d", "--dir", metavar="DIR", help="将目录下所有 docx 文件转换为 md 文件")
    args = parser.parse_args(argv)

    # 目录模式：转换目录下所有 docx
    if args.dir:
        dir_path = Path(args.dir)
        if not dir_path.is_dir():
            print(f"错误: 目录不存在: {dir_path}", file=sys.stderr)
            return 1
        docx_files = sorted(dir_path.glob("*.docx"))
        if not docx_files:
            print(f"错误: 目录中没有 .docx 文件: {dir_path}", file=sys.stderr)
            return 1
        success = 0
        for f in docx_files:
            if _convert_one(f):
                success += 1
        print(f"完成: 成功 {success} / 共 {len(docx_files)} 个文件")
        return 0 if success == len(docx_files) else 1

    # 单文件模式
    if not args.path:
        parser.error("必须提供文件路径或 -d/--dir 目录")
    src = Path(args.path)
    if not src.exists():
        print(f"错误: 文件不存在: {src}", file=sys.stderr)
        return 1
    if src.suffix.lower() != ".docx":
        print(f"错误: 仅支持 .docx 文件: {src}", file=sys.stderr)
        return 1

    return 0 if _convert_one(src) else 1


if __name__ == "__main__":
    sys.exit(main())
