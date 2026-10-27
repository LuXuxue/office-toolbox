#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx2md — 将 PowerPoint (.pptx) 文件转换为 Markdown 的单文件独立脚本

用法:
    python pptx2md.py path/to/file.pptx       转换单个文件
    python pptx2md.py -d path/to/dir          转换目录下所有 .pptx

转换结果写在与 pptx 同一目录、文件名相同的 .md 文件。
图片导出到 `<md名>_media/` 子目录，并在原位置写 `![说明](相对路径)`。

支持: 幻灯片标题（按节标题自动分层）、正文与多级列表、表格（含合并单元格
      与多级表头）、图片、图表数据、组合图形、嵌入对象的预览图、演讲者备注。

依赖:
    pip install python-pptx
    （Pillow 由 python-pptx 一并安装；缺了它只是不转换 EMF/WMF 图片）

"""

import argparse
import hashlib
import io
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

try:
    from PIL import Image
except ImportError:  # python-pptx 正常会带上 Pillow；缺失时保留原图
    Image = None


# ══════════════════════════════════════════════════════════════════
# 基础工具
# ══════════════════════════════════════════════════════════════════

# 表格里出现这类值就说明该行是数据行而不是表头
_CODE_PATTERN = re.compile(r'[A-Za-z]{2}\d{3,}')
_VALUE_CELL_RE = re.compile(r'^[≈≥≤<>＜＞约近]?\s*\d[\d,. ]*\s*[^\d\s]{0,4}$')

# 各种模板对「节标题」「目录」的叫法
_SECTION_LAYOUT_RE = re.compile(r'(节标题|章节标题|section\s*header)', re.I)
_TOC_LAYOUT_RE = re.compile(r'(目录|议程|toc|agenda|contents)', re.I)

# 版式占位符里不该写进正文的元信息
_META_PLACEHOLDERS = {
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.HEADER,
    PP_PLACEHOLDER.SLIDE_NUMBER,
}

# 正文类占位符：默认带项目符号
_BODY_PLACEHOLDERS = {
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.VERTICAL_BODY,
    PP_PLACEHOLDER.VERTICAL_OBJECT,
}

_BOOK_TITLE_RE = re.compile(r'《[^》]*》')

# 主文本上方紧挨着的短语（≤30 字）若水平对齐、垂直方向也贴着，视为小节标题，
# 不再把主文本提前。判据用「两形状顶端距离」而不是「上一形状底边」：占位符的
# 文本框往往比文字高得多，用底边会把整页大框误判成标题。
_HEADING_GAP = 1500000      # 两形状顶端相差不超过 1500000 EMU（约 4 厘米）
_HEADING_MAX_HEIGHT = 1600000  # 小节标题框本身不能太高（约 4.4 厘米）
_HEADING_OVERLAP = 0.5      # 与主文本至少重叠标题宽度的 50%


def _strip_cjk_gaps(text: str) -> str:
    text = re.sub(r'(?<=[\u4e00-\u9fff]) +(?=[\u4e00-\u9fff])', '', text)
    text = re.sub(r'(?<=[\u4e00-\u9fff])\u3000+(?=[\u4e00-\u9fff])', '', text)
    return text


def remove_cjk_spaces(text: str) -> str:
    """去除中文字符之间的空格，书名号里的空格保持原样"""
    out = []
    pos = 0
    for m in _BOOK_TITLE_RE.finditer(text):
        out.append(_strip_cjk_gaps(text[pos:m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(_strip_cjk_gaps(text[pos:]))
    return ''.join(out)


def _collapse_spaces(text: str) -> str:
    """把换行、制表符、全角空格与连续空白压成单个半角空格"""
    text = text.replace('\x0b', ' ').replace('\r', ' ').replace('\n', ' ')
    text = text.replace('\u00a0', ' ').replace('\u3000', ' ')
    return re.sub(r'[ \t]+', ' ', text).strip()


def _fmt_value(value) -> str:
    """图表数值：整数值不带小数点，None 写成空串"""
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# ══════════════════════════════════════════════════════════════════
# 图片导出
# ══════════════════════════════════════════════════════════════════

class _MediaDir:
    """图片落盘器：懒建目录，同一张图只写一次"""

    def __init__(self, directory: Path, rel_prefix: str):
        self.directory = directory
        self.rel_prefix = rel_prefix
        self.written = {}
        self.count = 0

    def save(self, blob: bytes, ext: str) -> str:
        key = hashlib.sha1(blob).hexdigest()
        if key in self.written:
            return self.written[key]

        safe_ext = re.sub(r'[^a-z0-9]', '', (ext or 'png').lower()) or 'png'
        self.count += 1
        name = f"image{self.count}.{safe_ext}"
        try:
            self.directory.mkdir(parents=True, exist_ok=True)
            (self.directory / name).write_bytes(blob)
        except OSError as e:
            print(f"警告: 图片写入失败 {name}: {e}", file=sys.stderr)
            return ""

        rel = f"{self.rel_prefix}/{name}"
        self.written[key] = rel
        return rel


_MEDIA = None


def _md_url(path: str) -> str:
    """图片地址里有空格或括号时用尖括号包住，避免 Markdown 解析截断"""
    if any(ch in path for ch in ' ()<>'):
        return f"<{path}>"
    return path


def _metafile_to_png(blob: bytes):
    """把 EMF/WMF 转成 PNG

    Markdown 阅读器基本都不支持 EMF/WMF，PowerPoint 里这类图又很常见
    （Visio 对象预览、粘贴的 CAD 图）。Pillow 在 Windows 上能读这两种格式，
    转换失败就保留原文件，不丢内容。
    """
    if Image is None:
        return None
    try:
        with Image.open(io.BytesIO(blob)) as image:
            image.load()
            out = io.BytesIO()
            image.save(out, format='PNG')
            return out.getvalue()
    except Exception:
        return None


def _save_image(blob: bytes, ext: str, alt: str) -> str:
    if _MEDIA is None:
        return ""
    ext = (ext or 'png').lower()
    if ext in ('emf', 'wmf'):
        png = _metafile_to_png(blob)
        if png is not None:
            blob, ext = png, 'png'
    path = _MEDIA.save(blob, ext)
    if not path:
        return f"> [图片导出失败：{alt or '未命名'}]"
    alt = (alt or "").replace(']', '\\]')
    return f"![{alt}]({_md_url(path)})"


def _shape_image(shape):
    """取形状里的图片；不是图片形状时返回 None（占位符会抛异常）"""
    try:
        return shape.image
    except (AttributeError, ValueError, KeyError, TypeError):
        return None


def _shape_alt(shape) -> str:
    """读图片的替换文字（alt text），没有就返回空串"""
    try:
        c_nv_pr = shape._element.xpath('.//p:cNvPr')
    except (AttributeError, KeyError):
        return ""
    for node in c_nv_pr:
        descr = (node.get('descr') or '').strip()
        if descr:
            return descr
    return ""


def _in_fallback(node) -> bool:
    """节点是否位于 mc:Fallback 里（与 mc:Choice 是同一份内容的两种写法）"""
    parent = node.getparent()
    while parent is not None:
        if parent.tag.endswith('}Fallback'):
            return True
        parent = parent.getparent()
    return False


def _related_part(part, rid: str):
    """按关系 id 取目标部件；缺失或指向外部链接时返回 None"""
    try:
        return part.rels[rid].target_part
    except (KeyError, AttributeError, TypeError):
        return None


# ══════════════════════════════════════════════════════════════════
# 行内格式
# ══════════════════════════════════════════════════════════════════

@dataclass
class TextLine:
    text: str
    is_list: bool = False


def _rpr_flags(rpr):
    """从 a:rPr 读加粗/斜体/删除线"""
    if rpr is None:
        return False, False, False
    bold = rpr.get('b') in ('1', 'true')
    italic = rpr.get('i') in ('1', 'true')
    strike = rpr.get('strike') in ('sngStrike', 'dblStrike')
    return bold, italic, strike


def _hyperlink_url(run) -> Optional[str]:
    try:
        return run.hyperlink.address
    except (AttributeError, KeyError, ValueError):
        return None


def _inline_text(para) -> str:
    """把一个 a:p 渲染成带行内标记的文本

    不能只遍历 para.runs：a:br（手动换行）不在 runs 里，直接拼 run.text
    会把两行粘成一行。这里按 a:p 的子元素顺序处理 r / br / fld。
    """
    runs = {run._r: run for run in para.runs}
    segs = []

    for child in para._p:
        if child.tag == qn('a:r'):
            run = runs.get(child)
            if run is None:
                continue
            text = run.text
            if not text:
                continue
            bold, italic, strike = _rpr_flags(run._r.find(qn('a:rPr')))
            segs.append((text, bold, italic, strike, _hyperlink_url(run)))
        elif child.tag == qn('a:br'):
            segs.append((' ', False, False, False, None))
        elif child.tag == qn('a:fld'):
            text = child.findtext(qn('a:t')) or ''
            if not text:
                continue
            bold, italic, strike = _rpr_flags(child.find(qn('a:rPr')))
            segs.append((text, bold, italic, strike, None))

    # 相邻且格式相同的片段先合并，避免出现 **前半****后半**
    merged = []
    for text, bold, italic, strike, url in segs:
        if not text:
            continue
        if merged:
            ptext, pbold, pitalic, pstrike, purl = merged[-1]
            if (pbold, pitalic, pstrike, purl) == (bold, italic, strike, url):
                merged[-1] = (ptext + text, bold, italic, strike, url)
                continue
        merged.append((text, bold, italic, strike, url))

    parts = []
    for text, bold, italic, strike, url in merged:
        # 首尾空格留在标记外面，否则 ** 紧贴空格会失效
        core = text.strip()
        lead = text[:len(text) - len(text.lstrip())]
        trail = text[len(text.rstrip()):]
        if not core:
            parts.append(text)
            continue

        decorated = core
        if strike:
            decorated = f"~~{decorated}~~"
        if bold:
            decorated = f"**{decorated}**"
        if italic:
            decorated = f"*{decorated}*"
        if url:
            decorated = f"[{decorated}]({url})"
        parts.append(lead + decorated + trail)

    return _collapse_spaces(''.join(parts))


# ══════════════════════════════════════════════════════════════════
# 段落、项目符号与编号
# ══════════════════════════════════════════════════════════════════

def _p_pr(para):
    return para._p.find(qn('a:pPr'))


def _bullet_kind(para, shape, single_short: bool) -> str:
    """判断段落的列表类型：none / bullet / number

    显式写了 buNone / buChar / buAutoNum 的按显式来；正文占位符没写时
    继承版式，默认当项目符号。只有一个短段落的正文占位符多半是副标题，
    不加项目符号更好看。
    """
    p_pr = _p_pr(para)
    if p_pr is not None:
        if p_pr.find(qn('a:buNone')) is not None:
            return 'none'
        if p_pr.find(qn('a:buAutoNum')) is not None:
            return 'number'
        if p_pr.find(qn('a:buChar')) is not None or p_pr.find(qn('a:buBlip')) is not None:
            return 'bullet'

    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        if ph_type in _BODY_PLACEHOLDERS:
            if single_short and para.level == 0:
                return 'none'
            return 'bullet'
    return 'none'


def _auto_num_prefix(p_pr, counters) -> str:
    """把 a:buAutoNum 还原成 1. / 1) / (1) / 1 这几种写法"""
    node = p_pr.find(qn('a:buAutoNum'))
    level = int(node.get('startAt') or 1) + counters.get('_num', 0)
    counters['_num'] = counters.get('_num', 0) + 1
    num_type = node.get('type') or 'arabicPeriod'
    if 'ParenR' in num_type:
        return f"{level})"
    if 'ParenBoth' in num_type:
        return f"({level})"
    if 'Plain' in num_type:
        return str(level)
    return f"{level}."


def _list_marker(para, shape, counters, single_short: bool):
    kind = _bullet_kind(para, shape, single_short)
    if kind == 'number':
        return _auto_num_prefix(_p_pr(para), counters)
    if kind == 'bullet':
        return '-'
    return None


def _text_entries(shape) -> list:
    """把一个文本框渲染成若干行，保留缩进、项目符号与行内格式"""
    paragraphs = list(shape.text_frame.paragraphs)
    nonempty = [p for p in paragraphs if _inline_text(p).strip()]
    single_short = len(nonempty) == 1 and len(_inline_text(nonempty[0])) <= 20

    entries = []
    counters = {}
    for para in paragraphs:
        text = _inline_text(para)
        if not text.strip():
            continue
        level = max(0, min(para.level, 8))
        marker = _list_marker(para, shape, counters, single_short)
        if marker is None:
            counters.clear()
            line = text.strip()
            if level:
                line = '  ' * level + line
            entries.append(TextLine(line, False))
        else:
            if marker == '-':
                counters.clear()
            entries.append(TextLine('  ' * level + marker + ' ' + text.strip(), True))
    return entries


def _entries_to_markdown(entries) -> str:
    """列表项之间不空行，普通段落之间空一行"""
    out = []
    prev_list = None
    for entry in entries:
        if out and not (entry.is_list and prev_list):
            out.append('')
        out.append(entry.text)
        prev_list = entry.is_list
    return '\n'.join(out)


def _text_block(shape) -> str:
    return _entries_to_markdown(_text_entries(shape))


# ══════════════════════════════════════════════════════════════════
# 表格
# ══════════════════════════════════════════════════════════════════

def _cell_text(cell) -> str:
    """单元格内容：多段落折成一行，保留行内加粗/斜体"""
    parts = []
    for para in cell.text_frame.paragraphs:
        text = _inline_text(para)
        if text.strip():
            parts.append(text.strip())
    return remove_cjk_spaces(_collapse_spaces(' '.join(parts)))


def _table_grid(table) -> list:
    """按网格提取表格，合并单元格只在左上角写一次，其余位置留空"""
    rows = len(table.rows)
    cols = len(table.columns)
    grid = [[''] * cols for _ in range(rows)]
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            if ci >= cols:
                continue
            if cell.is_spanned:
                continue
            grid[ri][ci] = _cell_text(cell)
    return grid


def _strip_inline_markup(text: str) -> str:
    """去掉行内 Markdown 标记，用于判断表格行是表头还是数据"""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'~~(.+?)~~', r'\1', text)
    text = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'\1', text)
    return text.replace('\\|', '|').strip()


def _is_header_row(row) -> bool:
    """按内容猜测某行是不是表头（短文本、无数值）"""
    non_empty = [_strip_inline_markup(str(t or '')) for t in row]
    non_empty = [text for text in non_empty if text]
    if not non_empty:
        return True
    for text in non_empty:
        if _CODE_PATTERN.match(text):
            return False
    if any(_VALUE_CELL_RE.match(text) for text in non_empty):
        return False
    if len(non_empty) >= 3:
        avg = sum(len(text) for text in non_empty) / len(non_empty)
        # PowerPoint 的多级表头常写成「估算价值（万元）」「设备及工器具购置费」
        # 这类长标签，阈值比 Word 宽一些，否则会把第二行表头当成数据
        return avg <= 8
    return True


def _header_count(grid) -> int:
    """推断表头行数

    从第一行开始数连续的表头行；如果整张表都被判成表头（纯文字表格），
    只留首行当表头，否则所有数据都会被并进表头。
    """
    if not grid:
        return 0
    count = 0
    for row in grid:
        if not _is_header_row(row):
            break
        count += 1
    if count == 0 or count >= len(grid):
        return 1
    return count


def _render_table(table) -> str:
    """把 pptx 表格渲染成 GFM 表格，多级表头拍平成一行"""
    grid = _table_grid(table)
    if not grid:
        return ''

    header_count = _header_count(grid)
    header_rows = grid[:header_count]
    data_rows = grid[header_count:]

    cols = max(len(row) for row in grid)
    headers = []
    for ci in range(cols):
        parts = []
        for row in header_rows:
            value = remove_cjk_spaces(str(row[ci] if ci < len(row) else '')).strip()
            # 纵向合并的单元格在多行表头里会重复同一文字，只保留一次
            if value and (not parts or parts[-1] != value):
                parts.append(value)
        headers.append(' '.join(parts))

    # 整列都空的列不输出
    keep = []
    for ci in range(cols):
        has_header = bool(headers[ci].strip())
        has_data = any(ci < len(row) and row[ci].strip() for row in data_rows)
        if has_header or has_data:
            keep.append(ci)
    if not keep:
        return ''

    headers = [headers[ci] for ci in keep]
    if not any(headers) and data_rows:
        # 判不出表头时把首行数据提上来，避免出现空的假表头
        headers = [str(data_rows[0][ci] if ci < len(data_rows[0]) else '') for ci in keep]
        data_rows = data_rows[1:]

    rows = []
    for row in data_rows:
        rows.append([str(row[ci] if ci < len(row) else '') for ci in keep])

    def escape(text: str) -> str:
        return _collapse_spaces(text).replace('|', '\\|')

    lines = ['| ' + ' | '.join(escape(h) for h in headers) + ' |',
             '| ' + ' | '.join('---' for _ in headers) + ' |']
    for row in rows:
        lines.append('| ' + ' | '.join(escape(c) for c in row) + ' |')
    return '\n'.join(lines)


# ══════════════════════════════════════════════════════════════════
# 图表
# ══════════════════════════════════════════════════════════════════

def _chart_title(chart) -> str:
    try:
        if chart.has_title:
            return _collapse_spaces(chart.chart_title.text_frame.text)
    except (AttributeError, ValueError):
        pass
    return ''


def _render_chart(shape) -> list:
    """把图表的数据系列写成 Markdown 表格，读文档时至少能看到原始数值"""
    chart = shape.chart
    title = _chart_title(chart)

    try:
        plot = chart.plots[0]
        categories = [str(c) for c in plot.categories]
    except (IndexError, AttributeError, ValueError, TypeError):
        categories = []

    series = []
    try:
        for item in chart.series:
            name = _collapse_spaces(str(item.name or ''))
            series.append((name, list(item.values)))
    except (AttributeError, ValueError, TypeError):
        series = []

    if not series and not categories:
        return [f"> **图表**：{title}" if title else '> **图表**：无数据']

    if not categories:
        length = max((len(values) for _, values in series), default=0)
        categories = [str(i + 1) for i in range(length)]

    headers = ['类别'] + [name or f'系列{i + 1}' for i, (name, _) in enumerate(series)]
    rows = []
    for ci, category in enumerate(categories):
        row = [category]
        for _, values in series:
            row.append(_fmt_value(values[ci]) if ci < len(values) else '')
        rows.append(row)

    lines = []
    if title:
        lines.append(f"**{title}**")
        lines.append('')
    lines.append('| ' + ' | '.join(headers) + ' |')
    lines.append('| ' + ' | '.join('---' for _ in headers) + ' |')
    for row in rows:
        lines.append('| ' + ' | '.join(_collapse_spaces(c).replace('|', '\\|') for c in row) + ' |')
    return ['\n'.join(lines)]


# ══════════════════════════════════════════════════════════════════
# 组合图形、嵌入对象与普通形状
# ══════════════════════════════════════════════════════════════════

def _shape_text(shape) -> str:
    if not getattr(shape, 'has_text_frame', False):
        return ''
    return _collapse_spaces('\n'.join(p.text for p in shape.text_frame.paragraphs))


def _shape_pos(shape):
    return (shape.top or 0, shape.left or 0)


def _simple_text_shape(shape) -> bool:
    """单段文本、没有表格/图片/图表的小标签，用于目录页合并「1 项目背景」"""
    if not getattr(shape, 'has_text_frame', False):
        return False
    if getattr(shape, 'has_table', False) or getattr(shape, 'has_chart', False):
        return False
    if shape.shape_type in (MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.PICTURE):
        return False
    paragraphs = [p for p in shape.text_frame.paragraphs if _inline_text(p).strip()]
    return len(paragraphs) == 1 and not shape.text_frame.text.strip().startswith('•')


def _render_toc_group(group) -> list:
    """目录页的组合形状：把同一行的序号与标题合成「1 项目背景」"""
    items = []
    for index, shape in enumerate(group.shapes):
        if _simple_text_shape(shape):
            items.append((_shape_pos(shape) + (index,), shape))
        else:
            # 非文字元素单独按位置排，不参与合行
            items.append((_shape_pos(shape) + (index,), shape))
    items.sort(key=lambda item: item[0])

    lines = []
    row = []
    row_top = None
    row_height = 0

    def flush():
        if not row:
            return
        ordered = sorted(row, key=lambda shape: shape.left or 0)
        texts = [_shape_text(shape).strip() for shape in ordered]
        texts = [text for text in texts if text]
        if texts:
            lines.append('- ' + ' '.join(texts))
        row.clear()

    for (top, _left, _index), shape in items:
        if not _simple_text_shape(shape):
            flush()
            row_top = None
            row_height = 0
            continue
        height = shape.height or 0
        if row and row_top is not None and top < row_top + row_height:
            row.append(shape)
            row_height = max(row_height, height)
        else:
            flush()
            row = [shape]
            row_top = top
            row_height = height
    flush()

    return ['\n'.join(lines)] if lines else []


def _render_group(group, toc: bool) -> list:
    """组合图形：按位置从上到下、从左到右展开"""
    if toc:
        return _render_toc_group(group)

    items = sorted(enumerate(group.shapes), key=lambda item: _shape_pos(item[1]) + (item[0],))
    blocks = []
    for _index, shape in items:
        blocks.extend(_render_shape(shape, toc=False))
    return blocks


def _ole_preview(shape) -> list:
    """嵌入对象（Visio/Excel 等）：导出 OLE 的预览图，保留视觉效果"""
    prog_id = ''
    for node in shape._element.iter():
        if node.tag.endswith('}oleObj'):
            prog_id = node.get('progId') or ''
            break

    # mc:Choice 里通常只有 OLE 引用，预览图在 mc:Fallback 的 p:pic 里
    for blip in shape._element.iter(qn('a:blip')):
        rid = blip.get(qn('r:embed'))
        if not rid:
            continue
        part = _related_part(shape.part, rid)
        if part is None:
            continue
        try:
            ext = Path(str(part.partname)).suffix.lstrip('.').lower() or 'png'
            alt = '嵌入对象预览'
            if prog_id:
                alt += f"（{prog_id}）"
            return [_save_image(part.blob, ext, alt)]
        except (AttributeError, TypeError):
            continue

    label = prog_id or '未提供预览'
    return [f"> **嵌入对象**：{label}"]


def _graphic_texts(shape) -> list:
    """SmartArt 等图形框架没有 has_text_frame，文字藏在图形数据里"""
    texts = []

    def collect(root):
        for node in root.iter(qn('a:t')):
            if _in_fallback(node):
                continue
            text = _collapse_spaces(node.text or '')
            if text and (not texts or texts[-1] != text):
                texts.append(text)

    # 先看形状自身（普通图示的回退内容）
    collect(shape._element)

    # SmartArt 的数据模型在单独的部件里，通过 dgm:relIds/@r:dm 关联
    if not texts:
        for rel_node in shape._element.iter():
            if rel_node.tag.endswith('}relIds'):
                for attr in ('dm', 'lo'):
                    rid = rel_node.get(qn(f'r:{attr}'))
                    if not rid:
                        continue
                    part = _related_part(shape.part, rid)
                    if part is None:
                        continue
                    try:
                        collect(part.element)
                    except (AttributeError, TypeError):
                        continue
                break
    return texts


def _render_shape(shape, toc: bool = False) -> list:
    """把一个形状渲染成若干 Markdown 块"""
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return _render_group(shape, toc)
    if getattr(shape, 'has_table', False):
        text = _render_table(shape.table)
        return [text] if text else []
    if getattr(shape, 'has_chart', False):
        return _render_chart(shape)
    if shape.shape_type in (MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT,
                            MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
        return _ole_preview(shape)

    image = _shape_image(shape)
    if image is not None:
        return [_save_image(image.blob, image.ext, _shape_alt(shape))]

    if getattr(shape, 'has_text_frame', False):
        text = _text_block(shape)
        return [text] if text else []

    # SmartArt / 图示：抽图形数据里的文字
    texts = _graphic_texts(shape)
    if texts:
        return ['\n'.join(texts)]

    # 纯装饰形状（线条、无文字矩形）直接跳过；图形框架留一个可见标记
    if shape.shape_type in (MSO_SHAPE_TYPE.DIAGRAM, MSO_SHAPE_TYPE.IGX_GRAPHIC):
        return [f"> [未支持的图形：{shape.name}]"]
    return []


# ══════════════════════════════════════════════════════════════════
# 幻灯片结构
# ══════════════════════════════════════════════════════════════════

def _title_of(slide):
    """返回 (标题文字, 标题形状)，没有标题时返回 (None, None)"""
    shape = slide.shapes.title
    if shape is None:
        return None, None
    text = remove_cjk_spaces(_collapse_spaces(shape.text_frame.text))
    return (text or None), shape


def _is_section_slide(slide, layout_name: str) -> bool:
    return bool(_SECTION_LAYOUT_RE.search(layout_name or ''))


def _section_number(slide, title_shape):
    """节标题页上的「1」一般放在正文占位符里，取出来后不重复输出"""
    if title_shape is None:
        return None, None
    for shape in slide.placeholders:
        if shape is title_shape:
            continue
        if shape.placeholder_format.type not in _BODY_PLACEHOLDERS:
            continue
        text = _collapse_spaces(shape.text_frame.text)
        if re.fullmatch(r'\d{1,3}', text):
            return text, shape
    return None, None


def _toc_slide(slide, layout_name: str) -> bool:
    if _TOC_LAYOUT_RE.search(layout_name or ''):
        return True
    title, _shape = _title_of(slide)
    return bool(title and _TOC_LAYOUT_RE.search(title))


def _is_primary_text(shape, slide_area: int) -> bool:
    """判断是不是这一页的主叙述文本

    主叙述文本先输出，图片、标注、图形按位置排到后面。否则像「设计规模」
    这种位于页面上方的标注会跑到正文前面。
    """
    if shape.shape_type in (MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.PICTURE):
        return False
    if getattr(shape, 'has_table', False) or getattr(shape, 'has_chart', False):
        return False
    if not getattr(shape, 'has_text_frame', False):
        return False
    if shape.is_placeholder and shape.placeholder_format.type in _BODY_PLACEHOLDERS:
        return bool(_shape_text(shape))
    area = (shape.width or 0) * (shape.height or 0)
    return area >= slide_area * 0.20 and len(_shape_text(shape)) >= 30


def _short_heading_above(shape, candidates) -> bool:
    """主文本正上方是否紧贴着一行小标题

    「问题与建议」这类小节标题本身很短、面积也小，会被当成次要形状排到后面，
    于是两条建议先出现、标题掉到末尾。判据是水平方向基本对齐、垂直方向紧挨着；
    页面上方远离正文的图形标注（如「设计规模12万吨/天」）不满足贴住条件，
    它们仍然排在正文之后。
    """
    top = shape.top or 0
    left = shape.left or 0
    width = shape.width or 0
    for other in candidates:
        if other is shape or (other.top or 0) >= top:
            continue
        text = _shape_text(other)
        if not text or len(text) > 30:
            continue
        if top - (other.top or 0) > _HEADING_GAP:
            continue
        if (other.height or 0) > _HEADING_MAX_HEIGHT:
            continue
        other_left = other.left or 0
        other_width = other.width or 0
        if not other_width:
            continue
        overlap = min(left + width, other_left + other_width) - max(left, other_left)
        if overlap >= other_width * _HEADING_OVERLAP:
            return True
    return False


def _render_notes(slide) -> list:
    if not slide.has_notes_slide:
        return []
    frame = slide.notes_slide.notes_text_frame
    if frame is None:
        return []
    paragraphs = [_collapse_spaces(p.text) for p in frame.paragraphs]
    paragraphs = [p for p in paragraphs if p]
    if not paragraphs:
        return []
    lines = []
    for index, text in enumerate(paragraphs):
        prefix = '> **备注**：' if index == 0 else '>'
        lines.append(f"{prefix} {text}")
    return ['\n'.join(lines)]


def _render_slide(slide, index: int, level: int, heading: Optional[str],
                  toc: bool, include_notes: bool, slide_markers: bool,
                  slide_area: int, skip_ids: set) -> list:
    """渲染一页幻灯片，返回若干个 Markdown 块"""
    blocks = []
    if slide_markers:
        blocks.append(f"<!-- slide {index} -->")
    if heading:
        blocks.append(f"{'#' * level} {heading}")

    candidates = []
    for shape in slide.shapes:
        if id(shape._element) in skip_ids:
            continue
        if shape.is_placeholder and shape.placeholder_format.type in _META_PLACEHOLDERS:
            continue
        candidates.append(shape)

    primary_ids = {id(shape._element)
                   for shape in candidates
                   if _is_primary_text(shape, slide_area)
                   and not _short_heading_above(shape, candidates)}
    primary = [shape for shape in candidates if id(shape._element) in primary_ids]
    secondary = [(index, shape) for index, shape in enumerate(candidates)
                 if id(shape._element) not in primary_ids]
    secondary.sort(key=lambda item: _shape_pos(item[1]) + (item[0],))

    for shape in primary:
        blocks.extend(_render_shape(shape, toc=toc))
    for _index, shape in secondary:
        blocks.extend(_render_shape(shape, toc=toc))

    # 目录页每一项如果来自不同组合，块之间会被空行隔开，这里合成一个紧凑列表
    if toc:
        merged = []
        toc_lines = []
        for block in blocks:
            if block.startswith('- '):
                toc_lines.extend(block.splitlines())
            else:
                if toc_lines:
                    merged.append('\n'.join(toc_lines))
                    toc_lines = []
                merged.append(block)
        if toc_lines:
            merged.append('\n'.join(toc_lines))
        blocks = merged

    if include_notes:
        blocks.extend(_render_notes(slide))
    return [block for block in blocks if block and block.strip()]


def _heading_with_number(title: Optional[str], number: Optional[str]) -> Optional[str]:
    """把节标题页的编号拼到标题前面，标题本身已带同类序号时不重复"""
    if not title or not number:
        return title
    if re.match(rf'^{re.escape(number)}[\s、.．]', title):
        return title
    if re.match(r'^第[一二三四五六七八九十0-9]+[章节部分]', title):
        return title
    return f"{number} {title}"


def render_presentation(prs, stem: str, include_notes: bool = True,
                        slide_markers: bool = False) -> str:
    """把整个演示文稿渲染成 Markdown"""
    slides = list(prs.slides)
    if not slides:
        return f"# {stem}\n"

    slide_area = (prs.slide_width or 1) * (prs.slide_height or 1)
    first_title, first_shape = _title_of(slides[0])
    first_layout = slides[0].slide_layout.name or ''
    first_is_section = _is_section_slide(slides[0], first_layout)
    needs_doc_title = not first_title or first_is_section

    blocks = []
    if needs_doc_title:
        blocks.append(f"# {stem}")

    section_seen = False
    for index, slide in enumerate(slides, 1):
        layout_name = slide.slide_layout.name or ''
        title, title_shape = _title_of(slide)
        is_section = _is_section_slide(slide, layout_name)
        toc = _toc_slide(slide, layout_name)

        number = None
        number_shape = None
        if title_shape is not None:
            number, number_shape = _section_number(slide, title_shape)
        title = _heading_with_number(title, number)

        if index == 1 and title and not is_section and first_shape is not None:
            level = 1
        elif is_section:
            level = 2
        else:
            level = 3 if section_seen else 2

        # 目录页没有标题占位符时，用版式名当标题
        if not title and toc:
            title = '目录'
            level = 3 if section_seen else 2

        skip_ids = set()
        if title_shape is not None:
            skip_ids.add(id(title_shape._element))
        if number_shape is not None:
            skip_ids.add(id(number_shape._element))

        blocks.extend(_render_slide(
            slide, index, level, title, toc=toc,
            include_notes=include_notes,
            slide_markers=slide_markers,
            slide_area=slide_area,
            skip_ids=skip_ids,
        ))

        if is_section:
            section_seen = True

    return '\n\n'.join(block for block in blocks if block and block.strip()) + '\n'


# ══════════════════════════════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════════════════════════════

def convert(pptx_path: str, include_notes: bool = True,
            slide_markers: bool = False) -> str:
    """将 pptx 转换为 markdown 文本"""
    global _MEDIA
    src = Path(pptx_path)
    media_dir = f"{src.stem}_media"
    _MEDIA = _MediaDir(src.parent / media_dir, media_dir)
    try:
        prs = Presentation(str(src))
        return render_presentation(prs, src.stem, include_notes=include_notes,
                                   slide_markers=slide_markers)
    finally:
        _MEDIA = None


def _convert_one(src: Path, include_notes: bool, slide_markers: bool) -> bool:
    """将单个 pptx 文件转换为同目录同名的 md 文件，成功返回 True"""
    try:
        content = convert(str(src), include_notes=include_notes,
                          slide_markers=slide_markers)
    except Exception as e:
        print(f"转换失败: {src}: {e}", file=sys.stderr)
        return False

    out_path = src.with_suffix('.md')
    try:
        out_path.write_text(content, encoding='utf-8')
    except OSError as e:
        print(f"写入失败: {out_path}: {e}", file=sys.stderr)
        return False

    print(f"已生成: {out_path}")
    return True


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog='pptx2md',
        description='将 PowerPoint (.pptx) 文件转换为 Markdown，输出同目录同名 .md 文件',
    )
    parser.add_argument('path', nargs='?', help='pptx 文件路径（与 -d 二选一）')
    parser.add_argument('-d', '--dir', metavar='DIR',
                        help='将目录下所有 pptx 文件转换为 md 文件')
    parser.add_argument('--no-notes', action='store_true',
                        help='不输出演讲者备注（默认输出）')
    parser.add_argument('--slide-markers', action='store_true',
                        help='每页前加 <!-- slide N --> 注释，便于溯源')
    args = parser.parse_args(argv)

    include_notes = not args.no_notes

    # 目录模式：转换目录下所有 pptx
    if args.dir:
        dir_path = Path(args.dir)
        if not dir_path.is_dir():
            print(f"错误: 目录不存在: {dir_path}", file=sys.stderr)
            return 1
        pptx_files = sorted(
            path for path in dir_path.iterdir()
            if path.is_file()
            and path.suffix.lower() in ('.pptx', '.pptm')
            and not path.name.startswith('~$')
        )
        if not pptx_files:
            print(f"错误: 目录中没有 .pptx 文件: {dir_path}", file=sys.stderr)
            return 1
        success = 0
        for path in pptx_files:
            if _convert_one(path, include_notes, args.slide_markers):
                success += 1
        print(f"完成: 成功 {success} / 共 {len(pptx_files)} 个文件")
        return 0 if success == len(pptx_files) else 1

    # 单文件模式
    if not args.path:
        parser.error('必须提供文件路径或 -d/--dir 目录')
    src = Path(args.path)
    if not src.exists():
        print(f"错误: 文件不存在: {src}", file=sys.stderr)
        return 1
    if src.suffix.lower() not in ('.pptx', '.pptm'):
        print(f"错误: 仅支持 .pptx / .pptm 文件: {src}", file=sys.stderr)
        return 1

    return 0 if _convert_one(src, include_notes, args.slide_markers) else 1


if __name__ == '__main__':
    sys.exit(main())
