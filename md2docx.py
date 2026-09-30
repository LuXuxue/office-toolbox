#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2docx — 将 Markdown 转换为 Word (.docx) 的单文件独立脚本

用法:
    python md2docx.py path/to/file.md           转换单个文件
    python md2docx.py -d path/to/dir            转换目录下所有 .md

结果写在与 md 同一目录、文件名相同的 .docx 文件。与 docx2md 配套使用：
公式（$$…$$ / $…$）会写成可编辑的 Word 公式而非纯文本。

依赖:
    pip install python-docx

"""

import sys
import re
import os
import argparse
import glob
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.oxml import OxmlElement


def _split_table_row(line):
    """按未转义的 | 切分表格行，并把 \\| 还原为 |

    docx2md 会把单元格里的竖线转义成 \\|，直接 split('|') 会把这样的单元格
    切成两列，并丢掉后面各列的内容。
    """
    line = line.strip()
    if line.startswith('|'):
        line = line[1:]
    if line.endswith('|') and not line.endswith('\\|'):
        line = line[:-1]

    cells, cur, i = [], '', 0
    while i < len(line):
        c = line[i]
        if c == '\\' and i + 1 < len(line) and line[i + 1] == '|':
            cur += '|'
            i += 2
            continue
        if c == '|':
            cells.append(cur.strip())
            cur = ''
            i += 1
            continue
        cur += c
        i += 1
    cells.append(cur.strip())
    return cells


def _looks_like_math(tex):
    """判断 $…$ 里装的是不是公式

    docx2md 只会为公式输出 $…$，但正文里也可能出现「$100 和 $200」这类金额，
    不能一律当公式。用两条经验规则区分：含 LaTeX 命令/花括号，或整段没有空白。
    """
    if '\\' in tex or '{' in tex or '}' in tex:
        return True
    return not re.search(r'\s', tex)


def _add_formatted_text(para, text):
    """写入段落文本，识别行内公式与 markdown 行内格式"""
    if not text:
        return

    for segment in re.split(r'(\$[^$\n]+\$)', text):
        if not segment:
            continue
        if len(segment) > 2 and segment[0] == '$' and segment[-1] == '$':
            body = segment[1:-1]
            if _looks_like_math(body):
                add_inline_math(para, body)
                continue
        _add_text_segment(para, segment)


def _add_text_segment(para, text):
    """写入非公式片段，处理 `代码` 段"""
    for segment in re.split(r'(`[^`]+`)', text):
        if segment.startswith('`') and segment.endswith('`'):
            para.add_run(segment[1:-1])
        elif segment:
            _add_formatted_runs(para, segment)


def _add_formatted_runs(para, text):
    """写入加粗、斜体、粗斜体与超链接格式"""
    pattern = re.compile(
        r'(\*\*\*(.+?)\*\*\*)'
        r'|(\*\*(.+?)\*\*)'
        r'|(\*(.+?)\*)'
        r'|(\[(.+?)\]\((.+?)\))'
    )

    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            para.add_run(text[pos:m.start()])

        if m.group(2) is not None:
            run = para.add_run(m.group(2))
            run.bold = True
            run.italic = True
        elif m.group(4) is not None:
            run = para.add_run(m.group(4))
            run.bold = True
        elif m.group(6) is not None:
            run = para.add_run(m.group(6))
            run.italic = True
        elif m.group(8) is not None:
            run = para.add_run(m.group(8))
            run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
            run.font.underline = True

        pos = m.end()

    if pos < len(text):
        para.add_run(text[pos:])


def _link_default_fonts_to_theme(doc):
    """让默认字体跟随主题，参数与 Word 2016 新建文档保持一致"""
    styles_elem = doc.styles.element
    doc_defaults = styles_elem.find(qn('w:docDefaults'))
    if doc_defaults is None:
        doc_defaults = OxmlElement('w:docDefaults')
        styles_elem.insert(0, doc_defaults)

    rpr_default = doc_defaults.find(qn('w:rPrDefault'))
    if rpr_default is None:
        rpr_default = OxmlElement('w:rPrDefault')
        doc_defaults.append(rpr_default)

    rpr = rpr_default.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        rpr_default.append(rpr)

    ppr_default = doc_defaults.find(qn('w:pPrDefault'))
    if ppr_default is not None:
        ppr = ppr_default.find(qn('w:pPr'))
        if ppr is not None:
            for child in list(ppr):
                ppr.remove(child)

    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:asciiTheme'), 'minorHAnsi')
    rfonts.set(qn('w:eastAsiaTheme'), 'minorEastAsia')
    rfonts.set(qn('w:hAnsiTheme'), 'minorHAnsi')
    rfonts.set(qn('w:cstheme'), 'minorBidi')

    sz = rpr.find(qn('w:sz'))
    if sz is None:
        sz = OxmlElement('w:sz')
        rpr.append(sz)
    sz.set(qn('w:val'), '21')
    sz_cs = rpr.find(qn('w:szCs'))
    if sz_cs is None:
        sz_cs = OxmlElement('w:szCs')
        rpr.append(sz_cs)
    sz_cs.set(qn('w:val'), '22')

    lang = rpr.find(qn('w:lang'))
    if lang is None:
        lang = OxmlElement('w:lang')
        rpr.append(lang)
    lang.set(qn('w:val'), 'en-US')
    lang.set(qn('w:eastAsia'), 'zh-CN')
    lang.set(qn('w:bidi'), 'ar-SA')

    _config_theme_cn(doc)


def _config_theme_cn(doc):
    """把内置主题改成简体中文主题（等线）"""
    _set_theme_font_lang(doc)

    theme_part = None
    for rel in doc.part.rels.values():
        if 'theme' in rel.reltype:
            theme_part = rel.target_part
            break
    if theme_part is None:
        return

    from lxml import etree
    root = etree.fromstring(theme_part.blob)
    font_scheme = root.find('.//' + qn('a:fontScheme'))
    if font_scheme is None:
        return

    major = font_scheme.find('.//' + qn('a:majorFont'))
    minor = font_scheme.find('.//' + qn('a:minorFont'))

    if major is not None:
        _set_theme_fonts(major, '等线 Light')
    if minor is not None:
        _set_theme_fonts(minor, '等线')

    new_blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                              standalone=True)
    theme_part._blob = new_blob


def _set_theme_fonts(node, face):
    """设置 fontScheme 中拉丁字体与中文（Hans）字体的字体名"""
    from lxml import etree
    latin = node.find(qn('a:latin'))
    if latin is None:
        latin = etree.SubElement(node, qn('a:latin'))
    latin.set('typeface', face)

    for f in node.findall(qn('a:font')):
        if f.get('script') == 'Hans':
            f.set('typeface', face)
            return
    hans = etree.SubElement(node, qn('a:font'))
    hans.set('script', 'Hans')
    hans.set('typeface', face)


def _set_theme_font_lang(doc):
    """把 settings.xml 的 themeFontLang 东亚语言设为 zh-CN"""
    settings = doc.settings.element
    tfl = settings.find(qn('w:themeFontLang'))
    if tfl is None:
        tfl = OxmlElement('w:themeFontLang')
        math_pr = settings.find(qn('m:mathPr'))
        if math_pr is not None:
            math_pr.addnext(tfl)
        else:
            settings.append(tfl)
    tfl.set(qn('w:val'), 'en-US')
    tfl.set(qn('w:eastAsia'), 'zh-CN')


# ══════════════════════════════════════════════════════════════════
# LaTeX → OMML（公式）
#
# docx2md 把 Word 公式转成 LaTeX（块级 $$…$$、行内 $…$），这里再转回
# OMML，让 round-trip 后仍是可编辑的 Word 公式而不是纯文本。
# 只覆盖 docx2md 会输出的那部分语法；解析不了就退回原始文本，绝不丢内容。
# ══════════════════════════════════════════════════════════════════

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def _mtag(name):
    return f'{{{MATH_NS}}}{name}'


_LATEX_SYMBOLS = {
    r'\times': '×', r'\cdot': '⋅', r'\div': '÷', r'\pm': '±', r'\mp': '∓',
    r'\leq': '≤', r'\le': '≤', r'\geq': '≥', r'\ge': '≥',
    r'\neq': '≠', r'\approx': '≈', r'\sim': '∼', r'\propto': '∝',
    r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
    r'\varepsilon': 'ε', r'\mu': 'μ', r'\sigma': 'σ', r'\tau': 'τ',
    r'\theta': 'θ', r'\lambda': 'λ', r'\xi': 'ξ', r'\pi': 'π',
    r'\rho': 'ρ', r'\psi': 'ψ', r'\phi': 'φ', r'\omega': 'ω',
    r'\eta': 'η', r'\nu': 'ν', r'\zeta': 'ζ', r'\kappa': 'κ',
    r'\Delta': 'Δ', r'\Omega': 'Ω', r'\Phi': 'Φ',
    r'\infty': '∞', r'\partial': '∂', r'\prime': '′', r'\circ': '°',
    r'\ldots': '…', r'\cdots': '⋯', r'\backslash': '\\',
    r'\lbrace': '{', r'\rbrace': '}', r'\{': '{', r'\}': '}',
    r'\,': ' ', r'\;': ' ', r'\ ': ' ', r'\quad': '　', r'\qquad': '　',
    r'\%': '%', r'\$': '$', r'\&': '&', r'\#': '#', r'\_': '_',
}
_LATEX_ACCENTS = {
    r'\hat': '̂', r'\bar': '̄', r'\tilde': '̃',
    r'\vec': '⃗', r'\dot': '̇', r'\ddot': '̈',
    r'\check': '̌', r'\breve': '̆',
}
_LATEX_NARY = {
    r'\sum': '∑', r'\prod': '∏', r'\coprod': '∐',
    r'\int': '∫', r'\iint': '∬', r'\iiint': '∭', r'\oint': '∮',
}
_LATEX_DELIMS = {
    '(': '(', '[': '[', '{': '{', '|': '|', '.': '',
    r'\langle': '⟨', r'\rangle': '⟩', r'\|': '‖',
}
_TEXT_STYLE_COMMANDS = {
    r'\text': 'nor', r'\mathrm': 'plain', r'\mathbf': 'bold',
    r'\mathit': None, r'\operatorname': 'plain', r'\textrm': 'nor',
    r'\textbf': 'bold',
}

_TOKEN_RE = re.compile(r"""
    (?P<cmd>\\[A-Za-z]+)
  | (?P<esc>\\.)
  | (?P<word>[A-Za-z]+)
  | (?P<num>\d+(?:\.\d+)?)
  | (?P<open>\{)   | (?P<close>\})
  | (?P<lbrack>\[) | (?P<rbrack>\])
  | (?P<sub>_)    | (?P<sup>\^)
  | (?P<amp>&)
  | (?P<other>.)
""", re.X)


class _LatexParser:
    """把 docx2md 生成的 LaTeX 子集解析成语法树"""

    def __init__(self, latex):
        self.tokens = [(m.lastgroup, m.group()) for m in _TOKEN_RE.finditer(latex)]
        self.pos = 0
        self._drop_space = False

    def peek(self):
        return self.tokens[self.pos] if self.pos < len(self.tokens) else (None, None)

    def take(self):
        kind, val = self.peek()
        self.pos += 1
        return kind, val

    def parse(self):
        return self.parse_expr()

    def parse_expr(self):
        nodes = []
        while True:
            kind, val = self.peek()
            if kind is None or kind == 'close' or kind == 'rbrack':
                break
            if kind == 'cmd' and val in (r'\right', r'\end'):
                break
            # 符号宏（\times \sigma …）后面紧跟的空格不必保留，
            # 否则再转回 LaTeX 时会出现多余空格；
            # 但空格后面跟着上/下标时必须留给 parse_scripted，否则会拆散「基+标」
            if (kind == 'other' and val == ' ' and self._drop_space
                    and not self._script_follows()):
                self.take()
                continue
            self._drop_space = False
            node = self.parse_scripted()
            if node is not None:
                nodes.append(node)
        return nodes

    def _script_follows(self) -> bool:
        """跳过空格后是否是上标/下标标记"""
        i = self.pos
        while i < len(self.tokens) and self.tokens[i] == ('other', ' '):
            i += 1
        return i < len(self.tokens) and self.tokens[i][0] in ('sub', 'sup')

    def parse_scripted(self):
        base = self.parse_atom()
        if base is None:
            return None
        sub = sup = None
        while True:
            kind, val = self.peek()
            if kind == 'other' and val == ' ' and self._script_follows():
                self.take()  # 吃掉「基 空格 上标」中间的空格，保持关联
                continue
            if kind == 'sub':
                self.take()
                sub = self.parse_group()
            elif kind == 'sup':
                self.take()
                sup = self.parse_group()
            else:
                break
        if sub is None and sup is None:
            return base
        if sub is not None and sup is not None:
            return ('subsup', base, sub, sup)
        if sub is not None:
            return ('sub', base, sub)
        return ('sup', base, sup)

    def parse_group(self):
        kind, val = self.peek()
        if kind == 'open':
            self.take()
            nodes = self.parse_expr()
            if self.peek()[0] == 'close':
                self.take()
            return nodes
        node = self.parse_atom()
        return [node] if node is not None else []

    def parse_atom(self):
        kind, val = self.take()
        if kind is None or kind in ('close', 'sub', 'sup'):
            return None
        if kind == 'open':
            nodes = self.parse_expr()
            if self.peek()[0] == 'close':
                self.take()
            return ('group', nodes)
        if kind == 'cmd':
            return self.parse_command(val)
        if kind == 'esc':
            return ('text', _LATEX_SYMBOLS.get(val, val[1:]), None)
        if kind == 'amp':
            return None
        if kind in ('word', 'num', 'other'):
            return ('text', val, None)
        return None

    def parse_command(self, name):
        if name == r'\frac':
            num = self.parse_group()
            den = self.parse_group()
            return ('frac', num, den)
        if name == r'\sqrt':
            deg = None
            if self.peek()[0] == 'lbrack':
                self.take()
                deg = self.parse_expr()
                if self.peek()[0] == 'rbrack':
                    self.take()
            return ('rad', deg, self.parse_group())
        if name == r'\left':
            beg = self._read_delim()
            body = self.parse_expr()
            end = ''
            if self.peek() == ('cmd', r'\right'):
                self.take()
                end = self._read_delim()
            return ('delim', beg, body, end)
        if name in _TEXT_STYLE_COMMANDS:
            style = _TEXT_STYLE_COMMANDS[name]
            return ('text', _plain_text(self.parse_group()), style)
        if name in _LATEX_ACCENTS:
            return ('acc', _LATEX_ACCENTS[name], self.parse_group())
        if name in (r'\overline', r'\underline'):
            pos = 'top' if name == r'\overline' else 'bot'
            return ('bar', pos, self.parse_group())
        if name in (r'\underbrace', r'\overbrace'):
            chr_ = '⏟' if name == r'\underbrace' else '⏞'
            body = self.parse_group()
            sub = None
            if self.peek()[0] == 'sub':
                self.take()
                sub = self.parse_group()
            return ('groupchr', chr_, body, sub)
        if name == r'\begin':
            return self.parse_env()
        if name in _LATEX_NARY:
            sub = sup = None
            while True:
                kind, val = self.peek()
                if kind == 'sub':
                    self.take()
                    sub = self.parse_group()
                elif kind == 'sup':
                    self.take()
                    sup = self.parse_group()
                else:
                    break
            return ('nary', _LATEX_NARY[name], sub, sup)
        if name == r'\limits':
            return None
        if name in _LATEX_SYMBOLS:
            value = _LATEX_SYMBOLS[name]
            if value.strip():
                self._drop_space = True
            return ('text', value, None)
        # 未知命令：保留字母本身，尽量不丢内容
        return ('text', name[1:], None)

    def _read_delim(self):
        kind, val = self.take()
        if kind in ('other', 'esc', 'lbrack', 'rbrack'):
            return _LATEX_DELIMS.get(val, val)
        if kind == 'cmd':
            return _LATEX_DELIMS.get(val, val[1:])
        return ''

    def _read_env_name(self):
        """读取 {name} 形式的 begin/end 参数名"""
        if self.peek()[0] == 'open':
            self.take()
            kind, val = self.take()
            if self.peek()[0] == 'close':
                self.take()
            return val if kind == 'word' else ''
        kind, val = self.take()
        return val if kind == 'word' else ''

    def parse_env(self):
        env = self._read_env_name()
        rows, cur = [], []
        while True:
            kind, val = self.peek()
            if kind is None:
                break
            if kind == 'cmd' and val == r'\end':
                self.take()
                self._read_env_name()
                break
            if kind == 'close':
                self.take()
                continue
            if kind == 'amp':
                self.take()
                continue
            if kind == 'esc' and val == '\\\\':
                self.take()
                rows.append(cur)
                cur = []
                continue
            node = self.parse_scripted()
            if node is not None:
                cur.append(node)
        rows.append(cur)
        return ('env', env, rows)


def _plain_text(nodes):
    """取出节点里的纯文本（用于 \\text / \\mathrm）"""
    out = []
    for node in nodes or []:
        if node is None:
            continue
        kind = node[0]
        if kind == 'text':
            out.append(node[1])
        elif kind == 'group':
            out.append(_plain_text(node[1]))
        elif kind == 'sub':
            out.append(_plain_text(node[1]))
        elif kind == 'sup':
            out.append(_plain_text(node[1]))
        elif kind == 'subsup':
            out.append(_plain_text(node[1]))
        elif kind == 'frac':
            out.append(_plain_text(node[1]))
            out.append(_plain_text(node[2]))
    return ''.join(out)


def _m_run(text, style=None):
    """<m:r>：公式里的文字"""
    from lxml import etree
    run = etree.Element(_mtag('r'))
    if style:
        rpr = etree.SubElement(run, _mtag('rPr'))
        if style == 'nor':
            etree.SubElement(rpr, _mtag('nor'))
        elif style == 'plain':
            sty = etree.SubElement(rpr, _mtag('sty'))
            sty.set(_mtag('val'), 'p')
        elif style == 'bold':
            sty = etree.SubElement(rpr, _mtag('sty'))
            sty.set(_mtag('val'), 'b')
    node = etree.SubElement(run, _mtag('t'))
    node.text = text
    return run


def _m_children(nodes):
    """把节点列表转成 OMML 元素列表"""
    out = []
    for node in nodes or []:
        if node is not None:
            out.extend(_m_node(node))
    return out


def _m_arg(tag, nodes):
    """<m:sub> / <m:sup> / <m:num> / <m:den> / <m:e> 等容器"""
    from lxml import etree
    arg = etree.Element(_mtag(tag))
    for child in _m_children(nodes):
        arg.append(child)
    return arg


def _m_wrap(tag, *args):
    """<m:f> / <m:sSub> … 由若干容器拼成的元素"""
    from lxml import etree
    el = etree.Element(_mtag(tag))
    for arg in args:
        el.append(arg)
    return el


def _m_node(node):
    from lxml import etree
    kind = node[0]

    if kind == 'text':
        return [_m_run(node[1], node[2])] if node[1] else []
    if kind == 'group':
        return _m_children(node[1])

    if kind == 'sub':
        return [_m_wrap('sSub', _m_arg('e', [node[1]]), _m_arg('sub', node[2]))]
    if kind == 'sup':
        return [_m_wrap('sSup', _m_arg('e', [node[1]]), _m_arg('sup', node[2]))]
    if kind == 'subsup':
        return [_m_wrap('sSubSup', _m_arg('e', [node[1]]),
                        _m_arg('sub', node[2]), _m_arg('sup', node[3]))]

    if kind == 'frac':
        return [_m_wrap('f', _m_arg('num', node[1]), _m_arg('den', node[2]))]

    if kind == 'rad':
        pr = etree.Element(_mtag('radPr'))
        hide = etree.SubElement(pr, _mtag('degHide'))
        hide.set(_mtag('val'), '0' if node[1] else '1')
        return [_m_wrap('rad', pr, _m_arg('deg', node[1]), _m_arg('e', node[2]))]

    if kind == 'delim':
        pr = etree.Element(_mtag('dPr'))
        beg = etree.SubElement(pr, _mtag('begChr'))
        beg.set(_mtag('val'), node[1])
        end = etree.SubElement(pr, _mtag('endChr'))
        end.set(_mtag('val'), node[3])
        return [_m_wrap('d', pr, _m_arg('e', node[2]))]

    if kind == 'nary':
        pr = etree.Element(_mtag('naryPr'))
        chr_ = etree.SubElement(pr, _mtag('chr'))
        chr_.set(_mtag('val'), node[1])
        lim = etree.SubElement(pr, _mtag('limLoc'))
        lim.set(_mtag('val'), 'undOvr')
        return [_m_wrap('nary', pr, _m_arg('sub', node[2]),
                        _m_arg('sup', node[3]), _m_arg('e', []))]

    if kind == 'acc':
        pr = etree.Element(_mtag('accPr'))
        chr_ = etree.SubElement(pr, _mtag('chr'))
        chr_.set(_mtag('val'), node[1])
        return [_m_wrap('acc', pr, _m_arg('e', node[2]))]

    if kind == 'bar':
        pr = etree.Element(_mtag('barPr'))
        pos = etree.SubElement(pr, _mtag('pos'))
        pos.set(_mtag('val'), node[1])
        return [_m_wrap('bar', pr, _m_arg('e', node[2]))]

    if kind == 'groupchr':
        pr = etree.Element(_mtag('groupChrPr'))
        chr_ = etree.SubElement(pr, _mtag('chr'))
        chr_.set(_mtag('val'), node[1])
        pos = etree.SubElement(pr, _mtag('pos'))
        pos.set(_mtag('val'), 'bot' if node[1] == '⏟' else 'top')
        return [_m_wrap('groupChr', pr, _m_arg('e', node[2]),
                        _m_arg('sub', node[3] or []))]

    if kind == 'env':
        return _m_env(node[1], node[2])

    return []


def _m_env(env, rows):
    """matrix / bmatrix / aligned / array 等环境"""
    from lxml import etree
    if (env or '').lower().startswith(('align', 'array', 'cases')):
        arr = etree.Element(_mtag('eqArr'))
        for row in rows:
            arr.append(_m_arg('e', row))
        return [arr]
    matrix = etree.Element(_mtag('m'))
    for row in rows:
        mr = etree.SubElement(matrix, _mtag('mr'))
        for cell in row:
            mr.append(_m_arg('e', cell))
    return [matrix]


def latex_to_omml(latex):
    """LaTeX 片段 → <m:oMath> 元素；无法解析时返回 None"""
    from lxml import etree
    latex = (latex or '').strip()
    if not latex:
        return None
    try:
        nodes = _LatexParser(latex).parse()
        if not nodes:
            return None
        omath = etree.Element(_mtag('oMath'))
        for child in _m_children(nodes):
            omath.append(child)
        return omath
    except Exception:
        return None


def add_inline_math(para, latex):
    """把行内公式写成 Word 公式；失败则退回原始 LaTeX 文本"""
    omath = latex_to_omml(latex)
    if omath is None:
        para.add_run(f'${latex}$')
        return False
    para._p.append(omath)
    return True


def add_display_math(para, latex):
    """把块级公式写成独立成行的 Word 公式"""
    from lxml import etree
    body = latex_to_omml(latex)
    if body is None:
        para.add_run(f'$${latex}$$')
        return False
    wrapper = etree.Element(_mtag('oMathPara'))
    wrapper.append(body)
    para._p.append(wrapper)
    return True


def _find_list_number_abstract_id(num_part):
    """找到 List Number 样式所引用的 abstractNumId"""
    w = qn('w:abstractNum')
    for abstract in num_part.element.findall(w):
        for lvl in abstract.findall(qn('w:lvl')):
            pstyle_elem = lvl.find(qn('w:pStyle'))
            if pstyle_elem is not None and pstyle_elem.get(qn('w:val')) == 'ListNumber':
                return abstract.get(qn('w:abstractNumId'))
    return None


def _add_num_override(num_part, num_id):
    """新建一个 w:num，用 lvlOverride 让有序列表从 1 重新编号"""
    from lxml import etree
    abstract_id = _find_list_number_abstract_id(num_part)
    if abstract_id is None:
        return None
    num_elem = etree.SubElement(num_part.element, qn('w:num'))
    num_elem.set(qn('w:numId'), str(num_id))
    abstract_ref = etree.SubElement(num_elem, qn('w:abstractNumId'))
    abstract_ref.set(qn('w:val'), abstract_id)
    override = etree.SubElement(num_elem, qn('w:lvlOverride'))
    override.set(qn('w:ilvl'), '0')
    start_override = etree.SubElement(override, qn('w:startOverride'))
    start_override.set(qn('w:val'), '1')
    return num_id


def _cancel_style_color_overrides(doc):
    """去掉 python-docx 内置样式里写死的颜色，让样式跟随主题"""
    for s in doc.styles:
        if s.type != 1:
            continue
        rpr = s.element.find(qn('w:rPr'))
        if rpr is None:
            continue
        color = rpr.find(qn('w:color'))
        if color is not None:
            rpr.remove(color)


def convert(md_path):
    """将单个 markdown 文件转换为同目录同名的 docx"""
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    doc = Document()
    _link_default_fonts_to_theme(doc)
    _cancel_style_color_overrides(doc)

    in_code_block = False
    code_content = []
    blockquote_lines = []
    in_blockquote = False
    list_level = 0
    ol_next_num_id = 10
    ol_group_active = False
    ol_current_num_id = None

    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\n')

        if in_code_block:
            if line.strip().startswith('```'):
                para = doc.add_paragraph()
                para.add_run('\n'.join(code_content))
                in_code_block = False
                code_content = []
                ol_group_active = False
            else:
                code_content.append(line)
            i += 1
            continue

        if line.strip().startswith('```'):
            in_code_block = True
            code_content = []
            i += 1
            continue

        if not line.strip():
            if in_blockquote:
                para = doc.add_paragraph()
                _add_formatted_text(para, ' '.join(blockquote_lines))
                in_blockquote = False
                blockquote_lines = []
            while list_level > 0:
                list_level -= 1
            ol_group_active = False
            i += 1
            continue

        if re.match(r'^[\-\*_]{3,}\s*$', line.strip()):
            para = doc.add_paragraph()
            para.add_run('_' * 50)
            ol_group_active = False
            i += 1
            continue

        heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            para = doc.add_heading(level=level)
            _add_formatted_text(para, text)
            list_level = 0
            ol_group_active = False
            i += 1
            continue

        # 块级公式 $$...$$：整行就是一条公式
        display_match = re.match(r'^\$\$(.+)\$\$\s*$', line.strip())
        if display_match:
            para = doc.add_paragraph()
            add_display_math(para, display_match.group(1))
            list_level = 0
            ol_group_active = False
            i += 1
            continue

        bq_match = re.match(r'^>\s?(.*)$', line)
        if bq_match:
            in_blockquote = True
            blockquote_lines.append(bq_match.group(1))
            ol_group_active = False
            i += 1
            continue

        ul_match = re.match(r'^(\s*)([-*+])\s+(.+)$', line)
        if ul_match:
            indent = len(ul_match.group(1))
            content = ul_match.group(3)
            new_level = indent // 2 + 1

            while list_level > new_level:
                list_level -= 1
            while list_level < new_level:
                list_level += 1

            para = doc.add_paragraph(style='List Bullet')
            _add_formatted_text(para, content)
            ol_group_active = False
            i += 1
            continue

        ol_match = re.match(r'^(\s*)\d+[.)]\s+(.+)$', line)
        if ol_match:
            indent = len(ol_match.group(1))
            content = ol_match.group(2)
            new_level = indent // 2 + 1

            while list_level > new_level:
                list_level -= 1
            while list_level < new_level:
                list_level += 1

            if not ol_group_active:
                _add_num_override(doc.part.numbering_part, ol_next_num_id)
                ol_current_num_id = ol_next_num_id
                ol_next_num_id += 1
                ol_group_active = True

            para = doc.add_paragraph(style='List Number')
            pPr = para._p.get_or_add_pPr()
            numPr = OxmlElement('w:numPr')
            ilvl = OxmlElement('w:ilvl')
            ilvl.set(qn('w:val'), '0')
            numId_elem = OxmlElement('w:numId')
            numId_elem.set(qn('w:val'), str(ol_current_num_id))
            numPr.append(ilvl)
            numPr.append(numId_elem)
            pPr.insert(0, numPr)
            _add_formatted_text(para, content)
            i += 1
            continue

        if '|' in line and i + 1 < len(lines) and re.match(r'^[\s|:\-]+$', lines[i + 1]):
            table_lines = []
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1

            if len(table_lines) >= 2:
                headers = _split_table_row(table_lines[0])
                start_row = 2 if re.match(r'^[\s|:\-]+$', table_lines[1]) else 1
                rows = []
                for tl in table_lines[start_row:]:
                    rows.append(_split_table_row(tl))

                num_cols = len(headers)
                table = doc.add_table(rows=1 + len(rows), cols=num_cols)
                table.style = 'Table Grid'

                for j, h in enumerate(headers):
                    if j < num_cols:
                        cell = table.rows[0].cells[j]
                        _add_formatted_text(cell.paragraphs[0], h)
                        for run in cell.paragraphs[0].runs:
                            run.bold = True

                for r_idx, row in enumerate(rows):
                    for c_idx, cell_text in enumerate(row):
                        if c_idx < num_cols:
                            cell = table.rows[r_idx + 1].cells[c_idx]
                            _add_formatted_text(cell.paragraphs[0], cell_text)

            list_level = 0
            ol_group_active = False
            continue

        if in_blockquote:
            para = doc.add_paragraph()
            _add_formatted_text(para, ' '.join(blockquote_lines))
            in_blockquote = False
            blockquote_lines = []

        para = doc.add_paragraph()
        _add_formatted_text(para, line.strip())
        ol_group_active = False
        i += 1

    if in_blockquote and blockquote_lines:
        para = doc.add_paragraph()
        _add_formatted_text(para, ' '.join(blockquote_lines))

    while list_level > 0:
        list_level -= 1

    out_path = os.path.splitext(md_path)[0] + '.docx'
    doc.save(out_path)
    print(f'Converted: {md_path} -> {out_path}')


def convert_dir(directory):
    """转换指定目录下的所有 .md 文件"""
    md_files = sorted(glob.glob(os.path.join(directory, '*.md')))
    if not md_files:
        print(f'No .md files found in: {directory}')
        return
    for md_path in md_files:
        convert(md_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert Markdown files to Word (.docx) documents.')
    parser.add_argument('path', nargs='?',
                        help='Path to a .md file, or with -d/--dir a directory')
    parser.add_argument('-d', '--dir', action='store_true',
                        help='Treat the path as a directory and convert all .md files in it')
    args = parser.parse_args()

    if not args.path:
        parser.print_help()
        sys.exit(1)

    if args.dir:
        if not os.path.isdir(args.path):
            print(f'Error: Directory not found: {args.path}')
            sys.exit(1)
        convert_dir(args.path)
    else:
        if not os.path.isfile(args.path):
            print(f'Error: File not found: {args.path}')
            sys.exit(1)
        convert(args.path)
