#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""normalize_docx — 将 .docx 的格式标准化

用法:
    python normalize_docx.py path/to/file.docx

标准化内容:
    - 主题字体: 标题中文黑体 / 英文 Times New Roman，正文中文宋体 / 英文 Times New Roman
    - 正文: 小四字号、1.5 倍行距
    - 标题 1: 小三字号、1.5 倍行距、段前分页、加粗、段前段后间距为 0
    - 标题 2~9: 小四字号、1.5 倍行距、加粗、段前段后间距为 0
    - 标题 1~9: 多级列表自动编号，1~7 级章节号（1 / 1.1 / 1.1.1…），
      8 级图号（图 1.1-1）、9 级表号（表 1.1-1），与参考模板一致
    - 自动替换: 常见计量单位写法规范化（如 m3/h → m³/h），替换表见 _REPLACEMENTS

结果直接覆盖原 .docx 文件。

依赖:
    pip install python-docx
"""

import sys
import argparse
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 字号（w:sz 以半磅为单位）：小四 = 12pt = 24，小三 = 15pt = 30，五号 = 10.5pt = 21
_SZ_XIAO4 = '24'
_SZ_XIAO3 = '30'
_SZ_WUHAO = '21'

# 行距（w:line 以 1/20 磅为单位）：单倍 = 240，1.5 倍 = 360
_LINE_1 = '240'
_LINE_1_5 = '360'

# 主题字体方案（中文环境：major 标题 / minor 正文，各分拉丁与中文）
# 标题：中文黑体、英文 Times New Roman；正文：中文宋体、英文 Times New Roman
_MAJOR_LATIN = 'Times New Roman'
_MAJOR_EA = '黑体'
_MINOR_LATIN = 'Times New Roman'
_MINOR_EA = '宋体'

# 自动替换表：(原文, 替换文)，按顺序应用
# 后续新增替换规则只需在此列表追加
_REPLACEMENTS = [
    ('m3/', 'm³/'),
    ('m2/', 'm²/'),
    ('/m3', '/m³'),
    ('/m2', '/m²'),
    ('km2', 'km²'),
    ('BOD5', 'BOD₅'),
    ('NH3', 'NH₃'),
    ('H2S', 'H₂S'),
    ('CO2', 'CO₂'),
]

# 无间隔样式名（Word 内置，styleId=NoSpacing）
_STYLE_NO_SPACING = 'No Spacing'


def _find_style(doc, names, style_id):
    """按名称（本地化兜底）或 styleId 找样式"""
    for name in names:
        try:
            return doc.styles[name]
        except KeyError:
            continue
    for style in doc.styles:
        if style.style_id == style_id:
            return style
    return None


def _set_theme_rfonts(rpr, major):
    """让 rPr 的字体跟随主题（标题用 major，正文用 minor）"""
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    prefix = 'major' if major else 'minor'
    rfonts.set(qn('w:asciiTheme'), prefix + 'HAnsi')
    rfonts.set(qn('w:eastAsiaTheme'), prefix + 'EastAsia')
    rfonts.set(qn('w:hAnsiTheme'), prefix + 'HAnsi')
    rfonts.set(qn('w:cstheme'), prefix + 'Bidi')
    # 清掉可能硬编码的字体名，避免盖住主题引用
    for attr in ('w:ascii', 'w:eastAsia', 'w:hAnsi', 'w:cs'):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def _get_style_rpr(style):
    """取样式的 w:rPr，没有就建一个"""
    rpr = style.element.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        # rPr 是 w:style 的最后一个子元素
        style.element.append(rpr)
    return rpr


def _set_style_format(style, sz, line, page_break=False, major=False,
                      bold=False, zero_spacing=False, no_italic=False,
                      center=False):
    """设置样式的字号、行距、段前分页、加粗、斜体、居中、段前段后间距与主题字体"""
    rpr = _get_style_rpr(style)
    _set_theme_rfonts(rpr, major)

    # 加粗（w:b 无 val 属性即为 true）
    if bold:
        b = rpr.find(qn('w:b'))
        if b is None:
            b = OxmlElement('w:b')
            rpr.insert(0, b)

    # 取消倾斜（w:i val=0）
    if no_italic:
        i_el = rpr.find(qn('w:i'))
        if i_el is None:
            i_el = OxmlElement('w:i')
            rpr.append(i_el)
        i_el.set(qn('w:val'), '0')

    # 字号
    sz_el = rpr.find(qn('w:sz'))
    if sz_el is None:
        sz_el = OxmlElement('w:sz')
        rpr.append(sz_el)
    sz_el.set(qn('w:val'), sz)
    sz_cs = rpr.find(qn('w:szCs'))
    if sz_cs is None:
        sz_cs = OxmlElement('w:szCs')
        rpr.append(sz_cs)
    sz_cs.set(qn('w:val'), sz)

    ppr = style.element.get_or_add_pPr()
    # 行距与段前段后间距
    spacing = ppr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        ppr.append(spacing)
    spacing.set(qn('w:line'), line)
    spacing.set(qn('w:lineRule'), 'auto')
    if zero_spacing:
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')

    # 居中对齐（center=True 设 center，center=False 移除 jc）
    if center:
        jc = ppr.find(qn('w:jc'))
        if jc is None:
            jc = OxmlElement('w:jc')
            ppr.append(jc)
        jc.set(qn('w:val'), 'center')
    else:
        jc = ppr.find(qn('w:jc'))
        if jc is not None:
            ppr.remove(jc)

    # 段前分页
    if page_break:
        pbb = ppr.find(qn('w:pageBreakBefore'))
        if pbb is None:
            pbb = OxmlElement('w:pageBreakBefore')
            spacing.addprevious(pbb)
    else:
        pbb = ppr.find(qn('w:pageBreakBefore'))
        if pbb is not None:
            ppr.remove(pbb)


def _set_theme_fonts(node, latin_face, ea_face):
    """设置 fontScheme 中拉丁字体与中文（Hans）字体的字体名"""
    from lxml import etree
    latin = node.find(qn('a:latin'))
    if latin is None:
        latin = etree.SubElement(node, qn('a:latin'))
    latin.set('typeface', latin_face)
    ea = node.find(qn('a:font'))
    if ea is None:
        ea = etree.SubElement(node, qn('a:font'))
    ea.set('script', 'Hans')
    ea.set('typeface', ea_face)


def _config_theme(doc):
    """把主题字体方案设为：标题=黑体/Calibri，正文=宋体/Cambria"""
    from lxml import etree
    theme_part = None
    for rel in doc.part.rels.values():
        if 'theme' in rel.reltype:
            theme_part = rel.target_part
            break
    if theme_part is None:
        return

    root = etree.fromstring(theme_part.blob)
    font_scheme = root.find('.//' + qn('a:fontScheme'))
    if font_scheme is None:
        return
    major = font_scheme.find(qn('a:majorFont'))
    minor = font_scheme.find(qn('a:minorFont'))
    if major is not None:
        _set_theme_fonts(major, _MAJOR_LATIN, _MAJOR_EA)
    if minor is not None:
        _set_theme_fonts(minor, _MINOR_LATIN, _MINOR_EA)

    new_blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                             standalone=True)
    theme_part._blob = new_blob


def _link_default_fonts_to_theme(doc):
    """让默认字体跟随主题，默认字号小四"""
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

    _set_theme_rfonts(rpr, major=False)

    sz = rpr.find(qn('w:sz'))
    if sz is None:
        sz = OxmlElement('w:sz')
        rpr.append(sz)
    sz.set(qn('w:val'), _SZ_XIAO4)
    sz_cs = rpr.find(qn('w:szCs'))
    if sz_cs is None:
        sz_cs = OxmlElement('w:szCs')
        rpr.append(sz_cs)
    sz_cs.set(qn('w:val'), _SZ_XIAO4)

    # 默认段距也统一为 1.5 倍行距（Word 新建段落继承这个值）
    ppr_default = doc_defaults.find(qn('w:pPrDefault'))
    if ppr_default is None:
        ppr_default = OxmlElement('w:pPrDefault')
        doc_defaults.append(ppr_default)
    ppr = ppr_default.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        ppr_default.append(ppr)
    spacing = ppr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        ppr.append(spacing)
    spacing.set(qn('w:line'), _LINE_1_5)
    spacing.set(qn('w:lineRule'), 'auto')


def _apply_numbering(doc):
    """给标题 1~9 套用同一个多级列表

    1~5 级是章节号（1 / 1.1 / 1.1.1…），6~7 级延续章节号，8 级是图号
    （图 1.1-1），9 级是表号（表 1.1-1），图/表号随二级章节重置，与参考
    模板一致。编号写在样式上，用户在 Word 里新加的同级标题也会自动接上序号。
    """
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        print('警告: 文档没有编号定义部件，无法套用多级列表', file=sys.stderr)
        return False

    # 找 1~9 级标题样式，缺失就创建
    heading_styles = []
    for level in range(1, 10):
        style = _find_style(doc, [f'Heading {level}', f'标题 {level}',
                                 f'{level}级标题'], f'Heading{level}')
        if style is None:
            style = _create_heading_style(doc, level)
        heading_styles.append(style)

    # 找没被占用的 abstractNumId / numId
    used_abstract = {e.get(qn('w:abstractNumId'))
                     for e in numbering.findall(qn('w:abstractNum'))}
    used_num = {e.get(qn('w:numId')) for e in numbering.findall(qn('w:num'))}
    abstract_id = next((str(i) for i in range(900, 1000)
                       if str(i) not in used_abstract), None)
    num_id = next((str(i) for i in range(900, 1000)
                  if str(i) not in used_num), None)
    if abstract_id is None or num_id is None:
        print('警告: 编号定义已用满，无法套用多级列表', file=sys.stderr)
        return False

    # 构建 abstractNum（9 个级别）
    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), abstract_id)
    mlt = OxmlElement('w:multiLevelType')
    mlt.set(qn('w:val'), 'multilevel')
    abstract.append(mlt)

    for ilvl in range(9):
        level = ilvl + 1
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))

        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)

        numfmt = OxmlElement('w:numFmt')
        numfmt.set(qn('w:val'), 'decimal')
        lvl.append(numfmt)

        pstyle = OxmlElement('w:pStyle')
        pstyle.set(qn('w:val'), heading_styles[ilvl].style_id)
        lvl.append(pstyle)

        if level <= 7:
            text = '.'.join('%' + str(i + 1) for i in range(level))
        elif level == 8:
            text = '图 %1.%2-%8'
        else:
            text = '表 %1.%2-%9'
        lvltxt = OxmlElement('w:lvlText')
        lvltxt.set(qn('w:val'), text)
        lvl.append(lvltxt)

        if level in (8, 9):
            restart = OxmlElement('w:lvlRestart')
            restart.set(qn('w:val'), '2')
            lvl.append(restart)

        lvljc = OxmlElement('w:lvlJc')
        lvljc.set(qn('w:val'), 'left')
        lvl.append(lvljc)

        suff = OxmlElement('w:suff')
        suff.set(qn('w:val'), 'space')
        lvl.append(suff)

        abstract.append(lvl)

    # abstractNum 必须在所有 num 之前插入（OOXML 顺序要求）
    first_num = numbering.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)

    num = OxmlElement('w:num')
    num.set(qn('w:numId'), num_id)
    ref = OxmlElement('w:abstractNumId')
    ref.set(qn('w:val'), abstract_id)
    num.append(ref)
    numbering.append(num)

    # 每个标题样式的 numPr 指向这个 numId
    for ilvl, style in enumerate(heading_styles):
        ppr = style.element.get_or_add_pPr()
        numpr = ppr.get_or_add_numPr()
        ilvl_el = numpr.get_or_add_ilvl()
        ilvl_el.set(qn('w:val'), str(ilvl))
        numid_el = numpr.get_or_add_numId()
        numid_el.set(qn('w:val'), num_id)



    # 移除标题段落上残留的 numPr（否则段落级 numPr 会覆盖样式级）
    heading_names = {f"Heading {i}" for i in range(1, 10)}
    for para in doc.paragraphs:
        if para.style.name in heading_names:
            ppr = para._p.find(qn('w:pPr'))
            if ppr is not None:
                numpr = ppr.find(qn('w:numPr'))
                if numpr is not None:
                    ppr.remove(numpr)
    return True


def _create_heading_style(doc, level):
    """创建缺失的第 level 级标题样式"""
    from docx.enum.style import WD_STYLE_TYPE
    style = doc.styles.add_style(f'Heading {level}',
                                 WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    style.style_id = f'Heading{level}'
    olvl = OxmlElement('w:outlineLvl')
    olvl.set(qn('w:val'), str(level - 1))
    style.element.get_or_add_pPr().append(olvl)
    return style


def _apply_replacements(doc):
    """对文档所有段落文本应用 _REPLACEMENTS 中的替换规则"""
    for para in doc.paragraphs:
        for run in para.runs:
            for old, new in _REPLACEMENTS:
                if old in run.text:
                    run.text = run.text.replace(old, new)


def normalize(docx_path):
    """标准化单个 .docx 文件的格式，覆盖原文件"""
    doc = Document(docx_path)

    _link_default_fonts_to_theme(doc)
    _config_theme(doc)

    # 正文（Normal）：小四、1.5 倍行距
    normal = _find_style(doc, ['Normal', '正文'], 'Normal')
    if normal is not None:
        _set_style_format(normal, _SZ_XIAO4, _LINE_1_5, page_break=False,
                          major=False)

    # 标题 1：小三、1.5 倍行距、段前分页、加粗、取消倾斜、段前段后间距为 0
    h1 = _find_style(doc, ['Heading 1', '标题 1', '1级标题'], 'Heading1')
    if h1 is not None:
        _set_style_format(h1, _SZ_XIAO3, _LINE_1_5, page_break=True,
                          major=True, bold=True, zero_spacing=True,
                          no_italic=True)

    # 标题 2~7：小四、1.5 倍行距、加粗、取消倾斜、段前段后间距为 0
    for level in range(2, 8):
        style = _find_style(doc, [f'Heading {level}', f'标题 {level}',
                                  f'{level}级标题'], f'Heading{level}')
        if style is not None:
            _set_style_format(style, _SZ_XIAO4, _LINE_1_5, page_break=False,
                             major=True, bold=True, zero_spacing=True,
                             no_italic=True)

    # 标题 8~9：小四、1.5 倍行距、加粗、取消倾斜、居中、段前段后间距为 0
    for level in (8, 9):
        style = _find_style(doc, [f'Heading {level}', f'标题 {level}',
                                  f'{level}级标题'], f'Heading{level}')
        if style is not None:
            _set_style_format(style, _SZ_XIAO4, _LINE_1_5, page_break=False,
                             major=True, bold=True, zero_spacing=True,
                             no_italic=True, center=True)

    # 无间隔样式：五号、1 倍行距
    _set_style_format(_find_style(doc, [_STYLE_NO_SPACING], 'NoSpacing'),
                      _SZ_WUHAO, _LINE_1, page_break=False,
                      major=False, no_italic=True)

    # 标题 1~9 的多级列表（1~5 章号、8 图号、9 表号）
    _apply_numbering(doc)

    # 自动替换（计量单位规范化等）
    _apply_replacements(doc)

    doc.save(docx_path)
    print(f'已标准化: {docx_path}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        prog='normalize_docx',
        description='将 .docx 的格式标准化（主题字体、字号行距、标题多级列表）')
    parser.add_argument('path', help='.docx 文件路径')
    args = parser.parse_args()

    if not __import__('os').path.isfile(args.path):
        print(f'错误: 文件不存在: {args.path}', file=sys.stderr)
        sys.exit(1)
    normalize(args.path)
