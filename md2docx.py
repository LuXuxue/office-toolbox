#!/usr/bin/env python3
"""Convert Markdown files to Word (.docx) documents."""

import sys
import re
import os
import argparse
import glob
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.oxml import OxmlElement


def _add_formatted_text(para, text):
    """Add text with inline markdown formatting to a paragraph."""
    if not text:
        return

    segments = re.split(r'(`[^`]+`)', text)

    for segment in segments:
        if segment.startswith('`') and segment.endswith('`'):
            para.add_run(segment[1:-1])
        elif segment:
            _add_formatted_runs(para, segment)


def _add_formatted_runs(para, text):
    """Add bold, italic, bold-italic, and link formatting to a paragraph."""
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
    """Link default fonts to the theme and match Word 2016's default document."""
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
    """Make the bundled theme a Simplified-Chinese one (等线 fonts)."""
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
    """Set latin and the Chinese (Hans) script font in a fontScheme slot."""
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
    """Set settings.xml themeFontLang eastAsia to zh-CN."""
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


def _find_list_number_abstract_id(num_part):
    """Return the abstractNumId for the List Number style."""
    w = qn('w:abstractNum')
    for abstract in num_part.element.findall(w):
        for lvl in abstract.findall(qn('w:lvl')):
            pstyle_elem = lvl.find(qn('w:pStyle'))
            if pstyle_elem is not None and pstyle_elem.get(qn('w:val')) == 'ListNumber':
                return abstract.get(qn('w:abstractNumId'))
    return None


def _add_num_override(num_part, num_id):
    """Add a new <w:num> entry with a lvlOverride that restarts numbering."""
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
    """Remove explicit colour overrides from python-docx's built-in styles."""
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
    """Convert a markdown file to docx."""
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
                headers = [c.strip() for c in table_lines[0].strip('|').split('|')]
                start_row = 2 if re.match(r'^[\s|:\-]+$', table_lines[1]) else 1
                rows = []
                for tl in table_lines[start_row:]:
                    rows.append([c.strip() for c in tl.strip('|').split('|')])

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
    """Convert all .md files under a directory to .docx files."""
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
