#!/usr/bin/env python3
"""Convert an ECPC Markdown deliverable to a styled, editable Word (.docx).

Local-only (python-docx); no external API. Handles the subset of Markdown used
in the ECPC briefs: #/##/### headings, > blockquote, - bullets, 1. numbered
lists, --- rules, and inline **bold** / *italic* / `code`.

Usage:  python md_to_docx.py INPUT.md OUTPUT.docx
"""
import re, sys
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GREEN_D = RGBColor(0x1B, 0x4B, 0x3A)   # dark green (titles)
GREEN   = RGBColor(0x2E, 0x7D, 0x64)   # ECPC green (headings)
GREY    = RGBColor(0x55, 0x55, 0x55)

INLINE = re.compile(r'(\*\*[^*]+\*\*|`[^`]+`|\*[^*]+\*)')

def add_runs(p, text):
    """Add text to paragraph p, honoring **bold**, *italic*, `code`."""
    for tok in INLINE.split(text):
        if not tok:
            continue
        if tok.startswith('**') and tok.endswith('**'):
            r = p.add_run(tok[2:-2]); r.bold = True
        elif tok.startswith('`') and tok.endswith('`'):
            r = p.add_run(tok[1:-1]); r.font.name = 'Consolas'; r.font.size = Pt(9.5)
        elif tok.startswith('*') and tok.endswith('*'):
            r = p.add_run(tok[1:-1]); r.italic = True
        else:
            p.add_run(tok)

def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for k, v in (('w:val','single'),('w:sz','6'),('w:space','1'),('w:color','C9C9C9')):
        bottom.set(qn(k), v)
    pbdr.append(bottom); pPr.append(pbdr)

def left_border(p):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    for k, v in (('w:val','single'),('w:sz','18'),('w:space','12'),('w:color','2E7D64')):
        left.set(qn(k), v)
    pbdr.append(left); pPr.append(pbdr)

def style_font(doc, name, size, color, bold=True, italic=False, before=0, after=4):
    st = doc.styles[name]
    st.font.name = 'Calibri'; st.font.size = Pt(size); st.font.bold = bold
    st.font.italic = italic; st.font.color.rgb = color
    st.paragraph_format.space_before = Pt(before); st.paragraph_format.space_after = Pt(after)
    st.paragraph_format.keep_with_next = True
    return st

def build(md_path, out_path):
    lines = open(md_path, encoding='utf-8').read().split('\n')
    doc = Document()

    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'; normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.12

    # restyle Word's built-in heading styles to the ECPC palette (kept as real
    # styles so they appear in the Navigation pane and restyle globally)
    style_font(doc, 'Title',     23, GREEN_D, after=2)
    style_font(doc, 'Subtitle',  14, GREEN,  after=10, italic=False)
    style_font(doc, 'Heading 1', 15, GREEN_D, before=14, after=4)
    style_font(doc, 'Heading 2', 12, GREEN,  before=8,  after=2)

    seen_subtitle = False
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue

        if line.startswith('# '):
            add_runs(doc.add_paragraph(style='Title'), line[2:].strip())
        elif line.startswith('## '):
            if not seen_subtitle:
                seen_subtitle = True
                add_runs(doc.add_paragraph(style='Subtitle'), line[3:].strip())
            else:
                add_runs(doc.add_paragraph(style='Heading 1'), line[3:].strip())
        elif line.startswith('### '):
            add_runs(doc.add_paragraph(style='Heading 2'), line[4:].strip())
        elif line.startswith('> '):
            p = doc.add_paragraph(); left_border(p)
            p.paragraph_format.left_indent = Inches(0.18)
            p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(8)
            add_runs(p, line[2:].strip())
        elif line.strip() == '---':
            p = doc.add_paragraph(); bottom_border(p)
            p.paragraph_format.space_after = Pt(2)
        elif re.match(r'^\d+\.\s', line):
            p = doc.add_paragraph(style='List Number')
            add_runs(p, re.sub(r'^\d+\.\s', '', line))
        elif line.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            add_runs(p, line[2:].strip())
        else:
            p = doc.add_paragraph()
            # a line that is entirely *italic* reads as a caption/byline
            if line.strip().startswith('*') and line.strip().endswith('*') and '**' not in line:
                r = p.add_run(line.strip()[1:-1]); r.italic = True; r.font.color.rgb = GREY
                r.font.size = Pt(10)
            else:
                add_runs(p, line.strip())

    doc.save(out_path)
    print('Wrote', out_path)

if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2])
