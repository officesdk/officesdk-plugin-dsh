#!/usr/bin/env python3
"""Build the empty templates document_create starts from (shimo-docx live mode).

Usage: python build_live_templates.py [--out <templates dir>]

gongwen : 公文版式——A4，上 37、下 35、左 28、右 26 mm，页码「— 1 —」奇右偶左；正文三号仿宋固定行距 29pt 首行缩进 2 字；
          标题二号小标宋居中；一级黑体、二级楷体、三级仿宋加粗，均 16pt。与离线脚本 plain 模式同一套参数（assets/style.json）。
redhead : 红头公文——与 gongwen 同一套页面、样式与页码，正文前预置两段占位：发文机关标志（「Shimo Issuer」样式：小标宋 42pt 红色居中）
          与发文字号（「Shimo Doc Number」样式：仿宋三号居中，下方 1.5pt 红线）。格式都在段落样式里，模型用 setText 换掉占位文字后
          红字、字号、红线照旧；没有空段，第一个插入的段落接在红线下面。不加离线版的「草稿 · 待核定」页眉。
basic   : 普通 A4 文档——页边距 25.4 / 31.8 mm，正文宋体小四 1.5 倍行距，标题与各级标题黑体，页脚居中页码「— 1 —」。

The Writer JS API applies paragraph styles by style id; python-docx's own template names built-in styles with English ids
(Normal, Title, Heading1…), which is what the live agent's `style` field maps to. gongwen and basic have one empty paragraph; redhead has its two placeholder paragraphs and no empty one.
"""
import argparse
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gongwen_docx as gw  # noqa: E402  (same page and type settings as the offline generator)

T = gw.T
STYLES = ('Normal', 'Title', 'Subtitle', 'Heading 1', 'Heading 2', 'Heading 3', 'Quote')


def font(style, family, size, bold=False, color='000000'):
    style.font.name = family
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    style.font.color.rgb = RGBColor.from_string(color)
    rfonts = style.element.get_or_add_rPr().get_or_add_rFonts()
    for key in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(key), family)
    for key in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        rfonts.attrib.pop(qn(key), None)


def paragraph(style, align=None, line=None, multiple=None, indent=0.0, before=0, after=0, outline=None):
    f = style.paragraph_format
    style.paragraph_format.alignment = align
    f.space_before, f.space_after = Pt(before), Pt(after)
    if line is not None:
        f.line_spacing, f.line_spacing_rule = Pt(line), WD_LINE_SPACING.EXACTLY
    elif multiple is not None:
        f.line_spacing = multiple
    f.first_line_indent = Pt(indent) if indent else None
    f.keep_with_next = outline is not None
    # No outline level, not even python-docx's own on its headings: office-kit's Kit cannot save a document once the
    # Writer API applied a style (or a paragraph setting) with one (it answers the save with a failure).
    ppr = style.element.get_or_add_pPr()
    for old in ppr.findall(qn('w:outlineLvl')):
        ppr.remove(old)


def strip_outline(d):
    for style in d.styles:
        ppr = style.element.pPr
        if ppr is not None:
            for old in ppr.findall(qn('w:outlineLvl')):
                ppr.remove(old)


def no_title_rule(style):
    # python-docx's Title has a bottom border; neither layout wants a rule under the title.
    ppr = style.element.get_or_add_pPr()
    for old in ppr.findall(qn('w:pBdr')):
        ppr.remove(old)


def mark(d):
    # The live agent applies the title and heading styles only in a document that has this style (its own templates,
    # whose headings carry no outline level); elsewhere it formats them directly.
    d.styles.add_style('Shimo Live', WD_STYLE_TYPE.PARAGRAPH).hidden = True


def clear_body(d, empty=True):
    body = d.element.body
    for child in list(body):
        if child.tag != qn('w:sectPr'):
            body.remove(child)
    if empty:
        d.add_paragraph()


def bottom_rule(style, color, eighth_points, space):
    # The red line under the document number lives in the paragraph style, so it survives the Writer API replacing the
    # paragraph's text and is not copied onto paragraphs inserted after it (they get another style).
    ppr = style.element.get_or_add_pPr()
    for old in ppr.findall(qn('w:pBdr')):
        ppr.remove(old)
    bdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for key, value in {'val': 'single', 'sz': str(eighth_points), 'space': str(space), 'color': color}.items():
        bottom.set(qn('w:' + key), value)
    bdr.append(bottom)
    # Schema order (CT_PPrBase): pBdr comes after keepNext…numPr and before shd, tabs, spacing, ind, jc….
    later = ('shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
             'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId',
             'cnfStyle', 'rPr', 'sectPr', 'pPrChange')
    anchor = next((child for child in ppr if child.tag in {qn('w:' + tag) for tag in later}), None)
    if anchor is None:
        ppr.append(bdr)
    else:
        anchor.addprevious(bdr)


def page_field(p, family, size):
    """「— PAGE —」 in one paragraph (the field Word and the editor fill with the page number)."""
    def piece(text=''):
        r = p.add_run(text)
        r.font.name, r.font.size = family, Pt(size)
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), family)
        return r
    piece('— ')
    r = piece()
    for tag, attrs, text in [('fldChar', {'fldCharType': 'begin'}, None), ('instrText', {}, ' PAGE '),
                             ('fldChar', {'fldCharType': 'separate'}, None), ('t', {}, '1'),
                             ('fldChar', {'fldCharType': 'end'}, None)]:
        el = OxmlElement('w:' + tag)
        for key, value in attrs.items():
            el.set(qn('w:' + key), value)
        if tag == 'instrText':
            el.set(qn('xml:space'), 'preserve')
        if text is not None:
            el.text = text
        r._element.append(el)
    piece(' —')


def gongwen():
    d = Document()
    fonts = dict(gw.STYLE['fonts'])
    gw.setup(d, fonts, False, None)
    s = d.styles
    body = T['body']
    font(s['Normal'], fonts['body'], body)
    paragraph(s['Normal'], WD_ALIGN_PARAGRAPH.JUSTIFY, line=T['body_line'], indent=2 * body)
    font(s['Title'], fonts['title'], T['title'])
    no_title_rule(s['Title'])
    paragraph(s['Title'], WD_ALIGN_PARAGRAPH.CENTER, line=T['title_line'], after=T['body_line'], outline=0)
    font(s['Subtitle'], fonts['secondary'], T['subtitle'])
    paragraph(s['Subtitle'], WD_ALIGN_PARAGRAPH.CENTER, line=T['body_line'])
    font(s['Heading 1'], fonts['heading'], body)
    paragraph(s['Heading 1'], WD_ALIGN_PARAGRAPH.JUSTIFY, line=T['body_line'], indent=2 * body, outline=0)
    font(s['Heading 2'], fonts['secondary'], body)
    paragraph(s['Heading 2'], WD_ALIGN_PARAGRAPH.JUSTIFY, line=T['body_line'], indent=2 * body, outline=1)
    font(s['Heading 3'], fonts['body'], body, bold=True)
    paragraph(s['Heading 3'], WD_ALIGN_PARAGRAPH.JUSTIFY, line=T['body_line'], indent=2 * body, outline=2)
    font(s['Quote'], fonts['secondary'], body)
    paragraph(s['Quote'], WD_ALIGN_PARAGRAPH.JUSTIFY, line=T['body_line'], indent=2 * body)
    strip_outline(d)
    mark(d)
    clear_body(d)
    return d


def redhead():
    d = gongwen()
    fonts = dict(gw.STYLE['fonts'])
    red = gw.STYLE['red_line']
    s = d.styles
    issuer = s.add_style('Shimo Issuer', WD_STYLE_TYPE.PARAGRAPH)
    issuer.base_style = s['Normal']
    font(issuer, fonts['title'], T['redhead'], color=red['color'])
    paragraph(issuer, WD_ALIGN_PARAGRAPH.CENTER, line=T['redhead_line'], before=98, after=2 * T['body_line'] - 16)
    number = s.add_style('Shimo Doc Number', WD_STYLE_TYPE.PARAGRAPH)
    number.base_style = s['Normal']
    font(number, fonts['body'], T['body'])
    paragraph(number, WD_ALIGN_PARAGRAPH.CENTER, line=T['body_line'], after=2 * T['body_line'])
    bottom_rule(number, red['color'], red['eighth_points'], 11)
    for style in (issuer, number):
        # Normal's first-line indent would push the centred text off centre.
        style.paragraph_format.first_line_indent = Pt(0)
    clear_body(d, empty=False)
    d.add_paragraph('【发文机关】文件', style=issuer)
    d.add_paragraph('【发文字号】', style=number)
    return d


def basic():
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin = sec.bottom_margin = Mm(25.4)
    sec.left_margin = sec.right_margin = Mm(31.8)
    s = d.styles
    font(s['Normal'], 'SimSun', 12)
    paragraph(s['Normal'], WD_ALIGN_PARAGRAPH.JUSTIFY, multiple=1.5)
    font(s['Title'], 'SimHei', 22, bold=True)
    no_title_rule(s['Title'])
    paragraph(s['Title'], WD_ALIGN_PARAGRAPH.CENTER, multiple=1.2, before=12, after=12, outline=0)
    font(s['Subtitle'], 'KaiTi', 14)
    paragraph(s['Subtitle'], WD_ALIGN_PARAGRAPH.CENTER, multiple=1.5, after=12)
    for name, size, level in (('Heading 1', 16, 0), ('Heading 2', 14, 1), ('Heading 3', 12, 2)):
        font(s[name], 'SimHei', size, bold=True)
        paragraph(s[name], WD_ALIGN_PARAGRAPH.LEFT, multiple=1.5, before=12 if level == 0 else 6, after=6, outline=level)
    font(s['Quote'], 'KaiTi', 12, color='555555')
    paragraph(s['Quote'], WD_ALIGN_PARAGRAPH.LEFT, multiple=1.5)
    sec.footer_distance = Mm(17.5)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    page_field(footer, 'SimSun', 10.5)
    strip_outline(d)
    mark(d)
    clear_body(d)
    return d


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('--out', default=str(Path(__file__).resolve().parents[1] / 'templates'))
    out = Path(parser.parse_args().out)
    out.mkdir(parents=True, exist_ok=True)
    for name, build in (('gongwen', gongwen), ('redhead', redhead), ('basic', basic)):
        target = out / f'{name}.docx'
        build().save(target)
        print(target)


if __name__ == '__main__':
    main()
