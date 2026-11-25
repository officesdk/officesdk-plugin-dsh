"""Builds a style pack's live-mode template: styles/<name>/template.pptx from styles/<name>/style.json.

    <python with python-pptx> scripts/build_template.py [--style <name>]   (default: every pack under styles/)

presentation_create starts from this template: 16:9 (10 x 5.625 in, 720 x 405 pt) with no slides, so the model adds
every page through presentation_edit and the user sees each one appear. The theme takes the pack's default palette
and fonts (mapping in style.json `template.note`) and five layouts with text sizes from `template.sizes_pt`. The
editor's converter drops layout names, so each layout is told apart by the placeholders it has (what
presentation_read lists): 封面 title+subtitle, 标题和内容 title+body, 仅标题 title, 结束页 body, 空白 none. Date,
footer and slide-number placeholders are removed. Fonts are names only; the machine must have them. Re-running
produces the same parts; the zip timestamps differ, so commit the output only when the style or a layout changed.
"""
import argparse
import json
import re
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Inches, Pt
from pptx.shapes.autoshape import Shape

STYLES = Path(__file__).resolve().parents[1] / 'styles'
STYLE_NAME = re.compile(r'^[a-z0-9]+(?:-[a-z0-9]+)*$')
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'a': A}
WIDTH, HEIGHT = Inches(10), Inches(5.625)


class Style:
    """The values a template takes from style.json; hex colours without '#'."""

    def __init__(self, data):
        palette = data['palettes'][data['default_palette']]
        hexof = lambda value: value.lstrip('#').upper()
        divider = palette.get('dividerOnPrimary', palette['lineOnPrimary'])  # as brand_deck.py: only color-1 has its own
        self.primary, self.paper = hexof(palette['primary']), hexof(palette['paper'])
        self.on_primary = hexof(palette['onPrimary'])
        self.scheme = {'dk1': hexof(palette['primary']), 'lt1': self.paper, 'dk2': hexof(palette['ink']),
                       'lt2': hexof(data['template']['surface']), 'accent1': self.primary,
                       'accent2': hexof(palette['secondary']), 'accent3': hexof(palette['chartTertiary']),
                       'accent4': hexof(palette['lineOnPrimary']), 'accent5': hexof(palette['lineOnPaper']),
                       'accent6': hexof(divider), 'hlink': hexof(palette['ink']), 'folHlink': hexof(divider)}
        self.latin, self.cjk = data['fonts']['latin'], data['fonts']['cjk']
        self.sizes = data['template']['sizes_pt']
        # Optional template.colors (palette role names) for styles the default rules do not suit, e.g. dark ones whose
        # primary is black: cover_bg / closing_bg / content_bg, cover_text, title, body. Absent keys keep the rules below.
        colors = data['template'].get('colors') or {}
        role = lambda key, fallback: hexof(palette[colors[key]]) if key in colors else fallback
        self.cover_bg = role('cover_bg', self.primary)
        self.closing_bg = role('closing_bg', self.cover_bg)
        self.cover_text = role('cover_text', self.on_primary)
        self.content_bg = role('content_bg', None)
        self.title = role('title', self.primary)
        self.body = role('body', None)
        # template.dark: the master maps text to lt1 and background to dk1 (PowerPoint's dark-theme mapping), so text
        # typed without a colour is light on a dark page.
        self.dark = bool(data['template'].get('dark'))
        # template.cover "brand": the brand cover/closing (subtitle top-right, 1pt rule at 2.3in, title bottom-left).
        self.brand_cover = data['template'].get('cover') == 'brand'


def fit_16_9(prs):
    """The default template is 4:3; squeeze every master and layout shape to the 16:9 height."""
    scale = HEIGHT / prs.slide_height
    prs.slide_width, prs.slide_height = WIDTH, HEIGHT
    sx = WIDTH / Inches(10)
    for master in prs.slide_masters:
        for owner in [master, *master.slide_layouts]:
            for shape in owner.shapes:
                # Only shapes with their own box: a layout placeholder without one follows the (scaled) master.
                if shape._element.find('.//a:xfrm', NS) is None:
                    continue
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                shape.left, shape.width = Emu(int(left * sx)), Emu(int(width * sx))
                shape.top, shape.height = Emu(int(top * scale)), Emu(int(height * scale))


def theme_of(prs):
    return prs.slide_master.part.part_related_by(RT.THEME)


def restyle_theme(prs, style):
    part = theme_of(prs)
    root = etree.fromstring(part.blob)
    scheme = root.find('.//a:clrScheme', NS)
    scheme.set('name', 'DSH')
    for name, value in style.scheme.items():
        slot = scheme.find(f'a:{name}', NS)
        for child in list(slot):
            slot.remove(child)
        etree.SubElement(slot, f'{{{A}}}srgbClr', val=value)
    fonts = root.find('.//a:fontScheme', NS)
    fonts.set('name', 'DSH')
    for kind in ('majorFont', 'minorFont'):
        node = fonts.find(f'a:{kind}', NS)
        node.find('a:latin', NS).set('typeface', style.latin)
        node.find('a:ea', NS).set('typeface', style.cjk)
    part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def text_style(placeholder, color, align=None, bullets=None, size=None, anchor=None):
    """Default colour (and alignment, bullets, size in pt) of text typed into a layout placeholder; other settings stay."""
    body = placeholder.text_frame._txBody
    if anchor:
        body_pr = body.find('a:bodyPr', NS)
        if body_pr is None:
            body_pr = etree.Element(f'{{{A}}}bodyPr')
            body.insert(0, body_pr)
        body_pr.set('anchor', anchor)
    lst = body.find('a:lstStyle', NS)
    if lst is None:
        lst = etree.Element(f'{{{A}}}lstStyle')
        body.insert(1, lst)
    level = lst.find('a:lvl1pPr', NS)
    if level is None:
        level = etree.SubElement(lst, f'{{{A}}}lvl1pPr')
    if align:
        level.set('algn', align)
    if bullets is False:
        for name in ('buNone', 'buChar', 'buAutoNum', 'buFont'):
            for old in level.findall(f'a:{name}', NS):
                level.remove(old)
        level.set('marL', '0')
        level.set('indent', '0')
        level.insert(0, etree.Element(f'{{{A}}}buNone'))
    rpr = level.find('a:defRPr', NS)
    if rpr is None:
        rpr = etree.SubElement(level, f'{{{A}}}defRPr')
    if size:
        rpr.set('sz', str(int(size * 100)))
    for old in rpr.findall('a:solidFill', NS):
        rpr.remove(old)
    fill = etree.Element(f'{{{A}}}solidFill')
    etree.SubElement(fill, f'{{{A}}}srgbClr', val=color)
    rpr.insert(0, fill)


def solid_background(layout, color):
    fill = layout.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(color)


def place(shape, left, top, width, height):
    shape.left, shape.top, shape.width, shape.height = Inches(left), Inches(top), Inches(width), Inches(height)


def hairline(layout, color):
    """The brand cover/closing rule: a 1pt bar across the slide at 2.3in, in lineOnPrimary (spec §4: 0, 165.6, 720, 1).
    A layout's shape collection has no add_shape, so the autoshape is written into its shape tree directly."""
    tree = layout.shapes._spTree
    shape_id = max((int(e) for e in tree.xpath('.//p:cNvPr/@id')), default=1) + 1
    sp = CT_Shape.new_autoshape_sp(shape_id, '细横线', 'rect', 0, Inches(2.3), WIDTH, Pt(1))
    tree.insert_element_before(sp, 'p:extLst')
    rule = Shape(sp, layout.shapes)
    rule.fill.solid()
    rule.fill.fore_color.rgb = RGBColor.from_string(color)
    rule.line.fill.background()


def style_layouts(prs, style):
    layouts = prs.slide_master.slide_layouts
    by_name = {layout.name: layout for layout in layouts}
    keep = {'Title Slide': '封面', 'Title and Content': '标题和内容', 'Title Only': '仅标题', 'Blank': '空白', 'Section Header': '结束页'}
    for layout in list(layouts):
        if layout.name not in keep:
            layouts.remove(layout)
    for old, new in keep.items():
        by_name[old].name = new
    cover, content, closing = by_name['Title Slide'], by_name['Title and Content'], by_name['Section Header']
    for layout in layouts:
        for shape in list(layout.placeholders):
            if any(name in str(shape.placeholder_format.type) for name in ('DATE', 'FOOTER', 'SLIDE_NUMBER')):
                shape._element.getparent().remove(shape._element)
    solid_background(cover, style.cover_bg)
    solid_background(closing, style.closing_bg)
    if style.content_bg:
        for layout in (content, by_name['Title Only'], by_name['Blank']):
            solid_background(layout, style.content_bg)
    if style.dark:
        cmap = prs.slide_master._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}clrMap')
        for key, value in {'bg1': 'dk1', 'tx1': 'lt1', 'bg2': 'dk2', 'tx2': 'lt2'}.items():
            cmap.set(key, value)
    if style.brand_cover:
        for layout in (cover, closing):
            hairline(layout, style.scheme['accent4'])
    # 封面: template.cover "brand" (default) as the brand cover: subtitle top-right, the rule, two-line title bottom-left
    # anchored to its box's bottom. Otherwise a big two-line title on the left with the subtitle under it.
    for shape in cover.placeholders:
        kind = str(shape.placeholder_format.type)
        if 'SUBTITLE' in kind:
            place(shape, *((5.4, 1.3, 4.3, 0.7) if style.brand_cover else (0.6, 3.6, 8.8, 0.8)))
            text_style(shape, style.cover_text, align='l', bullets=False, size=style.sizes['cover_subtitle'])
        elif 'TITLE' in kind:
            place(shape, *((0.6, 2.6, 8.2, 2.2) if style.brand_cover else (0.6, 1.2, 8.8, 2.2)))
            text_style(shape, style.cover_text, align='l', size=style.sizes['cover_title'],
                       anchor='b' if style.brand_cover else None)
    # 结束页: THANKS in its body placeholder (it has no title, so it is told apart from 仅标题); brand cover: in the
    # cover title's box, left-aligned, under the same rule; otherwise one centred word.
    for shape in list(closing.placeholders):
        kind = str(shape.placeholder_format.type)
        if 'TITLE' in kind and 'SUBTITLE' not in kind:
            shape._element.getparent().remove(shape._element)
        elif 'BODY' in kind:
            if style.brand_cover:
                place(shape, 0.6, 2.6, 8.2, 2.2)
                text_style(shape, style.cover_text, align='l', bullets=False, size=style.sizes['closing_title'], anchor='b')
            else:
                place(shape, 0.6, 1.9, 8.8, 1.6)
                text_style(shape, style.cover_text, align='ctr', bullets=False, size=style.sizes['closing_title'])
    # Content pages: title top-left in the primary colour, body below.
    for layout in (content, by_name['Title Only']):
        for shape in layout.placeholders:
            kind = str(shape.placeholder_format.type)
            if 'TITLE' in kind and 'SUBTITLE' not in kind:
                place(shape, 0.5, 0.35, 9.0, 0.9)
                text_style(shape, style.title, align='l', size=style.sizes['page_title'])
            elif 'BODY' in kind or 'OBJECT' in kind:
                place(shape, 0.5, 1.4, 9.0, 3.7)
                if style.body:
                    text_style(shape, style.body)


def build(name, root=None):
    root = root or STYLES
    data = json.loads((root / name / 'style.json').read_text(encoding='utf-8'))
    style = Style(data)
    prs = Presentation()
    fit_16_9(prs)
    restyle_theme(prs, style)
    style_layouts(prs, style)
    out = root / name / 'template.pptx'
    prs.save(out)
    names = [layout.name for layout in prs.slide_master.slide_layouts]
    print(f'{name}/template.pptx: {len(prs.slides)} slides, layouts {names}')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--style', help='style pack name (a folder under styles/); default: every pack')
    ap.add_argument('--styles-dir', type=Path, help='folder holding the packs (style-distill staging); default: styles/')
    a = ap.parse_args()
    root = a.styles_dir.resolve() if a.styles_dir else STYLES
    if not root.is_dir():
        ap.error(f'样式目录 {a.styles_dir} 不存在')
    packs = sorted(p.name for p in root.iterdir() if (p / 'style.json').is_file() and STYLE_NAME.match(p.name))
    if a.style is not None and a.style not in packs:
        ap.error(f'未知样式 {a.style}，可选 {", ".join(packs) or "无"}')
    for name in [a.style] if a.style else packs:
        build(name, root)
    return 0


if __name__ == '__main__':
    sys.exit(main())
