#!/usr/bin/env python3
"""Preview a style pack's live mode without DSH (style-distill tooling; developer machines only).

    "$PY" -I tooling/style-distill/live_preview.py --style <id> [--styles-dir <dir>] --out <dir> [--no-png]

In live mode the model starts from styles/<id>/template.pptx (presentation_create) and builds slides with the JS API:
addSlide on a template layout, setText into its placeholders, addTextBox / addImage at the slot boxes that
layouts.json "live" hands it. This script does the same with python-pptx: one slide per live layout with sample text in
its placeholders, then one slide per slot on that slot's layout (a fixture photo cropped to the box, darkened when the
slot says so, plus its text box; or sample text in a text slot), in the colours and sizes the pack declares. It writes
<out>/live.pptx and, unless --no-png, page images and <out>/live.png (an overview) through pptx2png.py.

It is a stand-in for watching DSH: what it shows is the template and the geometry the model is told, not the model's
choices. Text colour follows the page: template.colors.body/title when given, else onPrimary on primary pages and ink
on the others.
"""
import argparse
import importlib.util
import json
import sys
from pathlib import Path

from PIL import Image, ImageEnhance
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

HERE = Path(__file__).resolve().parent
SKILL = HERE.parents[1] / 'packages' / 'dsh-adapter' / 'skills' / 'shimo-pptx'
FIXTURES = HERE / 'fixtures' / 'img'
SAMPLE = {
    'title': '季度经营回顾', 'subtitle': '2026 年第三季度 · 管理层汇报', 'body': '收入同比增长 18%，毛利率回升至 42%\n新客户贡献三成增量\n下季度重点：渠道效率与交付周期',
    'closing': 'THANKS',
}


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.dont_write_bytecode = True
    spec.loader.exec_module(module)
    return module


def rgb(hex_color):
    return RGBColor.from_string(hex_color.lstrip('#').upper())


def crop_to(src, w, h, darken, out):
    im = Image.open(src).convert('RGB')
    target = w / h
    iw, ih = im.size
    if iw / ih > target:
        nw = int(ih * target)
        im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / target)
        im = im.crop((0, (ih - nh) // 2, iw, (ih - nh) // 2 + nh))
    if darken:
        im = ImageEnhance.Brightness(im).enhance(1 - darken)
    im.save(out, 'JPEG', quality=85)
    return out


def build(style_id, styles_dir, out_dir, photo_dir=None):
    pack = styles_dir / style_id
    style = json.loads((pack / 'style.json').read_text(encoding='utf-8'))
    live = json.loads((pack / 'layouts.json').read_text(encoding='utf-8'))['live']
    P = style['palettes'][style['default_palette']]
    tcol = style['template'].get('colors') or {}
    sizes = style['template']['sizes_pt']
    prs = Presentation(str(pack / 'template.pptx'))
    by_name = {layout.name: layout for layout in prs.slide_layouts}
    layout_of = {item['id']: item for item in live['layouts']}

    def ink_on(layout_id):
        bg = layout_of[layout_id].get('background', 'paper')
        if 'body' in tcol:
            return P[tcol['body']]
        return P['onPrimary'] if bg == 'primary' else P['ink']

    def add(layout_id):
        return prs.slides.add_slide(by_name[layout_of[layout_id]['name']])

    def textbox(slide, box, text, color, size):
        l, t, w, h = box
        tb = slide.shapes.add_textbox(Pt(l), Pt(t), Pt(w), Pt(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, line in enumerate(text.split('\n')):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run()
            r.text = line
            r.font.size = Pt(size)
            r.font.color.rgb = rgb(color)
        return tb

    # 1. Each live layout with its placeholders filled.
    for item in live['layouts']:
        slide = add(item['id'])
        for ph in slide.placeholders:
            kind = str(ph.placeholder_format.type)
            if 'SUBTITLE' in kind:
                ph.text = SAMPLE['subtitle']
            elif 'TITLE' in kind:
                ph.text = SAMPLE['title'] if item['id'] != 'closing' else SAMPLE['closing']
            elif 'BODY' in kind or 'OBJECT' in kind:
                ph.text = SAMPLE['closing'] if item['id'] == 'closing' else SAMPLE['body']
    # 2. Each slot on its layout.
    src = Path(photo_dir) if photo_dir else FIXTURES
    photos = sorted(p for p in src.iterdir() if p.suffix.lower() in ('.png', '.jpg', '.jpeg') and not p.name.startswith('logo'))
    tmp = out_dir / 'live-img'
    tmp.mkdir(parents=True, exist_ok=True)
    for k, slot in enumerate(live['slots']):
        slide = add(slot['layout'])
        for ph in slide.placeholders:
            if 'TITLE' in str(ph.placeholder_format.type) and 'SUBTITLE' not in str(ph.placeholder_format.type):
                ph.text = f'槽位 {slot["id"]}：{slot.get("note", "")}'
        color = ink_on(slot['layout'])
        if slot['kind'] == 'image':
            l, t, w, h = slot['box']
            img = crop_to(photos[k % len(photos)], w, h, slot.get('darken'), tmp / f'{k:02d}.jpg')
            slide.shapes.add_picture(str(img), Pt(l), Pt(t), Pt(w), Pt(h))
            if slot.get('text_box'):
                on_photo = slot.get('darken') or slot.get('text_on_image')
                textbox(slide, slot['text_box'], SAMPLE['body'] if not on_photo else SAMPLE['title'],
                        P['onPrimary'] if on_photo else color, sizes.get('body', 14) if not on_photo else sizes.get('page_title', 24))
        else:
            textbox(slide, slot['box'], SAMPLE['body'], color, sizes.get('body', 14))
    out_dir.mkdir(parents=True, exist_ok=True)
    deck = out_dir / 'live.pptx'
    prs.save(str(deck))
    return deck


def overview(images, out, cols=4, width=1280):
    cell_w = width // cols
    cell_h = cell_w * 9 // 16
    rows = -(-len(images) // cols)
    sheet = Image.new('RGB', (cell_w * cols, cell_h * rows), 'white')
    for i, p in enumerate(images):
        tile = Image.open(p).convert('RGB').resize((cell_w - 4, cell_h - 4), Image.LANCZOS)
        sheet.paste(tile, ((i % cols) * cell_w + 2, (i // cols) * cell_h + 2))
    sheet.save(out, optimize=True)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--style', required=True)
    ap.add_argument('--styles-dir', type=Path, help='folder holding the packs (staging); default: the skill styles/')
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--no-png', action='store_true')
    ap.add_argument('--photos', type=Path, help='folder of photos for the image slots (default: tooling fixtures)')
    a = ap.parse_args()
    styles_dir = a.styles_dir.resolve() if a.styles_dir else SKILL / 'styles'
    a.out.mkdir(parents=True, exist_ok=True)
    deck = build(a.style, styles_dir, a.out, a.photos)
    result = {'pptx': str(deck)}
    if not a.no_png:
        p2p = _load('style_distill_pptx2png', HERE / 'pptx2png.py')
        try:
            pages = p2p.convert(deck, a.out / 'live-pages', dpi=96)
        except p2p.ConvertError as e:
            print(str(e), file=sys.stderr)
            return e.code
        result.update(pages=len(pages), overview=str(overview(pages, a.out / 'live.png')))
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
