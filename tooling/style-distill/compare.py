#!/usr/bin/env python3
"""Side-by-side pages and automatic checks: the original vs our style pack (style-distill tooling; this Mac only).

    <python with python-pptx and Pillow> -I compare.py --theirs <original.pptx | .pdf | images dir>
        (--ours <ours.pptx> | --spec <deck.json> --style <id>) --out <dir> [--labels 原稿,样式包] [--dpi 96]

Both sides become page images (PPTX through pptx2png.py: Microsoft PowerPoint + poppler, see there; PDF through
pdftoppm; an images directory is used as is, sorted by name), then page N of each is put side by side:
<out>/pages/NN.png and an overview <out>/compare.png. Each PPTX side is also checked (<out>/checks.json and
<out>/report.md):
  overflow      text estimated (same character widths as brand_deck.py) to run past its box
  overlap       two text boxes whose estimated ink overlaps by more than 10% of the smaller one
  font size     smallest run, and text boxes under 9pt
  duplicates    the same picture (by SHA-256) used more than once, logos excepted
  editability   text boxes, paragraphs, single-line boxes, pictures, native charts and tables, full-page pictures,
                external links, and "one line per text box" (2+ single-line boxes stacked at line spacing)
Exit 0 when our side has no overflow, overlap, duplicate, sub-9pt text or one-line-per-box finding, else 1.
"""
import argparse
import hashlib
import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent
SKILL = HERE.parents[1] / 'packages' / 'dsh-adapter' / 'skills' / 'shimo-pptx'
PAGE = (720, 405)
LABEL_H = 34
EMU_PER_PX = 9525
MIN_PT = 9


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    was, sys.dont_write_bytecode = sys.dont_write_bytecode, True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = was
    return module


def _font(size):
    for path in ('/System/Library/Fonts/PingFang.ttc', '/System/Library/Fonts/STHeiti Light.ttc', '/System/Library/Fonts/Hiragino Sans GB.ttc'):
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


# ---------- page images ----------
def page_images(src, work, dpi):
    src = Path(src)
    if src.is_dir():
        return sorted(p for p in src.iterdir() if p.suffix.lower() in ('.png', '.jpg', '.jpeg'))
    return _load('style_distill_pptx2png', HERE / 'pptx2png.py').convert(src, work, dpi=dpi)


def _fit(path):
    canvas = Image.new('RGB', PAGE, (236, 236, 236))
    if path is None:
        ImageDraw.Draw(canvas).text((PAGE[0] // 2 - 40, PAGE[1] // 2 - 10), '（无此页）', fill=(120, 120, 120), font=_font(20))
        return canvas
    img = Image.open(path).convert('RGB')
    img.thumbnail(PAGE, Image.LANCZOS)
    canvas.paste(img, ((PAGE[0] - img.width) // 2, (PAGE[1] - img.height) // 2))
    return canvas


def side_by_side(theirs, ours, labels, out):
    pages_dir = out / 'pages'
    pages_dir.mkdir(parents=True, exist_ok=True)
    font = _font(18)
    pairs = []
    for i in range(max(len(theirs), len(ours))):
        sheet = Image.new('RGB', (PAGE[0] * 2 + 12, PAGE[1] + LABEL_H), 'white')
        draw = ImageDraw.Draw(sheet)
        for k, (side, label) in enumerate(zip((theirs, ours), labels)):
            x = k * (PAGE[0] + 12)
            draw.text((x + 6, 6), f'{label} · 第 {i + 1} 页', fill=(40, 40, 40), font=font)
            sheet.paste(_fit(side[i] if i < len(side) else None), (x, LABEL_H))
        path = pages_dir / f'{i + 1:02d}.png'
        sheet.save(path, optimize=True)
        pairs.append(path)
    if pairs:
        thumb_w = 960
        rows = [Image.open(p) for p in pairs]
        scale = thumb_w / rows[0].width
        h = int(rows[0].height * scale)
        overview = Image.new('RGB', (thumb_w, h * len(rows)), 'white')
        for k, row in enumerate(rows):
            overview.paste(row.resize((thumb_w, h), Image.LANCZOS), (0, k * h))
        overview.save(out / 'compare.png', optimize=True)
    return pairs


# ---------- checks ----------
def _px(emu):
    return (emu or 0) / EMU_PER_PX


def _ink(shape, est):
    """Estimated ink rectangle (px) of a text box: as wide as its widest line (or the box when wrapping), as tall as its lines."""
    x, y, w, h = _px(shape.left), _px(shape.top), _px(shape.width), _px(shape.height)
    widest, height = est
    align = shape.text_frame.paragraphs[0].alignment
    ink_w = min(w, widest) if not shape.text_frame.word_wrap or widest < w else w
    if align is not None and int(align) == 2:      # centre
        x += (w - ink_w) / 2
    elif align is not None and int(align) == 3:    # right
        x += w - ink_w
    return (x, y, ink_w, min(h, height) if height else h)


def _overlap(a, b):
    ix = max(0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]))
    iy = max(0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))
    smaller = min(a[2] * a[3], b[2] * b[3]) or 1
    return ix * iy / smaller


def check_pptx(path, renderer):
    from pptx import Presentation
    bd = _load('style_distill_brand_deck', renderer)
    prs = Presentation(str(path))
    slide_w, slide_h = _px(prs.slide_width), _px(prs.slide_height)
    findings = {'overflow': [], 'overlap': [], 'small_text': [], 'duplicate_images': [], 'one_line_per_box': [], 'full_page_pictures': []}
    stats = {'slides': len(prs.slides), 'text_boxes': 0, 'paragraphs': 0, 'single_line_boxes': 0, 'pictures': 0, 'charts': 0, 'tables': 0}
    min_pt, seen = None, {}
    for n, slide in enumerate(prs.slides, 1):
        texts = []
        for sh in slide.shapes:
            if sh.shape_type == 13:
                stats['pictures'] += 1
                digest = hashlib.sha256(sh.image.blob).hexdigest()
                if sh.name != 'brand-logo':
                    if digest in seen:
                        findings['duplicate_images'].append({'slide': n, 'shape': sh.name, 'same_as': seen[digest]})
                    else:
                        seen[digest] = f'第 {n} 页 {sh.name}'
                if _px(sh.width) * _px(sh.height) >= 0.9 * slide_w * slide_h:
                    findings['full_page_pictures'].append({'slide': n, 'shape': sh.name})
            if getattr(sh, 'has_chart', False) and sh.has_chart:
                stats['charts'] += 1
            if getattr(sh, 'has_table', False) and sh.has_table:
                stats['tables'] += 1
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            tf = sh.text_frame
            stats['text_boxes'] += 1
            stats['paragraphs'] += len(tf.paragraphs)
            sizes = [r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size]
            size_pt = max(sizes) if sizes else 18.0
            if sizes:
                low = min(sizes)
                min_pt = low if min_pt is None else min(min_pt, low)
                if low < MIN_PT:
                    findings['small_text'].append({'slide': n, 'shape': sh.name, 'pt': round(low, 2), 'text': tf.text[:20]})
            size_px = size_pt / 0.75
            spacing = tf.paragraphs[0].line_spacing   # None, a multiple (float) or an exact Length
            line_px = 1.2 * size_px if spacing is None else spacing * size_px if isinstance(spacing, float) else spacing.pt / 0.75
            lines = tf.text.split('\n')
            w, h = _px(sh.width), _px(sh.height)
            widest = max(bd.text_width(line, size_px) for line in lines)
            if tf.word_wrap is False:
                rows = len(lines)
                if widest > w * 1.05:
                    findings['overflow'].append({'slide': n, 'shape': sh.name, 'kind': '宽度', 'need_px': round(widest), 'box_px': round(w), 'text': lines[0][:20]})
            else:
                rows = sum(max(1, -(-int(bd.text_width(line, size_px)) // max(1, int(w)))) for line in lines)
                if rows * line_px > h + size_px * 0.3:
                    findings['overflow'].append({'slide': n, 'shape': sh.name, 'kind': '高度', 'need_px': round(rows * line_px), 'box_px': round(h), 'text': lines[0][:20]})
            single = len(tf.paragraphs) == 1 and len(lines) == 1 and rows == 1
            stats['single_line_boxes'] += single
            texts.append({'shape': sh, 'ink': _ink(sh, (widest, rows * line_px)), 'single': single, 'size': round(size_pt, 1), 'x': round(_px(sh.left)), 'y': _px(sh.top), 'line': line_px})
        for i, a in enumerate(texts):
            for b in texts[i + 1:]:
                ratio = _overlap(a['ink'], b['ink'])
                if ratio > 0.10:
                    findings['overlap'].append({'slide': n, 'shapes': [a['shape'].name, b['shape'].name], 'ratio': round(ratio, 2),
                                                'texts': [a['shape'].text_frame.text[:12], b['shape'].text_frame.text[:12]]})
        # One line per text box: 2+ single-line boxes, same left edge and size, stacked at about one line apart.
        groups = {}
        for t in texts:
            if t['single']:
                groups.setdefault((round(t['x'] / 2), t['size']), []).append(t)
        for (_, size), items in groups.items():
            items.sort(key=lambda t: t['y'])
            run = [items[0]] if items else []
            for prev, cur in zip(items, items[1:]):
                gap = cur['y'] - prev['y']
                if 0.95 * size / 0.75 <= gap <= 1.45 * size / 0.75:
                    run.append(cur)
                else:
                    if len(run) >= 2:
                        findings['one_line_per_box'].append({'slide': n, 'boxes': len(run), 'size_pt': size, 'first': run[0]['shape'].text_frame.text[:16]})
                    run = [cur]
            if len(run) >= 2:
                findings['one_line_per_box'].append({'slide': n, 'boxes': len(run), 'size_pt': size, 'first': run[0]['shape'].text_frame.text[:16]})
    with zipfile.ZipFile(path) as zf:
        external = sorted({name for name in zf.namelist() if name.endswith('.rels') and b'TargetMode="External"' in zf.read(name)})
    stats['single_line_ratio'] = round(stats['single_line_boxes'] / stats['text_boxes'], 2) if stats['text_boxes'] else 0
    stats['min_pt'] = round(min_pt, 2) if min_pt is not None else None
    stats['external_links'] = external
    return {'stats': stats, 'findings': findings}


BLOCKING = ('overflow', 'overlap', 'small_text', 'duplicate_images', 'one_line_per_box')


def report_md(result, labels):
    out = [f'# 对比：{labels[0]} vs {labels[1]}', '', f'并排图：`pages/NN.png`，总览：`compare.png`（左 {labels[0]}，右 {labels[1]}）。', '']
    for side, label in (('theirs', labels[0]), ('ours', labels[1])):
        r = result.get(side)
        if not r:
            out += [f'## {label}', '', '不是 PPTX，只出图，不做检查。', '']
            continue
        s, f = r['stats'], r['findings']
        out += [f'## {label}', '',
                f'{s["slides"]} 页；文本框 {s["text_boxes"]}（段落 {s["paragraphs"]}，单行文本框占 {s["single_line_ratio"]:.0%}）；图片 {s["pictures"]}；原生图表 {s["charts"]}；原生表格 {s["tables"]}；最小字号 {s["min_pt"]}pt；外链 {len(s["external_links"])}。', '',
                '| 检查 | 结果 |', '|---|---|']
        names = {'overflow': '溢出', 'overlap': '重叠', 'small_text': f'小于 {MIN_PT}pt', 'duplicate_images': '重复图片',
                 'one_line_per_box': '疑似一行一个文本框', 'full_page_pictures': '整页图片'}
        for key, name in names.items():
            items = f[key]
            detail = '；'.join(json.dumps(i, ensure_ascii=False) for i in items[:4]) + (f' …共 {len(items)} 处' if len(items) > 4 else '')
            out.append(f'| {name} | {"无" if not items else ("**" + str(len(items)) + " 处** " + detail)} |')
        out.append('')
    return '\n'.join(out) + '\n'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--theirs', type=Path, required=True)
    ap.add_argument('--ours', type=Path)
    ap.add_argument('--spec', type=Path)
    ap.add_argument('--style', default='default')
    ap.add_argument('--renderer', type=Path, default=SKILL / 'scripts' / 'brand_deck.py')
    ap.add_argument('--styles-dir', type=Path, help='folder holding the packs (staging); default: the skill styles/')
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--labels', default='原稿,样式包')
    ap.add_argument('--dpi', type=int, default=96)
    a = ap.parse_args()
    if bool(a.ours) == bool(a.spec):
        ap.error('给 --ours，或给 --spec 加 --style，二选一')
    labels = (a.labels.split(',') + ['原稿', '样式包'])[:2]
    a.out.mkdir(parents=True, exist_ok=True)
    pptx2png = _load('style_distill_pptx2png', HERE / 'pptx2png.py')
    with tempfile.TemporaryDirectory(prefix='style-compare-') as tmp:
        tmp = Path(tmp)
        ours = a.ours
        if a.spec:
            ours = tmp / 'ours.pptx'
            run = subprocess.run([sys.executable, '-I', str(a.renderer), 'build', '--spec', str(a.spec), '--out', str(ours), '--style', a.style] + (['--styles-dir', str(a.styles_dir)] if a.styles_dir else []),
                                 capture_output=True, text=True)
            if run.returncode != 0:
                print(f'渲染失败：{run.stderr.strip()}', file=sys.stderr)
                return 2
        try:
            theirs_pages = page_images(a.theirs, tmp / 'theirs', a.dpi)
            ours_pages = page_images(ours, tmp / 'ours', a.dpi)
        except pptx2png.ConvertError as e:
            print(str(e), file=sys.stderr)
            return e.code
        pairs = side_by_side(theirs_pages, ours_pages, labels, a.out)
        result = {'labels': labels, 'pages': {'theirs': len(theirs_pages), 'ours': len(ours_pages)},
                  'ours': check_pptx(ours, a.renderer),
                  'theirs': check_pptx(a.theirs, a.renderer) if a.theirs.suffix.lower() == '.pptx' else None}
    (a.out / 'checks.json').write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding='utf-8')
    (a.out / 'report.md').write_text(report_md(result, labels), encoding='utf-8')
    blocking = {k: len(result['ours']['findings'][k]) for k in BLOCKING if result['ours']['findings'][k]}
    print(json.dumps({'out': str(a.out), 'pages': len(pairs), 'ours_blocking': blocking}, ensure_ascii=False))
    return 1 if blocking else 0


if __name__ == '__main__':
    sys.exit(main())
