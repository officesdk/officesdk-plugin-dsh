#!/usr/bin/env python3
"""Measure reference pages for a hand-written layouts.json: colours, ink boxes, lines, font-size estimates (style-distill).

    "$PY" -I measure_image.py <page.png | intake.json> [--page N] [--roi name=x,y,w,h …] [--units canvas|src]
          [--threshold 24] [--colors 6] [--out measure.json]

Per page: pixel size, the scale to our 960-wide canvas, the background (most common colour of the outer frame) and the
main colours (exact colour counts on a ≤480-wide nearest-neighbour copy, top --colors with their share).

--roi name=x,y,w,h (canvas coordinates by default; --units src for source pixels; one page only): the region's own
background is the most common colour of its frame; ink is every pixel differing from it by more than --threshold in
any channel. Reports the ink box (canvas and source), the ink colours (exact counts, top 3), text lines from the row
projection (top/bottom/height), the line pitch when there are two or more lines, and font-size guesses:
size_px_if_caps = height / 0.70, size_px_if_mixed = height / 0.95, size_px_from_pitch = pitch / 1.2 (canvas px; pt =
px × 0.75). These are estimates: in trial D reading sizes off screenshots came out about 20% too large.

Given an intake.json whose source is a PDF with a text layer, each page also gets the text lines from
`pdftotext -bbox-layout` (text, box in pt and canvas px, a size guess from the line height), marked
source: "pdf-text-layer". Without a text layer the text must be read off the image by the agent, each piece marked
with where it came from. Measures only: no scoring, no layouts written.
"""
import argparse
import importlib.util
import json
import re
import subprocess
import sys
import warnings
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
CANVAS_W = 960
Image.MAX_IMAGE_PIXELS = 100_000_000
warnings.simplefilter('error', Image.DecompressionBombWarning)
XHTML = '{http://www.w3.org/1999/xhtml}'


def _load(name, file):
    sys.dont_write_bytecode = True   # loading a sibling by path must not leave __pycache__ in the tooling folder
    spec = importlib.util.spec_from_file_location(name, HERE / file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


preflight = _load('style_distill_preflight', 'preflight.py')


class MeasureError(Exception):
    pass


def hexc(c):
    return '#%02X%02X%02X' % tuple(c[:3])


def open_rgb(path):
    try:
        with Image.open(path) as im:
            im.load()
            return im.convert('RGB')
    except (Image.DecompressionBombWarning, Image.DecompressionBombError, OSError) as e:
        raise MeasureError(f'{path} 读不了：{e}')


def frame_mode(im):
    w, h = im.size
    counts = {}
    for box in ((0, 0, w, 1), (0, h - 1, w, h), (0, 0, 1, h), (w - 1, 0, w, h)):
        for n, colour in im.crop(box).getcolors(w * h) or []:
            counts[colour] = counts.get(colour, 0) + n
    return max(counts.items(), key=lambda kv: kv[1])[0]


def ink_mask(im, bg, threshold):
    bands = [band.point(lambda v, ref=ref: 255 if abs(v - ref) > threshold else 0) for band, ref in zip(im.split(), bg)]
    merged = Image.merge('RGB', bands).convert('L')
    return merged.point(lambda v: 255 if v else 0)


def top_colours(im, n, total=None, exclude=None):
    colours = im.getcolors(im.size[0] * im.size[1]) or []
    if exclude is not None:
        colours = [c for c in colours if c[1] != exclude]
    total = total or sum(c[0] for c in colours) or 1
    return [{'hex': hexc(c), 'share': round(k / total, 4)} for k, c in sorted(colours, key=lambda kc: -kc[0])[:n]]


def palette(im, n):
    w, h = im.size
    small = im if w <= 480 else im.resize((480, max(1, round(h * 480 / w))), Image.NEAREST)
    return top_colours(small, n)


def row_runs(mask, gap=2):
    w, h = mask.size
    col = mask.convert('F').resize((1, h), Image.BOX)
    rows = [col.getpixel((0, y)) > 0 for y in range(h)]
    runs, start, empty = [], None, 0
    for y, v in enumerate(rows):
        if v:
            if start is None:
                start = y
            empty, end = 0, y + 1
        elif start is not None:
            empty += 1
            if empty >= gap:
                runs.append((start, end))
                start, empty = None, 0
    if start is not None:
        runs.append((start, end))
    return runs


def size_guesses(height_px, pitch_px=None):
    out = {'ink_height_px': round(height_px, 1), 'size_px_if_caps': round(height_px / 0.70, 1), 'size_px_if_mixed': round(height_px / 0.95, 1)}
    if pitch_px:
        out['size_px_from_pitch'] = round(pitch_px / 1.2, 1)
    out['pt'] = {k.replace('_px', ''): round(v * 0.75, 1) for k, v in out.items() if k.startswith('size_px')}
    return out


def measure_roi(im, scale, name, box, units, threshold):
    x, y, w, h = box
    if units == 'canvas':
        src = (round(x / scale), round(y / scale), round((x + w) / scale), round((y + h) / scale))
    else:
        src = (round(x), round(y), round(x + w), round(y + h))
    W, H = im.size
    src = (max(0, src[0]), max(0, src[1]), min(W, src[2]), min(H, src[3]))
    if src[2] - src[0] < 2 or src[3] - src[1] < 2:
        raise MeasureError(f'ROI {name} 在图外或太小：{box}')
    region = im.crop(src)
    bg = frame_mode(region)
    mask = ink_mask(region, bg, threshold)
    bbox = mask.getbbox()
    out = {'name': name, 'roi_src': list(src), 'region_background': hexc(bg)}
    if not bbox:
        out['ink'] = None
        return out
    ink_pixels = mask.histogram()[255]
    only_ink = Image.composite(region, Image.new('RGB', region.size, bg), mask)
    ink_src = [src[0] + bbox[0], src[1] + bbox[1], src[0] + bbox[2], src[1] + bbox[3]]
    lines = []
    for a, b in row_runs(mask):
        lines.append({'top': round((src[1] + a) * scale, 1), 'bottom': round((src[1] + b) * scale, 1), 'height': round((b - a) * scale, 1)})
    tops = [line['top'] for line in lines]
    pitches = sorted(b - a for a, b in zip(tops, tops[1:]))
    pitch = pitches[len(pitches) // 2] if pitches else None
    tallest = max((line['height'] for line in lines), default=0)
    out['ink'] = {
        'box_canvas': {'x': round(ink_src[0] * scale, 1), 'y': round(ink_src[1] * scale, 1),
                       'w': round((ink_src[2] - ink_src[0]) * scale, 1), 'h': round((ink_src[3] - ink_src[1]) * scale, 1)},
        'box_src': ink_src, 'pixels': ink_pixels,
        'colours': top_colours(only_ink, 3, total=ink_pixels, exclude=bg),
        'lines': lines, 'pitch_px': pitch,
        'size_estimate': size_guesses(tallest, pitch) if lines else None,
    }
    return out


def pdf_text_lines(pdf, page):
    tool = preflight._poppler('pdftotext')
    if not tool:
        return None, '没有 pdftotext（brew install poppler）'
    try:
        run = subprocess.run([tool, '-bbox-layout', '-f', str(page), '-l', str(page), str(pdf), '-'], capture_output=True, timeout=120)
    except subprocess.TimeoutExpired:
        return None, 'pdftotext 超时'
    if run.returncode != 0:
        return None, 'pdftotext 失败：' + run.stderr.decode('utf-8', 'replace')[:200]
    text = run.stdout.decode('utf-8', 'replace')
    text = re.sub(r'<!DOCTYPE[^>]*>', '', text, count=1)
    root = ET.fromstring(text)
    lines = []
    for pg in root.iter(f'{XHTML}page'):
        scale = CANVAS_W / float(pg.get('width'))
        for line in pg.iter(f'{XHTML}line'):
            x0, y0, x1, y1 = (float(line.get(k)) for k in ('xMin', 'yMin', 'xMax', 'yMax'))
            words = [w.text or '' for w in line.iter(f'{XHTML}word')]
            h = (y1 - y0) * scale
            lines.append({'text': ' '.join(words), 'source': 'pdf-text-layer', 'box_pt': [round(v, 2) for v in (x0, y0, x1, y1)],
                          'box_canvas': {'x': round(x0 * scale, 1), 'y': round(y0 * scale, 1), 'w': round((x1 - x0) * scale, 1), 'h': round(h, 1)},
                          'size_px_est': round(h / 1.2, 1), 'size_pt_est': round(h / 1.2 * 0.75, 1)})
    return lines, None


def measure_page(path, rois, units, threshold, colours):
    im = open_rgb(path)
    w, h = im.size
    scale = CANVAS_W / w
    page = {'image': str(path), 'size_src': [w, h], 'scale_to_canvas': round(scale, 6), 'canvas': [CANVAS_W, round(h * scale, 1)],
            'background': hexc(frame_mode(im)), 'palette': palette(im, colours)}
    if rois:
        page['rois'] = [measure_roi(im, scale, name, box, units, threshold) for name, box in rois]
    return page


def parse_roi(text):
    m = re.fullmatch(r'([^=]+)=\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)', text)
    if not m:
        raise argparse.ArgumentTypeError(f'ROI 要写成 名字=x,y,w,h：{text}')
    return m.group(1).strip(), tuple(float(v) for v in m.groups()[1:])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('source', type=Path, help='一张页面图，或 ref_intake.py 写的 intake.json')
    ap.add_argument('--page', type=int, help='intake.json 里只量第几页（从 1 起）')
    ap.add_argument('--roi', type=parse_roi, action='append', default=[], help='名字=x,y,w,h，可重复；只能配单页')
    ap.add_argument('--units', choices=['canvas', 'src'], default='canvas', help='ROI 坐标：960 画布（默认）或源图像素')
    ap.add_argument('--threshold', type=int, default=24, help='与区域底色的通道差超过它算墨迹，默认 24')
    ap.add_argument('--colors', type=int, default=6, help='每页列几个主色，默认 6')
    ap.add_argument('--out', type=Path, help='写 JSON 到这里（默认打印）')
    a = ap.parse_args()
    try:
        result = {'tool': 'measure_image', 'version': 1, 'canvas_width': CANVAS_W, 'notes': [], 'pages': []}
        if a.source.suffix.lower() == '.json':
            intake = json.loads(a.source.read_text(encoding='utf-8'))
            if intake.get('tool') != 'ref_intake':
                raise MeasureError(f'{a.source} 不是 ref_intake.py 写的 intake.json')
            rows = intake.get('pages') or []
            if a.page is not None:
                rows = [r for r in rows if r['page'] == a.page]
                if not rows:
                    raise MeasureError(f'intake.json 里没有第 {a.page} 页')
            if a.roi and len(rows) != 1:
                raise MeasureError('--roi 只能配单页：加 --page N')
            src = intake.get('source', {})
            pdf = intake.get('pdf') or {}
            if src.get('kind') == 'pdf' and pdf.get('text_layer') is False:
                result['notes'].append('PDF 无文字层：文字靠 agent 看图识读，notes.md 里逐条标来源（agent 识读 / 低可信）')
            for row in rows:
                page = measure_page(a.source.parent / row['file'], a.roi, a.units, a.threshold, a.colors)
                page['page'] = row['page']
                if src.get('kind') == 'pdf' and (row.get('text_chars') or 0) > 0:
                    lines, why = pdf_text_lines(src["path"], row["from"]["pdf_page"])
                    page['text_lines'] = lines
                    if why:
                        result['notes'].append(f'第 {row["page"]} 页：{why}')
                result['pages'].append(page)
        else:
            if a.page is not None:
                raise MeasureError('--page 只用于 intake.json')
            result['pages'].append(measure_page(a.source, a.roi, a.units, a.threshold, a.colors))
        result['notes'].append('字号是估计值：只看图时偏大约 20%（试验 D），以对比关卡为准')
    except (MeasureError, OSError, ValueError, KeyError, ET.ParseError) as e:
        print(f'测量失败：{e}', file=sys.stderr)
        return 2
    text = json.dumps(result, ensure_ascii=False, indent=1)
    if a.out:
        a.out.write_text(text, encoding='utf-8')
        print(json.dumps({'out': str(a.out), 'pages': len(result['pages'])}, ensure_ascii=False))
    else:
        print(text)
    return 0


if __name__ == '__main__':
    sys.exit(main())
