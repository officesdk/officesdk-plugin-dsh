#!/usr/bin/env python3
"""Reference pages in one shape: images / one overview image / PDF / PPTX → pages/NN.png + intake.json (style-distill).

    "$PY" -I ref_intake.py <图片目录 | 单张图片 | x.pdf | x.pptx> --out <dir>
          [--width 1920] [--single | --montage] [--boxes boxes.json] [--no-render] [--tolerance 12] [--min-gap 4]

- Image directory: png/jpg/jpeg/webp/gif/bmp/tif in natural file-name order, each saved as RGB PNG (not resized).
- One image: a slide-shaped image (16:9 or 4:3, ±2%) is one page. Otherwise it is split as an overview image by seam
  projection: background = most common colour of the outer frame; mask = pixels differing from it by more than
  --tolerance in any channel; recursive XY cuts along empty mask rows/columns at least --min-gap wide; each block
  tightened to its mask. It is an overview when there are ≥2 page-like blocks (≥1% of the image, ≥85% filled, aspect
  0.5–2.4) with aspects within ±4% of their median, covering ≥30% of the image, and no content left outside them
  (a page whose own background matches the overview's falls apart into fragments: when the other pages sit on a
  regular grid, the empty grid cell holding those fragments is taken as that page and noted). Anything
  else exits 3 asking for --boxes (JSON [[x0,y0,x1,y1],…] in source pixels) or --single. A slide-shaped image that
  also looks like an overview stays one page with a warning; pass --montage to split it.
- PDF: poppler (pdfinfo, pdftoppm scaled to --width, pdftotext for the text layer, pdffonts).
- PPTX: inventory read with the standard library after the zip checks (preflight.py); pages rendered through
  pptx2png.py (PowerPoint → PDF → pdftoppm) unless --no-render.

Input is untrusted data: nothing in it is executed; Pillow's pixel cap is lowered and a decompression bomb is an error;
PPTX XML with a DOCTYPE is refused. The output directory must be new or empty.
Exit codes: 0 ok; 2 input or environment problem; 3 overview image not decidable, or PowerPoint permission/timeout.
"""
import argparse
import hashlib
import importlib.util
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
import warnings
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
Image.MAX_IMAGE_PIXELS = 100_000_000
warnings.simplefilter('error', Image.DecompressionBombWarning)
IMAGE_EXT = {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.tif', '.tiff'}
SLIDE_ASPECTS = (16 / 9, 4 / 3)
PAGE_FILL = 0.85
PAGE_MIN_AREA = 0.01
MONTAGE_COVER = 0.30
ASPECT_SPREAD = 0.04
STRAY_MAX = 0.0005
CELL_MIN_CONTENT = 0.01


def _load(name, file):
    sys.dont_write_bytecode = True   # loading a sibling by path must not leave __pycache__ in the tooling folder
    spec = importlib.util.spec_from_file_location(name, HERE / file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


preflight = _load('style_distill_preflight', 'preflight.py')


class IntakeError(Exception):
    def __init__(self, message, code=2):
        super().__init__(message)
        self.code = code


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def natural_key(path):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', path.name)]


def open_rgb(path):
    try:
        with Image.open(path) as im:
            im.load()
            return im.convert('RGB')
    except (Image.DecompressionBombWarning, Image.DecompressionBombError) as e:
        raise IntakeError(f'{path.name} 像素过多，拒绝读取：{e}')
    except OSError as e:
        raise IntakeError(f'{path.name} 不是可读的图片：{e}')


# ---------- seam projection ----------
def frame_background(im):
    """Most common colour of the one-pixel outer frame."""
    w, h = im.size
    counts = {}
    for box in ((0, 0, w, 1), (0, h - 1, w, h), (0, 0, 1, h), (w - 1, 0, w, h)):
        for n, colour in im.crop(box).getcolors(w * h) or []:
            counts[colour] = counts.get(colour, 0) + n
    return max(counts.items(), key=lambda kv: kv[1])[0]


def content_mask(im, bg, tolerance):
    """L image: 255 where any channel differs from bg by more than tolerance."""
    channels = []
    for band, value in zip(im.split(), bg):
        channels.append(band.point(lambda v, value=value: 255 if abs(v - value) > tolerance else 0))
    r, g, b = channels
    return Image.eval(Image.merge('RGB', (r, g, b)).convert('L'), lambda v: 255 if v else 0)


def _runs(profile, min_gap):
    """Spans of non-empty entries, split at empty runs at least min_gap long (short empty runs stay inside)."""
    spans, start, gap = [], None, 0
    for i, v in enumerate(profile):
        if v:
            if start is None:
                start = i
            gap = 0
            end = i + 1
        elif start is not None:
            gap += 1
            if gap >= min_gap:
                spans.append((start, end))
                start, gap = None, 0
    if start is not None:
        spans.append((start, end))
    return spans


def _profile(mask, box, axis):
    """Per row (axis=1) or column (axis=0) of box: True when the mask has any pixel there."""
    region = mask.crop(box).convert('F')
    w, h = region.size
    line = region.resize((w, 1) if axis == 0 else (1, h), Image.BOX)
    return [line.getpixel((i, 0) if axis == 0 else (0, i)) > 0 for i in range(w if axis == 0 else h)]


def xy_cut(mask, box, min_gap, depth=0):
    """Leaf boxes of a recursive XY cut, each tightened to the mask."""
    tight = mask.crop(box).getbbox()
    if not tight:
        return []
    box = (box[0] + tight[0], box[1] + tight[1], box[0] + tight[2], box[1] + tight[3])
    if depth > 12:
        return [box]
    for axis in (1, 0):
        spans = _runs(_profile(mask, box, axis), min_gap)
        if len(spans) > 1:
            out = []
            for a, b in spans:
                sub = (box[0], box[1] + a, box[2], box[1] + b) if axis == 1 else (box[0] + a, box[1], box[0] + b, box[3])
                out += xy_cut(mask, sub, min_gap, depth + 1)
            return out
    return [box]


def _fill(mask, box):
    region = mask.crop(box)
    hist = region.histogram()
    return hist[255] / max(1, region.size[0] * region.size[1])


def order_rows(boxes):
    """Rows first (a box joins the current row when its vertical span overlaps the row's), then left to right."""
    rows = []
    for b in sorted(boxes, key=lambda b: (b[1], b[0])):
        if rows and b[1] < rows[-1]['y1'] and b[3] > rows[-1]['y0']:
            rows[-1]['items'].append(b)
            rows[-1]['y0'], rows[-1]['y1'] = min(rows[-1]['y0'], b[1]), max(rows[-1]['y1'], b[3])
        else:
            rows.append({'y0': b[1], 'y1': b[3], 'items': [b]})
    return [b for row in rows for b in sorted(row['items'], key=lambda b: b[0])]


def _inside(inner, outer, slack=2):
    return inner[0] >= outer[0] - slack and inner[1] >= outer[1] - slack and inner[2] <= outer[2] + slack and inner[3] <= outer[3] + slack


def fill_grid(pages, strays):
    """Pages of one size on a regular lattice: an empty lattice cell holding stray content is a page whose own
    background matches the overview background (it fell apart into fragments). Returns (added cells, strays left)."""
    if len(pages) < 2:
        return [], strays
    w0, h0 = pages[0][2] - pages[0][0], pages[0][3] - pages[0][1]
    if any(abs((b[2] - b[0]) - w0) > 2 or abs((b[3] - b[1]) - h0) > 2 for b in pages):
        return [], strays

    def lattice(values):
        out = []
        for v in sorted(values):
            if not out or v - out[-1] > 2:
                out.append(v)
        return out
    xs, ys = lattice(b[0] for b in pages), lattice(b[1] for b in pages)
    added = []
    for y in ys:
        for x in xs:
            cell = (x, y, x + w0, y + h0)
            if any(abs(b[0] - x) <= 2 and abs(b[1] - y) <= 2 for b in pages):
                continue
            inside = [s for s in strays if _inside(s, cell)]
            # Only real content fills a cell: specks (compression noise) in an empty cell are not a page.
            if inside and sum((s[2] - s[0]) * (s[3] - s[1]) for s in inside) >= w0 * h0 * CELL_MIN_CONTENT:
                added.append(cell)
                strays = [s for s in strays if s not in inside]
    return added, strays


def analyse_montage(im, tolerance, min_gap):
    w, h = im.size
    bg = frame_background(im)
    mask = content_mask(im, bg, tolerance)
    leaves = xy_cut(mask, (0, 0, w, h), min_gap)
    total = w * h
    pages, strays = [], []
    for b in leaves:
        bw, bh = b[2] - b[0], b[3] - b[1]
        if bw * bh >= total * PAGE_MIN_AREA and _fill(mask, b) >= PAGE_FILL and 0.5 <= bw / bh <= 2.4:
            pages.append(b)
        else:
            strays.append(b)
    filled, strays = fill_grid(pages, strays)
    pages += filled
    # Tiny specks (compression noise) may be ignored; any real content outside the pages makes the split unsafe.
    stray_area = sum((b[2] - b[0]) * (b[3] - b[1]) for b in strays)
    ok, why = False, ''
    if len(pages) < 2:
        why = f'只找到 {len(pages)} 个像页面的块'
    else:
        ratios = sorted((b[2] - b[0]) / (b[3] - b[1]) for b in pages)
        median = ratios[len(ratios) // 2]
        cover = sum((b[2] - b[0]) * (b[3] - b[1]) for b in pages) / total
        if any(abs(r - median) / median > ASPECT_SPREAD for r in ratios):
            why = f'{len(pages)} 个块的宽高比不一致（{ratios[0]:.2f}–{ratios[-1]:.2f}）'
        elif cover < MONTAGE_COVER:
            why = f'像页面的块只占全图 {cover:.0%}'
        elif stray_area > total * STRAY_MAX:
            why = f'页外还有 {len(strays)} 块内容（文字、图形，或某页底色与背景相近）'
        else:
            ok = True
    return {'ok': ok, 'why': why, 'background': '#%02X%02X%02X' % bg[:3], 'boxes': order_rows(pages) if ok else pages,
            'grid_filled': [list(b) for b in filled], 'strays': len(strays)}


def slide_shaped(w, h):
    r = w / h
    return any(abs(r - a) / a <= 0.02 for a in SLIDE_ASPECTS)


def read_boxes(path, w, h):
    try:
        boxes = json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, ValueError) as e:
        raise IntakeError(f'读不了框文件 {path}：{e}')
    if not isinstance(boxes, list) or not boxes:
        raise IntakeError('框文件要是 [[x0,y0,x1,y1], …]，至少 1 个')
    out = []
    for b in boxes:
        if not (isinstance(b, list) and len(b) == 4 and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in b)):
            raise IntakeError(f'框格式不对：{b}')
        x0, y0, x1, y1 = (int(round(v)) for v in b)
        if not (0 <= x0 < x1 <= w and 0 <= y0 < y1 <= h):
            raise IntakeError(f'框超出图片 {w}×{h} 或宽高为 0：{b}')
        out.append((x0, y0, x1, y1))
    for i, a in enumerate(out):
        for b in out[i + 1:]:
            if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
                raise IntakeError(f'框互相重叠：{list(a)} 与 {list(b)}')
    return out


# ---------- writers ----------
class Pages:
    def __init__(self, out):
        self.dir = out / 'pages'
        self.dir.mkdir(parents=True)
        self.rows = []

    def add(self, im, source, **extra):
        n = len(self.rows) + 1
        path = self.dir / f'{n:02d}.png'
        im.save(path, 'PNG')
        data = path.read_bytes()
        self.rows.append({'page': n, 'file': f'pages/{n:02d}.png', 'width': im.size[0], 'height': im.size[1],
                          'sha256': sha256_bytes(data), 'from': source, **extra})


def intake_images(files, pages):
    for f in files:
        data = f.read_bytes()
        pages.add(open_rgb(f), {'file': f.name, 'sha256': sha256_bytes(data)})


def intake_single(path, pages, a, notes):
    im = open_rgb(path)
    w, h = im.size
    if a.boxes:
        boxes = read_boxes(a.boxes, w, h)
        for b in boxes:
            pages.add(im.crop(b), {'box': list(b)})
        return 'montage' if len(boxes) > 1 else 'image', {'method': 'boxes', 'boxes': [list(b) for b in boxes]}
    if a.single:
        pages.add(im, {'file': path.name})
        return 'image', None
    result = analyse_montage(im, a.tolerance, a.min_gap)
    meta = {'method': 'seam', 'background': result['background'], 'tolerance': a.tolerance, 'min_gap': a.min_gap,
            'boxes': [list(b) for b in result['boxes']], 'grid_filled': result['grid_filled']}
    if not a.montage and slide_shaped(w, h):
        if result['ok']:
            warn = f'发现 {len(result["boxes"])} 个像页面的块；如果这是总览图，加 --montage'
            notes.append(warn)
            print(f'注意：{path.name} 是幻灯片画幅，按一页处理。{warn}', file=sys.stderr)
        pages.add(im, {'file': path.name})
        return 'image', None
    if not result['ok']:
        raise IntakeError(f'{path.name} 判不清是不是总览图（{result["why"]}）。请用 --boxes 给出每页的框（源图像素 [[x0,y0,x1,y1],…]），'
                          f'或用 --single 当一页', 3)
    if result['grid_filled']:
        notes.append(f'{len(result["grid_filled"])} 页底色与背景相近，按网格补齐：{result["grid_filled"]}，请核对')
    for b in result['boxes']:
        pages.add(im.crop(b), {'box': list(b)})
    return 'montage', meta


def render_pdf(pdf, out_dir, width, pages_expected):
    tool = preflight._poppler('pdftoppm')
    if not tool:
        raise IntakeError('没有找到 pdftoppm：请先 brew install poppler')
    try:
        run = subprocess.run([tool, '-scale-to-x', str(width), '-scale-to-y', '-1', '-png', str(pdf), str(out_dir / 'p')],
                             capture_output=True, text=True, timeout=max(120, 10 * pages_expected))
    except subprocess.TimeoutExpired:
        raise IntakeError('pdftoppm 超时')
    if run.returncode != 0:
        raise IntakeError(f'pdftoppm 失败：{run.stderr.strip()[:200]}')
    return sorted(out_dir.glob('p-*.png'), key=natural_key)


def intake_pdf(pdf, pages, width, chars=None):
    inv = preflight.pdf_inventory(pdf)
    if inv['approximate']:
        raise IntakeError('没有找到 pdfinfo：请先 brew install poppler')
    with tempfile.TemporaryDirectory(prefix='ref-intake-pdf-') as tmp:
        files = render_pdf(pdf, Path(tmp), width, inv['pages'])
        if len(files) != inv['pages']:
            raise IntakeError(f'pdftoppm 出了 {len(files)} 页，pdfinfo 说 {inv["pages"]} 页')
        per_page = inv['chars_per_page'] or [None] * len(files)
        for i, f in enumerate(files, 1):
            pages.add(open_rgb(f), {'pdf_page': i}, text_chars=per_page[i - 1])
    return inv


def intake_pptx(pptx, pages, a, notes):
    inv = preflight.pptx_inventory(pptx)
    if a.no_render:
        notes.append('--no-render：只盘点，没有出图')
        return inv, None
    pptx2png = _load('style_distill_pptx2png', 'pptx2png.py')
    with tempfile.TemporaryDirectory(prefix='ref-intake-pptx-') as tmp:
        pdf = Path(tmp) / 'deck.pdf'
        try:
            pptx2png._pdftoppm()
            if not pptx2png.POWERPOINT.exists():
                raise pptx2png.ConvertError(f'没有找到 {pptx2png.POWERPOINT}：PPTX 出图需要 Microsoft PowerPoint for Mac（或加 --no-render 只盘点）', 2)
            pptx2png.pptx_to_pdf(Path(pptx).resolve(), pdf, a.timeout)
        except pptx2png.ConvertError as e:
            raise IntakeError(str(e), e.code)
        pdf_inv = intake_pdf(pdf, pages, a.width)
    return inv, {'pages': pdf_inv['pages'], 'page_size_pt': pdf_inv['page_size_pt']}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('source', type=Path)
    ap.add_argument('--out', type=Path, required=True, help='输出目录（新建或空目录）')
    ap.add_argument('--width', type=int, default=1920, help='PDF / PPTX 出图宽度（px），默认 1920')
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument('--single', action='store_true', help='单张图片当一页，不切')
    mode.add_argument('--montage', action='store_true', help='单张图片当总览图切（判不清就退出 3）')
    mode.add_argument('--boxes', type=Path, help='单张图片按给定的框切：JSON [[x0,y0,x1,y1],…]（源图像素）')
    ap.add_argument('--no-render', action='store_true', help='PPTX 只盘点，不经 PowerPoint 出图')
    ap.add_argument('--tolerance', type=int, default=12, help='与背景色的通道差超过它才算内容，默认 12')
    ap.add_argument('--min-gap', type=int, default=4, help='页间缝至少多宽（px），默认 4')
    ap.add_argument('--timeout', type=int, default=180, help='PowerPoint 转图超时（秒）')
    a = ap.parse_args()
    src, out = a.source, a.out
    created = not out.exists()
    pages = None
    try:
        if out.exists() and (not out.is_dir() or any(out.iterdir())):
            raise IntakeError(f'输出目录 {out} 已存在且不是空目录，不覆盖')
        if not src.exists():
            raise IntakeError(f'找不到 {src}')
        ext = src.suffix.lower()
        single_flags = a.single or a.montage or a.boxes
        if single_flags and not (src.is_file() and ext in IMAGE_EXT):
            raise IntakeError('--single / --montage / --boxes 只用于单张图片')
        out.mkdir(parents=True, exist_ok=True)
        pages, notes = Pages(out), []
        report = {'tool': 'ref_intake', 'version': 1}
        if src.is_dir():
            files = sorted((p for p in src.iterdir() if p.is_file() and not p.is_symlink() and p.suffix.lower() in IMAGE_EXT
                            and not p.name.startswith('.')), key=natural_key)
            skipped = sorted(p.name for p in src.iterdir() if p not in files and not p.name.startswith('.'))
            if not files:
                raise IntakeError(f'{src} 里没有图片（png/jpg/jpeg/webp/gif/bmp/tif）')
            if skipped:
                notes.append('没读的文件：' + '、'.join(skipped[:20]) + (f' 等 {len(skipped)} 个' if len(skipped) > 20 else ''))
            intake_images(files, pages)
            report['source'] = {'path': str(src.resolve()), 'kind': 'images', 'files': len(files)}
        else:
            data_sha = preflight.sha256_file(src)
            report['source'] = {'path': str(src.resolve()), 'sha256': data_sha, 'bytes': src.stat().st_size}
            if ext in IMAGE_EXT:
                kind, meta = intake_single(src, pages, a, notes)
                report['source']['kind'] = kind
                if meta:
                    report['montage'] = meta
            elif ext == '.pdf':
                report['source']['kind'] = 'pdf'
                inv = intake_pdf(src, pages, a.width)
                report['pdf'] = {k: inv[k] for k in ('pages', 'page_size_pt', 'text_layer', 'text_chars', 'fonts')}
            elif ext == '.pptx' or ext in preflight.MACRO_PPT_EXT:
                report['source']['kind'] = 'pptx'
                inv, rendered = intake_pptx(src, pages, a, notes)
                report['pptx'] = inv
                if rendered:
                    report['pptx']['rendered'] = rendered
            else:
                raise IntakeError(f'不支持的类型 {ext or src.name}：只收图片目录、单张图片、PDF、PPTX')
    except Exception as e:  # noqa: BLE001 - untrusted input must end in a clear message and a clean output dir
        if isinstance(e, preflight.Refused):
            e = IntakeError(f'拒绝：{e}')
        elif not isinstance(e, IntakeError):
            e = IntakeError(f'读不了输入（内容损坏或格式不对）：{type(e).__name__}: {e}')
        # Clean up only what this run made: its pages/ (never an earlier run's), and the out dir if it created it.
        if pages is not None and out.is_dir() and not out.is_symlink():
            shutil.rmtree(pages.dir, ignore_errors=True)
            if created and not any(out.iterdir()):
                out.rmdir()
        print(str(e), file=sys.stderr)
        return e.code
    report['pages'] = pages.rows
    report['notes'] = notes
    (out / 'intake.json').write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding='utf-8')
    print(json.dumps({'out': str(out), 'kind': report['source']['kind'], 'pages': len(pages.rows),
                      'notes': notes}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
