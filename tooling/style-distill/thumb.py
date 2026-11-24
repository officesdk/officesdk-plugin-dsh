#!/usr/bin/env python3
"""Make a style pack's thumbnail: styles/<id>/thumb.png (style-distill tooling; developer machines only).

    <python with python-pptx and Pillow> -I thumb.py --style <id> [--pages 1,2,5,7] [--out <png>]
    <python with python-pptx and Pillow> -I thumb.py --style <id> --previews 3,4,8

--previews writes one 480x270 page image per given page (2–4 pages) as styles/<id>/preview-<n>.png — the sample pages
the plugin's template gallery shows under the thumbnail — and prints the `meta.previews` list to put in style.json
(the file is not rewritten: style.json keeps its hand formatting).
--styles-dir reads (and writes) the pack in another folder; --spec renders another deck instead of example.json (the
same deck with photos from tooling/style-distill/fixtures/img, for packs whose example has none).

Renders the pack's example.json with the offline renderer (brand_deck.py), turns it into page images with
pptx2png.py (Microsoft PowerPoint + poppler, see there), and writes a 480x270 PNG: a 2x2 grid of the chosen pages
(default 1,2,5,7), or the page itself when one page is given. Fonts the style names but this Mac lacks are replaced
by PowerPoint, so the thumbnail shows the layout and colours, not the exact typeface.
"""
import argparse
import importlib.util
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
SKILL = HERE.parents[1] / 'packages' / 'dsh-adapter' / 'skills' / 'shimo-pptx'
SIZE = (480, 270)


def _pptx2png():
    spec = importlib.util.spec_from_file_location('style_distill_pptx2png', HERE / 'pptx2png.py')
    module = importlib.util.module_from_spec(spec)
    sys.dont_write_bytecode = True
    spec.loader.exec_module(module)
    return module


def _render(style, tmp, renderer=None, styles_dir=None, spec=None):
    """The pack's example.json (or spec) rendered and turned into page images; returns their paths."""
    renderer = renderer or SKILL / 'scripts' / 'brand_deck.py'
    root = Path(styles_dir).resolve() if styles_dir else SKILL / 'styles'
    example = Path(spec).resolve() if spec else root / style / 'example.json'
    deck = Path(tmp) / 'example.pptx'
    run = subprocess.run([sys.executable, '-I', str(renderer), 'build', '--spec', str(example), '--out', str(deck), '--style', style, '--styles-dir', str(root)],
                         capture_output=True, text=True)
    if run.returncode != 0:
        raise SystemExit(f'渲染 example.json 失败：{run.stderr.strip()}')
    return _pptx2png().convert(deck, Path(tmp) / 'png', dpi=144)


def _check_pages(images, pages):
    missing = [p for p in pages if p < 1 or p > len(images)]
    if missing:
        raise SystemExit(f'example.json 只有 {len(images)} 页，没有第 {missing} 页')


def make_previews(style, pages, renderer=None, styles_dir=None, spec=None):
    """styles/<style>/preview-<n>.png for each page; returns their file names in page order."""
    if not 2 <= len(pages) <= 4 or len(set(pages)) != len(pages):
        raise SystemExit('--previews 给 2–4 个不重复的页码')
    folder = (Path(styles_dir).resolve() if styles_dir else SKILL / 'styles') / style
    with tempfile.TemporaryDirectory(prefix='style-previews-') as tmp:
        images = _render(style, tmp, renderer, styles_dir, spec)
        _check_pages(images, pages)
        names = []
        for page in pages:
            name = f'preview-{page}.png'
            Image.open(images[page - 1]).convert('RGB').resize(SIZE, Image.LANCZOS).save(folder / name, optimize=True)
            names.append(name)
    return names


def make(style, pages, out, renderer=None, styles_dir=None, spec=None):
    with tempfile.TemporaryDirectory(prefix='style-thumb-') as tmp:
        images = _render(style, tmp, renderer, styles_dir, spec)
        _check_pages(images, pages)
        if len(pages) == 1:
            sheet = Image.open(images[pages[0] - 1]).convert('RGB').resize(SIZE, Image.LANCZOS)
        else:
            if len(pages) != 4:
                raise SystemExit('--pages 给 1 页或 4 页')
            cell = (SIZE[0] // 2, SIZE[1] // 2)
            sheet = Image.new('RGB', SIZE, 'white')
            for k, page in enumerate(pages):
                tile = Image.open(images[page - 1]).convert('RGB').resize(cell, Image.LANCZOS)
                sheet.paste(tile, ((k % 2) * cell[0], (k // 2) * cell[1]))
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out, optimize=True)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--style', required=True)
    ap.add_argument('--pages', default='1,2,5,7')
    ap.add_argument('--out', type=Path)
    ap.add_argument('--styles-dir', type=Path, help='folder holding the packs (staging); default: the skill styles/')
    ap.add_argument('--spec', type=Path, help='deck to render instead of the pack example.json (e.g. the same deck with photos)')
    ap.add_argument('--previews', help='write preview-<n>.png for these 2–4 pages instead of the thumbnail')
    a = ap.parse_args()
    if a.previews:
        try:
            names = make_previews(a.style, [int(p) for p in a.previews.split(',') if p.strip()], styles_dir=a.styles_dir, spec=a.spec)
        except _pptx2png().ConvertError as e:
            print(str(e), file=sys.stderr)
            return e.code
        print(json.dumps({'previews': names, 'meta': {'previews': names}, 'size': list(SIZE)}, ensure_ascii=False))
        return 0
    pages = [int(p) for p in a.pages.split(',') if p.strip()]
    out = a.out or (a.styles_dir.resolve() if a.styles_dir else SKILL / 'styles') / a.style / 'thumb.png'
    try:
        path = make(a.style, pages, out, styles_dir=a.styles_dir, spec=a.spec)
    except _pptx2png().ConvertError as e:
        print(str(e), file=sys.stderr)
        return e.code
    print(json.dumps({'thumb': str(path), 'bytes': path.stat().st_size, 'size': list(SIZE)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
