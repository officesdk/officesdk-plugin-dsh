#!/usr/bin/env python3
"""Read a .pptx into comparable forms (style-distill tooling; developer machines only).

    <python with python-pptx> -I inspect_pptx.py dump  <deck.pptx>     # readable shape list per slide (JSON)
    <python with python-pptx> -I inspect_pptx.py parts <deck.pptx>     # sha256 of every normalised part (JSON)

parts(): every zip part except docProps/* is compared as canonical XML (C14N) when it parses as XML, else by bytes.
Embedded chart workbooks (ppt/embeddings/*.xlsx) are opened and each of their XML parts compared the same way,
except the workbook's own docProps/* (it carries a timestamp). Zip entry dates never count.
"""
import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path

from lxml import etree

EMU_PER_PX = 9525


def _canonical(data):
    try:
        root = etree.fromstring(data, etree.XMLParser(remove_blank_text=False, resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError:
        return None
    return etree.tostring(root, method='c14n')


def _walk(zf, prefix=''):
    out = {}
    for name in sorted(zf.namelist()):
        if name.startswith('docProps/') or name.endswith('/'):
            continue
        data = zf.read(name)
        key = prefix + name
        if name.lower().endswith('.xlsx'):
            with zipfile.ZipFile(io.BytesIO(data)) as inner:
                out.update(_walk(inner, key + '!'))
            continue
        canon = _canonical(data) if name.lower().endswith(('.xml', '.rels')) else None
        out[key] = canon if canon is not None else data
    return out


def part_bytes(path):
    """{part name: canonical bytes} for every comparable part."""
    with zipfile.ZipFile(path) as zf:
        return _walk(zf)


def parts(path):
    return {name: hashlib.sha256(data).hexdigest() for name, data in part_bytes(path).items()}


def _color(fill):
    try:
        if fill.type is None:
            return None
        if fill.type == 5:  # background (no fill)
            return 'none'
        return str(fill.fore_color.rgb)
    except Exception:  # noqa: BLE001 - theme colours, gradients
        return '?'


def _px(v):
    return None if v is None else round(v / EMU_PER_PX, 2)


def dump(path):
    """Readable per-slide shape list: kind, name, px geometry, fill, text runs with size/colour/font."""
    from pptx import Presentation
    prs = Presentation(str(path))
    slides = []
    for index, slide in enumerate(prs.slides, 1):
        shapes = []
        for sh in slide.shapes:
            item = {'kind': sh.shape_type and str(sh.shape_type).split('.')[-1].split(' ')[0], 'name': sh.name,
                    'box': [_px(sh.left), _px(sh.top), _px(sh.width), _px(sh.height)]}
            if sh.has_text_frame and sh.text_frame.text:
                paras = []
                for p in sh.text_frame.paragraphs:
                    runs = [{'t': r.text, 'pt': r.font.size and r.font.size.pt,
                             'rgb': str(r.font.color.rgb) if r.font.color and r.font.color.type is not None else None,
                             'face': r._r.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}latin').get('typeface')
                             if r._r.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}latin') is not None else None}
                            for r in p.runs]
                    paras.append(runs)
                item['text'] = paras
            elif getattr(sh, 'shape_type', None) is not None and sh.shape_type == 1:
                item['fill'] = _color(sh.fill)
            if sh.shape_type == 13:  # picture
                item['image'] = hashlib.sha256(sh.image.blob).hexdigest()[:16]
            if getattr(sh, 'has_chart', False) and sh.has_chart:
                ch = sh.chart
                item['chart'] = {'type': str(ch.chart_type), 'categories': list(ch.plots[0].categories),
                                 'series': [[s.name, list(s.values)] for s in ch.plots[0].series]}
            if getattr(sh, 'has_table', False) and sh.has_table:
                item['table'] = [[c.text for c in row.cells] for row in sh.table.rows]
            shapes.append(item)
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
        slides.append({'slide': index, 'background': _color(slide.background.fill), 'shapes': shapes, 'notes': notes})
    return slides


def main(argv):
    if len(argv) != 3 or argv[1] not in ('dump', 'parts'):
        print(__doc__, file=sys.stderr)
        return 2
    path = Path(argv[2])
    print(json.dumps(dump(path) if argv[1] == 'dump' else parts(path), ensure_ascii=False, indent=1))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
