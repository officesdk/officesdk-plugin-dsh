#!/usr/bin/env python3
"""Draft a layouts.json "offline" part from measured HTML slides (style-distill tooling; standard library only).

    python3 -I tooling/style-distill/gs2layouts.py <measure.json> --layouts <source layouts.json> --out draft.json
        [--ids cover,overview,...] [--notes conversion-notes.md] [--min-px 12]

Input: measure.json from measure_html.mjs (boxes and computed styles of every region, text slot and visual slot,
in px on a 1920×1080 canvas) and the source layouts.json (Genspark layout: texts / visuals / regions per page).
Output: our layouts.json draft (960×540 px) plus a notes file of everything that needs a person.

Rules:
- Geometry ÷ 2. Font size px ÷ 2 is our px (pt = px × 0.75); anything under --min-px (12 px = 9 pt, the validator's
  floor) is raised to it and listed, since text that was 3–5 pt on the source cannot just be halved.
- Text: the slots of one region become text boxes by column. Slots stacked in one column (same left edge, same size)
  are one box with a paragraph each (never a box per line); side-by-side columns are a box per column; a region whose
  slots form a grid (≥ 3 rows × ≥ 2 columns) is a table candidate: drawn as column boxes, flagged in the notes.
- Colours: rgb()/rgba() → #RRGGBB; partial opacity (colour alpha × element opacity) is premixed with the page
  background, since one run cannot be transparent in our renderer. The page background becomes role primary, the
  commonest text colour onPrimary; others stay hex.
- Fonts: the first family of each text maps to a font role (serif / sans / mono) for style.json fonts.roles.
- letter-spacing → spacing (pt); italic → italic (the renderer skips it on Chinese lines); text-transform uppercase →
  noted (write the text in capitals); bold weights ≥ 600 → bold.
- Shapes: a region with a background colour → rect (1–2 px thick → hline / vline; border → outline; 50% radius →
  ellipse shape); linear-gradient backgrounds → rect.gradient; radial/conic gradients, filters other than
  brightness, blend modes, transforms → listed as not converted, with the substitution the notes suggest.
- Visual slots of kind image → image slots (needed = required); filter brightness(x) → mask 1 − x on the image.
"""
import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

RGB = re.compile(r'rgba?\(\s*([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\s*\)')


def parse_rgb(value):
    m = RGB.search(value or '')
    if not m:
        return None
    r, g, b = (float(m.group(k)) for k in (1, 2, 3))
    a = float(m.group(4)) if m.group(4) is not None else 1.0
    return (r, g, b, a)


def to_hex(rgb):
    return '#%02X%02X%02X' % tuple(int(round(max(0, min(255, c)))) for c in rgb[:3])


def premix(rgba, alpha, bg):
    a = rgba[3] * alpha
    return to_hex(tuple(rgba[k] * a + bg[k] * (1 - a) for k in range(3)))


def px(v):
    return round(v / 2, 1) if v is not None else None


def num(v):
    return int(v) if float(v).is_integer() else round(v, 1)


def font_role(family):
    first = (family or '').split(',')[0].strip().strip('"\'').lower()
    if 'mono' in first or 'code' in first:
        return 'mono'
    if re.search(r'serif|times|garamond|fraunces|playfair|tinos|stix|cormorant|slab', first) and 'sans' not in first:
        return 'serif'
    return 'sans'


def gradient(value, bg):
    """linear-gradient(...) → {angle, stops: [[pos, hex, alpha], ...]}, or None."""
    m = re.match(r'linear-gradient\((.*)\)$', (value or '').strip())
    if not m:
        return None
    body = m.group(1)
    parts = re.split(r',\s*(?![^()]*\))', body)
    angle = 180.0
    if parts and re.match(r'^-?[\d.]+deg$', parts[0].strip()):
        angle = float(parts.pop(0).strip()[:-3])
    elif parts and parts[0].strip().startswith('to '):
        angle = {'to bottom': 180, 'to top': 0, 'to right': 90, 'to left': 270}.get(parts.pop(0).strip(), 180)
    stops = []
    for k, part in enumerate(parts):
        c = parse_rgb(part)
        if not c:
            return None
        pm = re.search(r'([\d.]+)%\s*$', part)
        pos = float(pm.group(1)) / 100 if pm else (k / max(1, len(parts) - 1))
        stops.append([round(pos, 3), to_hex(c), round(c[3], 3)])
    return {'angle': angle, 'stops': stops} if len(stops) >= 2 else None


def radius(value, box):
    """True when the corner radius makes the box a circle/ellipse (50%, or px ≥ half the short side); first corner only."""
    first = (value or '0px').split()[0]
    if first.endswith('%'):
        return float(first[:-1] or 0) >= 50
    try:
        return float(first.rstrip('px') or 0) >= min(box[2], box[3]) / 2 - 1
    except ValueError:
        return False


def columns(entries):
    """Group text entries into columns: same left edge (±6 px) and stacked; returns lists sorted top→bottom."""
    cols = []
    for e in sorted(entries, key=lambda e: (e['box'][0], e['box'][1])):
        for col in cols:
            if abs(col[0]['box'][0] - e['box'][0]) <= 6:
                col.append(e)
                break
        else:
            cols.append([e])
    for col in cols:
        col.sort(key=lambda e: e['box'][1])
    return cols


def split_runs(col):
    """Within a column, consecutive entries with the same size and colour form one box (paragraphs)."""
    runs, cur = [], []
    for e in col:
        key = (e['style']['font-size'], e['style']['color'], e['style']['font-family'], e['style']['font-weight'])
        if cur and key == cur[-1][0]:
            cur.append((key, e))
        else:
            if cur:
                runs.append([x for _, x in cur])
            cur = [(key, e)]
    if cur:
        runs.append([x for _, x in cur])
    return runs


def convert(measure, source, ids, min_px):
    notes, fonts_used = [], Counter()
    out_layouts = {}
    pages = measure['pages']
    order = sorted(pages)   # page ids as strings: '01' … '12' (JSON key order from Node puts 10–12 first)
    ids = ids or [f'p{pid}' for pid in order]
    # Page background: the slide container's colour isn't in the regions; take the commonest full-page dark/solid fill.
    for lid, pid in zip(ids, order):
        page = pages[pid]
        src = source['layouts'][pid]
        page_bg = parse_rgb(page.get('background') or '') or (255, 255, 255, 1)
        bg_hex = to_hex(page_bg)
        objects, fields = [], {}
        # Regions in z order (then document order).
        regs = sorted(page['regions'].items(), key=lambda kv: (int(kv[1]['style'].get('z-index') or 0) if (kv[1]['style'].get('z-index') or '').lstrip('-').isdigit() else 0))
        slot_region = {}
        for rid, r in page['regions'].items():
            for t in src['regions'].get(rid, {}).get('texts', []):
                slot_region[t] = rid
        for rid, r in regs:
            st, box = r['style'], r['box']
            if not box:
                continue
            x, y, w, h = (px(v) for v in box)
            kind = r['kind']
            filt = st.get('filter', 'none')
            if kind == 'image':
                for vid in src['regions'].get(rid, {}).get('visuals', []):
                    vis = src.get('visuals', {}).get(vid, {})
                    vstyle = page['visuals'].get(vid, {}).get('style') or {}
                    if vis.get('kind') != 'image':
                        notes.append(f'{lid}：视觉槽 {vid} 是 {vis.get("kind")}，没有转换（SVG/图示要重画成形状或原生图表）')
                        continue
                    name = f'photo{len([f for f in fields if f.startswith("photo")]) + 1}' if any(f.startswith('photo') for f in fields) else 'photo'
                    fields[name] = {'type': 'image', 'needed': bool(vis.get('required'))}
                    img = {'type': 'image', 'slot': name, 'x': num(x), 'y': num(y), 'w': num(w), 'h': num(h), 'panel': 'panel'}
                    bm = re.search(r'brightness\(([\d.]+)\)', vstyle.get('filter', '') + ' ' + vis.get('style', ''))
                    if bm:
                        img['mask'] = round(1 - float(bm.group(1)), 2)
                        notes.append(f'{lid}：{vid} 的 CSS filter brightness({bm.group(1)}) → 黑色遮罩 mask {img["mask"]}（contrast 等其余滤镜放弃）')
                    objects.append(img)
                continue
            bgimg = st.get('background-image', 'none')
            fill = parse_rgb(st.get('background-color'))
            opacity = float(st.get('opacity') or 1)
            if bgimg and bgimg != 'none':
                g = gradient(bgimg, page_bg)
                if g:
                    objects.append({'type': 'rect', 'x': num(x), 'y': num(y), 'w': num(w), 'h': num(h), 'fill': None, 'gradient': g, 'name': rid})
                else:
                    notes.append(f'{lid}：{rid} 背景 {bgimg[:60]}… 做不到（径向/锥形渐变或图片），改纯色或删掉')
            if fill and fill[3] * opacity > 0.01:
                color = premix(fill, opacity, page_bg)
                border = float((st.get('border-top-width') or '0px')[:-2] or 0)
                if h <= 1.5 or w <= 1.5:
                    objects.append({'type': 'hline', 'x': num(x), 'y': num(y), 'w': num(w), 'color': color} if h <= w else
                                   {'type': 'vline', 'x': num(x), 'y': num(y), 'h': num(h), 'color': color})
                elif radius(st.get('border-radius'), box):
                    objects.append({'type': 'shape', 'kind': 'ellipse', 'x': num(x), 'y': num(y), 'w': num(w), 'h': num(h), 'fill': color})
                else:
                    rect = {'type': 'rect', 'x': num(x), 'y': num(y), 'w': num(w), 'h': num(h), 'fill': color, 'name': rid}
                    bc = parse_rgb(st.get('border-top-color'))
                    if border and bc:
                        rect['outline'] = premix(bc, opacity, page_bg)
                    objects.append(rect)
            elif kind == 'shape':
                border = float((st.get('border-top-width') or '0px')[:-2] or 0)
                bc = parse_rgb(st.get('border-top-color'))
                if border and bc:
                    objects.append({'type': 'rect', 'x': num(x), 'y': num(y), 'w': num(w), 'h': num(h), 'fill': None, 'outline': premix(bc, opacity, page_bg), 'name': rid})
            for key, label in (('mix-blend-mode', '混合模式'), ('transform', '旋转/变形')):
                if st.get(key) not in (None, '', 'none', 'normal'):
                    notes.append(f'{lid}：{rid} 有 {label} {st[key]}，没有转换')
            if filt not in (None, '', 'none') and kind != 'image':
                notes.append(f'{lid}：{rid} 有 filter {filt}，没有转换')
            # Texts of this region.
            entries = [dict(page['texts'][t], id=t) for t in src['regions'].get(rid, {}).get('texts', []) if t in page['texts'] and page['texts'][t].get('box')]
            seen, uniq = set(), []
            for e in entries:   # one element can carry several slots (data-text-slots): measure it once
                key = tuple(e['box'])
                if key not in seen:
                    seen.add(key)
                    uniq.append(e)
            cols = columns(uniq)
            if len(cols) >= 2 and min(len(c) for c in cols) >= 3:
                notes.append(f'{lid}：{rid} 的文字排成 {max(len(c) for c in cols)} 行 × {len(cols)} 列，像表格：草稿按列画，建议改成原生 table')
            for col in cols:
                for run in split_runs(col):
                    first = run[0]
                    st0 = first['style']
                    size_src = float(st0['font-size'][:-2])
                    size = size_src / 2
                    if size < min_px:
                        notes.append(f'{lid}：{first["id"]} 原 {size_src:g}px（{size * 0.75:.1f}pt）→ 抬到 {min_px}px（9pt）')
                        size = min_px
                    color = parse_rgb(st0['color']) or (255, 255, 255, 1)
                    op = float(st0.get('opacity') or 1)
                    lh_raw = st0.get('line-height', 'normal')
                    lh = round(float(lh_raw[:-2]) / size_src, 2) if lh_raw.endswith('px') else 1.2
                    field = re.sub(r'^.*\.', '', first['id'])
                    fields[field] = {'type': 'text'}
                    xs = [e['box'][0] for e in run]
                    box = [min(xs), run[0]['box'][1], max(e['box'][0] + e['box'][2] for e in run) - min(xs), run[-1]['box'][1] + run[-1]['box'][3] - run[0]['box'][1]]
                    role = font_role(st0.get('font-family'))
                    fonts_used[(role, (st0.get('font-family') or '').split(',')[0].strip().strip('"'))] += 1
                    obj = {'type': 'text', 'name': field, 'text': f"d.get('{field}', '')", 'x': num(px(box[0])), 'y': num(px(box[1])), 'w': num(max(px(box[2]), 20)),
                           'size': num(size), 'color': premix(color, op, page_bg), 'lh': lh, 'max_h': num(max(px(box[3]), size * lh)), 'font': role}
                    weight = int(st0.get('font-weight') or 400)
                    if weight >= 600:
                        obj['bold'] = True
                    elif weight <= 300:
                        obj['weight'] = 300
                    if st0.get('font-style') == 'italic':
                        obj['italic'] = True
                    ls = st0.get('letter-spacing', 'normal')
                    if ls.endswith('px') and float(ls[:-2]) != 0:
                        obj['spacing'] = round(float(ls[:-2]) / 2 * 0.75, 2)
                    if st0.get('text-align') in ('center', 'right'):
                        obj['align'] = st0['text-align']
                    if st0.get('text-transform') == 'uppercase':
                        notes.append(f'{lid}：{field} 原文 CSS 大写（text-transform），英文内容直接写大写')
                    if len(run) > 1:
                        notes.append(f'{lid}：{field} 合并了 {len(run)} 个槽（{", ".join(e["id"].split(".")[-1] for e in run)}），一段一个')
                    fields[field]['example'] = '\n'.join(e['text'] for e in run)
                    objects.append(obj)
        out_layouts[lid] = {'title': page.get('name', lid), 'description': f"{page.get('name', lid)}（字段：{', '.join(fields)}）", 'background': bg_hex, 'fields': fields, 'objects': objects}
    return {'version': 1, 'deck': {'first': ids[0], 'last': ids[-1]},
            'offline': {'canvas': {'width': 960, 'height': 540, 'unit': 'px'}, 'order': ids, 'layouts': out_layouts},
            '_fonts': [{'role': r, 'family': f, 'uses': n} for (r, f), n in fonts_used.most_common()]}, notes


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('measure', type=Path)
    ap.add_argument('--layouts', type=Path, required=True)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--ids')
    ap.add_argument('--notes', type=Path)
    ap.add_argument('--min-px', type=float, default=12)
    a = ap.parse_args()
    measure = json.loads(a.measure.read_text(encoding='utf-8'))
    source = json.loads(a.layouts.read_text(encoding='utf-8'))
    ids = a.ids.split(',') if a.ids else None
    if ids and len(ids) != len(measure['pages']):
        print(f'--ids 有 {len(ids)} 个，测量结果有 {len(measure["pages"])} 页', file=sys.stderr)
        return 2
    draft, notes = convert(measure, source, ids, a.min_px)
    draft['_notes'] = notes
    a.out.write_text(json.dumps(draft, ensure_ascii=False, indent=1), encoding='utf-8')
    if a.notes:
        by = defaultdict(list)
        for n in notes:
            by[n.split('：')[0]].append(n.split('：', 1)[1])
        lines = ['# 转换说明（gs2layouts.py 草稿）', '', '字体角色：' + '；'.join(f"{f['role']} ← {f['family']}（{f['uses']} 处）" for f in draft['_fonts']), '']
        for lid in draft['offline']['order']:
            lay = draft['offline']['layouts'][lid]
            texts = sum(1 for o in lay['objects'] if o['type'] == 'text')
            lines.append(f"## {lid} {lay['title']}（{len(lay['objects'])} 个对象，文本框 {texts}）")
            lines += [f'- {n}' for n in by.get(lid, [])] or ['- 无']
            lines.append('')
        a.notes.write_text('\n'.join(lines), encoding='utf-8')
    print(json.dumps({'layouts': draft['offline']['order'], 'notes': len(notes), 'fonts': draft['_fonts']}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
