#!/usr/bin/env python3
"""Draft a layouts.json "offline" part from an object-level master JSON (style-distill tooling; standard library only).

    python3 -I tooling/style-distill/master2layouts.py <master.json> --out draft-layouts.json [--ids cover,index,...] [--report draft.md]

The master is the brand-ppt / Songyan kind: {width: 960, height: 540, pages: [{name, background, objects: [...]}]}, each
object with id, type (text, rect, image, triangle), x/y/w/h in px on a 960×540 canvas, and for text fontSize, weight,
color, lineHeight, maxLines. The same canvas as ours, so geometry is copied. What the script does beyond copying:

- text → one `text` object reading d['<id>'] (several lines are paragraphs of that one box, never a box per line);
  maxLines 1 → fit() so a long line shrinks (to the validator's 0.62 floor) instead of running out of its box; several
  lines → wrap with max_h, so the renderer warns on overflow.
- numbered siblings with the same step (row1_label/row1_value/row1_rule, row2_…; step1_*, phase1, action1_*) → one
  `repeat` over a list field (rows: [{label, value}, …]) with x/y offset by the step: fewer fields, any count up to max.
- rect of height or width 1 → hline / vline; triangle → shape (kind triangle, flip_v from verticalFlip); image → image
  slot with mask = maskOpacity (needed: true when the master requires it).
- the page background and the commonest text colour become palette roles primary / onPrimary; other colours stay hex.

The output is a draft: field names, which fields are required, extra page types (fewer photos) and charts drawn by the
source's code still need a person. Unknown object types are listed in the report, not guessed.
"""
import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

NUMBERED = re.compile(r'^([a-z]+?)(\d+)(?:_(\w+))?$')


def hex6(c):
    c = c.strip()
    if re.fullmatch(r'#[0-9A-Fa-f]{3}', c):
        c = '#' + ''.join(ch * 2 for ch in c[1:])
    return c.upper()


def color_ref(c, roles):
    c = hex6(c)
    return roles.get(c, c)


MIN_PX = 12   # 9pt: the validator's floor for text (px × 0.75)


def text_obj(o, roles, ref, x, y, name, notes=None):
    size = o['fontSize']
    if size < MIN_PX:
        if notes is not None:
            notes.append(f'{name}：{size}px（{size * 0.75:g}pt）小于 9pt，抬到 {MIN_PX}px')
        size = MIN_PX
    lines = o.get('maxLines', 1)
    # fit() may shrink a one-line text, but never under 9pt: its floor ratio is raised to MIN_PX / size when 0.62 is too low.
    ratio = max(0.62, -(-MIN_PX * 100 // size) / 100)
    fit = f"fit({ref}, {num(size)}, {num(o['w'])})" if ratio == 0.62 else f"fit({ref}, {num(size)}, {num(o['w'])}, {ratio})"
    out = {'type': 'text', 'name': name, 'text': f"{ref}", 'x': x, 'y': y, 'w': o['w'],
           'size': (fit if ratio < 1 else size) if lines == 1 else size,
           'color': color_ref(o.get('color', '#000000'), roles), 'lh': o.get('lineHeight', 1.17)}
    if o.get('weight', 400) != 400:
        out['weight'] = o['weight']
    if lines == 1:
        out['wrap'] = False
    else:
        out['max_h'] = o['h']
    return out


def shape_obj(o, roles, x, y):
    t = o['type']
    if t == 'rect':
        fill = o.get('fill')
        if o['h'] <= 1 and fill:
            return {'type': 'hline', 'x': x, 'y': y, 'w': o['w'], 'color': color_ref(fill, roles)}
        if o['w'] <= 1 and fill:
            return {'type': 'vline', 'x': x, 'y': y, 'h': o['h'], 'color': color_ref(fill, roles)}
        return {'type': 'rect', 'x': x, 'y': y, 'w': o['w'], 'h': o['h'], 'fill': color_ref(fill, roles) if fill and fill != 'none' else None}
    if t in ('triangle', 'rtTriangle', 'ellipse', 'parallelogram'):
        out = {'type': 'shape', 'kind': 'triangle' if t == 'triangle' else t, 'x': x, 'y': y, 'w': o['w'], 'h': o['h'], 'fill': color_ref(o.get('fill', '#000000'), roles)}
        if o.get('verticalFlip'):
            out['flip_v'] = True
        if o.get('horizontalFlip'):
            out['flip_h'] = True
        return out
    return None


def groups_of(objects):
    """{base: {index: {suffix: object}}} for numbered ids whose members repeat with one constant step."""
    found = defaultdict(lambda: defaultdict(dict))
    for o in objects:
        m = NUMBERED.match(o.get('id', ''))
        if m:
            found[m.group(1)][int(m.group(2))][m.group(3) or ''] = o
    out = {}
    for base, members in found.items():
        idx = sorted(members)
        if len(idx) < 2 or idx != list(range(idx[0], idx[0] + len(idx))):
            continue
        suffixes = set(members[idx[0]])
        if any(set(members[i]) != suffixes for i in idx):
            continue
        if any(o['type'] == 'image' for i in idx for o in members[i].values()):
            continue   # photos stay separate slots: each is its own field, counted by the validator
        first = members[idx[0]]
        steps = set()
        for k in range(1, len(idx)):
            for s in suffixes:
                a, b = members[idx[k - 1]][s], members[idx[k]][s]
                steps.add((round(b['x'] - a['x'], 3), round(b['y'] - a['y'], 3)))
                if (a['w'], a['h'], a['type']) != (b['w'], b['h'], b['type']):
                    steps.add(None)
        if len(steps) == 1 and None not in steps:
            out[base] = {'indices': idx, 'suffixes': sorted(suffixes), 'step': steps.pop(), 'first': first}
    return out


def num(v):
    return int(v) if float(v).is_integer() else round(v, 3)


def offset(base, step, axis):
    base = num(base)
    return base if not step else f"{base} + i * {num(step)}"


def convert_page(page, roles, layout_id):
    objects = page['objects']
    groups = groups_of(objects)
    in_group = {o['id'] for g in groups.values() for o in g['first'].values()}
    for g_base, g in groups.items():
        for i in g['indices']:
            for s in g['suffixes']:
                in_group.add(f"{g_base}{i}" + (f"_{s}" if s else ''))
    fields, out, notes = {}, [], []
    emitted = set()
    for o in objects:
        oid = o.get('id', '')
        m = NUMBERED.match(oid)
        if m and m.group(1) in groups and oid in in_group:
            base = m.group(1)
            if base in emitted:
                continue
            emitted.add(base)
            g = groups[base]
            dx, dy = g['step']
            field = base + 's'
            n = len(g['indices'])
            inner = []
            item_fields = []
            for s in g['suffixes']:
                fo = g['first'][s]
                x, y = offset(fo['x'], dx, 'x'), offset(fo['y'], dy, 'y')
                if fo['type'] == 'text':
                    key = s or 'text'
                    item_fields.append(key)
                    inner.append(text_obj(fo, roles, f"item.get('{key}', '')", x, y, f"{base}-{key}", notes))
                elif fo['type'] == 'image':
                    notes.append(f'{layout_id}：{base} 组里有图片，按单独槽位处理更好（已跳过）')
                else:
                    sh = shape_obj(fo, roles, x, y)
                    if sh:
                        inner.append(sh)
            fields[field] = {'type': 'list', 'max': n, 'items': item_fields}
            out.append({'type': 'repeat', 'items': f"d.get('{field}', [])", 'max': n, 'label': field, 'objects': inner})
            continue
        t = o['type']
        if t == 'text':
            fields[oid] = {'type': 'text', 'required': bool(o.get('requiresContent'))} if oid == 'title' else {'type': 'text'}
            obj = text_obj(o, roles, f"d.get('{oid}', '')", o['x'], o['y'], oid, notes)
            if oid != 'title':
                obj['when'] = f"d.get('{oid}')"
            out.append(obj)
        elif t == 'image':
            fields[oid] = {'type': 'image', 'needed': bool(o.get('requiresContent'))}
            img = {'type': 'image', 'slot': oid, 'x': o['x'], 'y': o['y'], 'w': o['w'], 'h': o['h'], 'panel': 'secondary'}
            if o.get('maskOpacity'):
                img['mask'] = o['maskOpacity']
            out.append(img)
        else:
            sh = shape_obj(o, roles, o['x'], o['y'])
            if sh:
                out.append(sh)
            else:
                notes.append(f'{layout_id}：不认识的对象类型 {t}（{oid}），没有转换')
    bg = color_ref(page.get('background', '#FFFFFF'), roles)
    return {'title': page.get('name', layout_id), 'description': f"{page.get('name', layout_id)}（字段：{', '.join(fields)}）",
            'background': bg, 'fields': fields, 'objects': out}, notes


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('master', type=Path)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--ids', help='comma-separated layout ids for the pages, in order (default: p1, p2, …)')
    ap.add_argument('--report', type=Path)
    a = ap.parse_args()
    m = json.loads(a.master.read_text(encoding='utf-8'))
    if (m.get('width'), m.get('height')) != (960, 540):
        print(f"画布是 {m.get('width')}×{m.get('height')}，本脚本只处理 960×540 的母版", file=sys.stderr)
        return 2
    pages = m['pages']
    ids = a.ids.split(',') if a.ids else [f'p{i + 1}' for i in range(len(pages))]
    if len(ids) != len(pages):
        print(f'--ids 有 {len(ids)} 个，母版有 {len(pages)} 页', file=sys.stderr)
        return 2
    bgs = Counter(hex6(p.get('background', '#FFFFFF')) for p in pages)
    texts = Counter(hex6(o['color']) for p in pages for o in p['objects'] if o['type'] == 'text' and o.get('color'))
    roles = {bgs.most_common(1)[0][0]: 'primary'}
    top_text = next((c for c, _ in texts.most_common() if c not in roles), None)
    if top_text:
        roles[top_text] = 'onPrimary'
    layouts, notes = {}, []
    for lid, page in zip(ids, pages):
        layouts[lid], n = convert_page(page, roles, lid)
        notes += n
    draft = {'version': 1, 'deck': {'first': ids[0], 'last': ids[-1]},
             'offline': {'canvas': {'width': 960, 'height': 540, 'unit': 'px'}, 'order': ids, 'layouts': layouts},
             '_roles': {v: k for k, v in roles.items()}, '_notes': notes}
    a.out.write_text(json.dumps(draft, ensure_ascii=False, indent=1), encoding='utf-8')
    if a.report:
        lines = [f'# 母版转换草稿：{a.master.name}', '', f'{len(pages)} 页 → {", ".join(ids)}。配色角色：{json.dumps(draft["_roles"], ensure_ascii=False)}', '']
        for lid, lay in layouts.items():
            reps = [o['label'] for o in lay['objects'] if o['type'] == 'repeat']
            lines.append(f"- {lid}「{lay['title']}」：{len(lay['objects'])} 个对象，字段 {len(lay['fields'])} 个" + (f"，合成 repeat：{', '.join(reps)}" if reps else ''))
        lines += [''] + [f'- 注：{n}' for n in notes]
        a.report.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    print(json.dumps({'layouts': ids, 'roles': draft['_roles'], 'notes': len(notes)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
