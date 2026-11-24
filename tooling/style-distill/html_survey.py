#!/usr/bin/env python3
"""Batch preflight of an HTML style library (style-distill tooling; Python standard library only, read-only).

    python3 -I tooling/style-distill/html_survey.py <package> --out survey.md --json survey.json [--pick 8] [--html shortlist.html]

<package> is a folder in the Genspark layout: assets/catalog.json (styles: id, title, category, color, typography,
layoutDescription, pages) and assets/styles/<id>/{layouts.json, slides/*.html, fonts.zh-CN.css}; a single-style package
such as Noir (assets/layouts.json + assets/slides/*.html) is read as one style. Nothing in the package is run: HTML and
CSS are read as text.

Per style it reports what stands between it and a clean PPTX (brand_deck.py can draw text, rectangles, lines, preset
shapes, linear gradients, pictures, native charts and tables): CSS filters, blend modes, backdrop filters, clip paths
and masks, rotated or vertical text, radial/conic gradients, background images, inline SVG diagrams, HTML tables,
regions that are not absolutely positioned; the photo slots (visuals of kind image: all, required, most on a page);
the smallest font (px on the 1920 canvas → pt on our 960 one: px / 2 × 0.75); the fonts and whether their licence is in
the package. It then scores each style for "converts cleanly, business-like, works in Chinese" and picks a varied
shortlist (palette darkness, accent hue, serif or sans, category).

--html writes a shortlist page with each style's reference images; with Pillow importable (run it with a Python that
has Pillow) they are embedded as small JPEGs, otherwise linked. Reference images are the source's own previews: keep the
page out of style packs and out of the plugin.
"""
import argparse
import base64
import colorsys
import io
import json
import re
import sys
from pathlib import Path

PT = 0.75 / 2   # 1920-wide px → pt on the 960 px / 720 pt slide

# Things PPTX (as brand_deck.py draws it) cannot do; weight = how much work or loss each one means.
BLOCKERS = {
    'filter': (re.compile(r'(?<![-\w])filter\s*:\s*(?!none)'), 2, 'CSS filter（照片可改成遮罩，其他只能放弃）'),
    'blend': (re.compile(r'mix-blend-mode\s*:\s*(?!normal)'), 3, '混合模式'),
    'backdrop': (re.compile(r'backdrop-filter\s*:'), 3, '背景模糊'),
    'clip': (re.compile(r'clip-path\s*:|(?<![-\w])mask(-image)?\s*:|-webkit-mask'), 3, '裁切路径 / 遮罩'),
    'rotate': (re.compile(r'transform\s*:[^;"]*(rotate|skew)'), 1, '旋转 / 斜切'),
    'vertical': (re.compile(r'writing-mode\s*:'), 2, '竖排文字'),
    'radial': (re.compile(r'radial-gradient|conic-gradient'), 2, '径向 / 锥形渐变'),
    'bgimage': (re.compile(r'background(-image)?\s*:[^;"]*url\('), 2, '背景图纹理'),
    'svg': (re.compile(r'<svg\b'), 2, '内联 SVG 图示（要改成形状或原生图表）'),
    'table': (re.compile(r'<table\b'), 1, 'HTML 表格（可转原生表格）'),
    'textfx': (re.compile(r'text-shadow\s*:|background-clip\s*:\s*text'), 1, '文字阴影 / 渐变字'),
}
FONT_SIZE = re.compile(r'font-size\s*:\s*([0-9.]+)px')
COLOR = re.compile(r'#([0-9A-Fa-f]{6})\b')
BUSINESS = {'企业战略': 3, '企业销售': 3, '融资': 3, '咨询': 3, '数据与指标': 3, '市场营销': 2, '产品管理': 2, '公共政策': 2, '培训': 2,
            '学术': 1, 'AI 素养': 1, '职业': 1, '设计制作': 1, '课程作业': 0, '生活': 0}
SERIF = re.compile(r'Serif|Garamond|Fraunces|Playfair|Times|Tinos|STIX|Slab|Cormorant|Lora|Merriweather', re.I)
NOT_DISTRIBUTABLE = ('reference times', 'reference arrow', 'times new roman', 'helvetica', 'arial', 'pingfang', 'microsoft yahei')


def read(path):
    try:
        return path.read_text(encoding='utf-8', errors='replace')
    except OSError:
        return ''


def licensed_families(pkg):
    lic = pkg / 'assets' / 'fonts' / 'licenses'
    names = set()
    if lic.is_dir():
        for f in lic.iterdir():
            m = re.match(r'(.+?)-(OFL|LICENSE)\b', f.name)
            if m:
                names.add(norm(m.group(1)))
    # Renamed families whose licence file carries the new name.
    if 'sourceserif4' in names:
        names.add('sourceserifpro')
    return names


def norm(family):
    return re.sub(r'[^a-z0-9]', '', family.lower())


def luminance(hex6):
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def palette_of(html_texts):
    """Page background (the slide container / body), and the most used saturated colour as accent."""
    bgs, counts = [], {}
    for html in html_texts:
        m = re.search(r'class="slide-container"[^>]*style="[^"]*background\s*:\s*#([0-9A-Fa-f]{6})', html) or \
            re.search(r'body\s*\{[^}]*background(?:-color)?\s*:\s*#([0-9A-Fa-f]{6})', html)
        if m:
            bgs.append(m.group(1).upper())
        for c in COLOR.findall(html):
            counts[c.upper()] = counts.get(c.upper(), 0) + 1
    bg = max(set(bgs), key=bgs.count) if bgs else 'FFFFFF'
    accent, best = None, 0
    for c, n in counts.items():
        r, g, b = (int(c[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if s > 0.35 and 0.2 < l < 0.8 and n > best:
            accent, best = c, n
    hue = None
    if accent:
        r, g, b = (int(accent[i:i + 2], 16) / 255 for i in (0, 2, 4))
        hue = round(colorsys.rgb_to_hls(r, g, b)[0] * 360)
    return {'background': '#' + bg, 'dark': luminance(bg) < 0.35, 'accent': '#' + accent if accent else None, 'hue': hue}


def survey_style(sid, entry, folder, layouts, licensed):
    slides = sorted((folder / 'slides').glob('*.html'))
    htmls = [read(p) for p in slides]
    blob = '\n'.join(htmls)
    hits = {}
    for key, (rx, weight, label) in BLOCKERS.items():
        pages = sum(1 for h in htmls if rx.search(h))
        if pages:
            hits[key] = {'pages': pages, 'weight': weight, 'label': label}
    pages = layouts.get('layouts', {})
    photo_all = photo_req = per_page = 0
    texts = 0
    not_absolute = 0
    for page in pages.values():
        imgs = [v for v in page.get('visuals', {}).values() if v.get('kind') == 'image']
        photo_all += len(imgs)
        photo_req += sum(1 for v in imgs if v.get('required'))
        per_page = max(per_page, len(imgs))
        texts += len(page.get('texts', {}))
        for region in page.get('regions', {}).values():
            if 'position:absolute' not in region.get('style', '').replace(' ', ''):
                not_absolute += 1
        for t in page.get('texts', {}).values():
            blob += '\n' + t.get('style', '')
    sizes = [float(x) for x in FONT_SIZE.findall(blob)]
    sizes = [x for x in sizes if x > 0]
    min_px = min(sizes) if sizes else None
    small = sum(1 for x in sizes if x * PT < 9)
    css = read(folder / 'fonts.zh-CN.css') + read(folder / 'fonts.en.css')
    families = sorted(set(re.findall(r"font-family\s*:\s*'([^']+)'", css)))
    fonts = []
    for fam in families:
        n = norm(fam)
        if any(bad in fam.lower() for bad in NOT_DISTRIBUTABLE):
            verdict = '不能分发'
        elif any(n.startswith(lic) or lic.startswith(n) for lic in licensed):
            verdict = '可分发'
        else:
            verdict = '待查'
        fonts.append({'family': fam, 'verdict': verdict})
    # Serif when a Latin serif face is loaded (Noto Serif SC is loaded by every style for Chinese, so it does not count).
    serif = any(SERIF.search(f) for f in families if not f.startswith('Noto Serif'))
    pal = palette_of(htmls)
    hard = sum(h['weight'] * min(h['pages'], 4) for h in hits.values())
    convert = max(0, 10 - hard / 2 - (not_absolute / max(1, len(pages))) * 0.5)
    business = BUSINESS.get(entry.get('category', ''), 1)
    photo_cost = min(3, photo_all / max(1, len(pages)) * 1.5)
    small_cost = 0 if min_px is None else max(0, (9 - min_px * PT)) / 3
    score = round(convert * 1.0 + business * 1.2 - photo_cost - small_cost * 0.5 - (2 if any(f['verdict'] == '不能分发' for f in fonts) else 0), 2)
    refs = []
    for pid in sorted(pages):
        ref = pages[pid].get('reference')
        if ref:
            refs.append(str((folder / ref).resolve()))
    return {
        'id': sid, 'title': entry.get('title', sid), 'name': entry.get('name', ''), 'category': entry.get('category', ''),
        'color': entry.get('color', ''), 'typography': entry.get('typography', ''), 'layout': entry.get('layoutDescription', ''),
        'pages': len(pages), 'texts': texts, 'variants': sorted({p.get('variant') for p in pages.values() if p.get('variant')}),
        'photos': {'slots': photo_all, 'required': photo_req, 'per_slide_max': per_page},
        'min_px': min_px, 'min_pt': round(min_px * PT, 2) if min_px else None, 'under_9pt': small,
        'blockers': hits, 'non_absolute_regions': not_absolute, 'fonts': fonts, 'serif': serif, 'palette': pal,
        'scores': {'convert': round(convert, 2), 'business': business, 'photo_cost': round(photo_cost, 2), 'small_cost': round(small_cost, 2), 'total': score},
        'references': refs,
    }


def load(pkg):
    licensed = licensed_families(pkg)
    catalog = pkg / 'assets' / 'catalog.json'
    out = []
    if catalog.is_file():
        cat = json.loads(read(catalog))
        for entry in cat.get('styles', []):
            folder = pkg / 'assets' / 'styles' / entry['id']
            lay = folder / 'layouts.json'
            if not lay.is_file():
                continue
            out.append(survey_style(entry['id'], entry, folder, json.loads(read(lay)), licensed))
    elif (pkg / 'assets' / 'layouts.json').is_file():
        out.append(survey_style(pkg.name, {'title': pkg.name}, pkg / 'assets', json.loads(read(pkg / 'assets' / 'layouts.json')), licensed))
    else:
        raise SystemExit(f'{pkg} 里没有 assets/catalog.json 或 assets/layouts.json')
    return out


def distance(a, b):
    d = 0.0
    if a['palette']['dark'] != b['palette']['dark']:
        d += 2
    ha, hb = a['palette']['hue'], b['palette']['hue']
    if ha is None or hb is None:
        d += 1 if ha != hb else 0
    else:
        d += min(abs(ha - hb), 360 - abs(ha - hb)) / 60
    d += 1 if a['serif'] != b['serif'] else 0
    d += 1 if a['category'] != b['category'] else 0
    return d


def pick(styles, n, avoid=()):
    """Greedy: best score first, then the best score among those far enough from every pick (and from `avoid`)."""
    pool = sorted((s for s in styles if s['scores']['convert'] >= 6 and s['scores']['business'] >= 1), key=lambda s: -s['scores']['total'])
    chosen = []
    for s in pool:
        if len(chosen) >= n:
            break
        near = [o for o in list(avoid) + chosen if distance(s, o) < 1.5]
        if not near:
            chosen.append(s)
    for s in pool:   # top up if diversity left too few
        if len(chosen) >= n:
            break
        if s not in chosen:
            chosen.append(s)
    return chosen


def report_md(styles, chosen, pkg):
    lines = [f'# HTML 样式库批量预检：{pkg.name}', '', f'{len(styles)} 套。只读源文件，不执行。字号按 1920 宽折算到 960 宽画布（px ÷ 2 × 0.75 = pt）。', '',
             '## 短名单', '', '| ID | 名称 | 类别 | 配色 | 页 | 照片槽（必需/全部/单页最多） | 最小字号 | 拦路项 | 分 |', '|---|---|---|---|---|---|---|---|---|']
    for s in chosen:
        lines.append(row(s))
    lines += ['', '## 全部（按分数）', '', '| ID | 名称 | 类别 | 配色 | 页 | 照片槽（必需/全部/单页最多） | 最小字号 | 拦路项 | 分 |', '|---|---|---|---|---|---|---|---|---|']
    for s in sorted(styles, key=lambda s: -s['scores']['total']):
        lines.append(row(s))
    lines += ['', '## 计分', '', '- convert：10 减去拦路项（权重 × 出现页数，封顶 4 页）的一半，再减去非绝对定位区域的占比；≥ 6 才进短名单。',
              '- business：类别权重（企业战略、销售、融资、咨询、数据 3；营销、产品、政策、培训 2；学术等 1；课程作业、生活 0）× 1.2。',
              '- 扣分：照片多（每页平均 × 1.5，封顶 3）、小字（低于 9pt 的差 ÷ 3 × 0.5）、不能分发的字体 2。',
              '- 短名单在分数之上按「深浅底、主色色相、衬线与否、类别」拉开差距。']
    return '\n'.join(lines) + '\n'


def row(s):
    b = '、'.join(f"{h['label'].split('（')[0]}×{h['pages']}" for h in s['blockers'].values()) or '无'
    p = s['photos']
    pal = s['palette']
    return (f"| {s['id']} | {s['title']} | {s['category']} | {'深' if pal['dark'] else '浅'}底 {pal['background']}，主色 {pal['accent'] or '-'} | {s['pages']} | "
            f"{p['required']}/{p['slots']}/{p['per_slide_max']} | {s['min_pt']}pt（<9pt {s['under_9pt']} 处） | {b} | {s['scores']['total']} |")


def thumb(path, width=300):
    try:
        from PIL import Image
    except ImportError:
        return Path(path).as_uri()
    with Image.open(path) as im:
        im = im.convert('RGB')
        im.thumbnail((width, width * 9 // 16 + 2))
        buf = io.BytesIO()
        im.save(buf, 'JPEG', quality=72)
    return 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode()


def shortlist_html(chosen, pkg):
    cards = []
    for s in chosen:
        refs = s['references']
        sample = [refs[i] for i in sorted({0, len(refs) // 3, (2 * len(refs)) // 3}) if i < len(refs)]
        imgs = ''.join(f'<img src="{thumb(r)}" alt="">' for r in sample)
        p = s['photos']
        b = '、'.join(h['label'].split('（')[0] for h in s['blockers'].values()) or '无'
        cards.append(f'''<section><h2>{s['id']} · {s['title']} <small>{s['category']}</small></h2>
<p>{s['color']}；{s['typography']}。{s['layout']}</p>
<div class="imgs">{imgs}</div>
<p class="facts">{s['pages']} 页 · 照片槽 必需 {p['required']} / 全部 {p['slots']} / 单页最多 {p['per_slide_max']} · 最小字号 {s['min_pt']}pt（要抬到 9pt） · 拦路项：{b} · 分 {s['scores']['total']}</p></section>''')
    return f'''<!doctype html><html lang="zh"><meta charset="utf-8"><title>HTML 样式库短名单</title>
<style>body{{font:14px/1.5 -apple-system,"PingFang SC",sans-serif;margin:24px;max-width:1000px;color:#222}}section{{border-top:1px solid #ddd;padding:12px 0}}
h2{{font-size:17px;margin:0 0 4px}}small{{color:#888;font-weight:400}}.imgs{{display:flex;gap:8px;flex-wrap:wrap}}img{{width:300px;border:1px solid #ccc}}.facts{{color:#555}}</style>
<h1>短名单：{pkg.name}</h1><p>预览图是源包自带的原稿参考图，只供挑选，不进样式包。选好后告诉我编号。</p>
{''.join(cards)}</html>
'''


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('package', type=Path)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--json', type=Path)
    ap.add_argument('--pick', type=int, default=8)
    ap.add_argument('--html', type=Path)
    a = ap.parse_args()
    pkg = a.package.expanduser().resolve()
    styles = load(pkg)
    # The registered default (light, blue-violet #2826C9, sans): a shortlist should not repeat it.
    default_like = {'palette': {'dark': False, 'hue': 241}, 'serif': False, 'category': ''}
    chosen = pick(styles, a.pick, avoid=[default_like])
    a.out.write_text(report_md(styles, chosen, pkg), encoding='utf-8')
    if a.json:
        a.json.write_text(json.dumps({'package': str(pkg), 'styles': styles, 'shortlist': [s['id'] for s in chosen]}, ensure_ascii=False, indent=1), encoding='utf-8')
    if a.html:
        a.html.write_text(shortlist_html(chosen, pkg), encoding='utf-8')
    print(json.dumps({'styles': len(styles), 'shortlist': [s['id'] for s in chosen]}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
