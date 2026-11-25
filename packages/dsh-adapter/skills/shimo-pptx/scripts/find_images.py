#!/usr/bin/env python3
"""Find openly licensed photos for a deck, download them into the task workspace and record where they came from.

  search  --query "Chongqing skyline" [--query ...] [--count 8] [--orientation landscape|portrait|any] [--dir images]
          Asks Openverse (commercial use and modification allowed) and Wikimedia Commons. Keeps CC0, public domain,
          CC BY and CC BY-SA only (no NC, no ND). Writes <dir>/candidates.json and prints the candidates numbered:
          title, creator, licence, size, source page, then tags or description.
  fetch   --pick N [--pick M ...] [--name cover ...] [--dir images] [--max-side 1920]
          Downloads candidates by number into <dir>, re-saved as JPEG with the long side <= max-side, refuses a
          picture already downloaded (same bytes or address), and records it in <dir>/sources.json.
  fit     --in <dir>/x.jpg --box 312x230 [--anchor center|top|bottom|left|right] [--darken 0.45] [--out <dir>/x-left.jpg] [--dir images]
          Crops a picture to the box's proportions (the live editor stretches a picture to its box). --anchor keeps
          that edge (default centre). --darken lays a black veil over it (for white text on a photo). Recorded as made
          from the original.
  credits [--dir images]
          Writes <dir>/sources.md, one line per picture (title, creator, licence, source), and prints it.

Needs the standard library and Pillow (shipped with python-pptx); no API key.
"""
import argparse
import hashlib
import html
import io
import json
import os
import re
import sys
import urllib.parse
import urllib.request

UA = 'shimo-pptx-image-search/0.0.2 (DSH plugin skill)'
TIMEOUT = 25
MAX_BYTES = 10 * 1024 * 1024
DOWNLOAD_MAX = 40 * 1024 * 1024  # an original larger than this is not worth decoding
# Optional version, and after it an optional port ("cn", "de") or "igo" / "international".
LICENSE = re.compile(r'^(?:cc[ -]?)?(by(?:-sa)?)(?:[ -](\d(?:\.\d)?)(?:[ -]([a-z]{2}|igo|international))?)?$')
RESTRICTED = re.compile(r'(?:^|[ -])(?:nc|nd)(?:$|[ -])')  # non-commercial, no-derivatives: never


def get(url, accept='application/json', limit=DOWNLOAD_MAX):
    """One HTTP GET of at most `limit` bytes, with one retry; the retry goes direct when a proxy is set and fails (tests replace this)."""
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept': accept})
    openers = [urllib.request.build_opener()]
    openers.append(urllib.request.build_opener(urllib.request.ProxyHandler({})) if urllib.request.getproxies() else openers[0])
    for i, opener in enumerate(openers):
        try:
            with opener.open(req, timeout=TIMEOUT) as resp:
                data = resp.read(limit + 1)
            if len(data) > limit:
                raise ValueError(f'larger than {limit // 1024 // 1024} MB')
            return data
        except (OSError, ValueError):
            if i == len(openers) - 1:
                raise


def norm_license(name, version=''):
    """'CC0', 'Public domain', 'CC BY 4.0', 'CC BY-SA 3.0'; None for anything else (NC, ND, GFDL only, unknown)."""
    name = re.sub(r'\s+', ' ', (name or '').strip().lower().replace('_', ' '))
    if name in ('cc0', 'cc0 1.0', 'cc-zero'):
        return 'CC0'
    if name in ('pdm', 'public domain', 'pd', 'public domain mark', 'public domain mark 1.0'):
        return 'Public domain'
    if RESTRICTED.search(name):
        return None
    m = LICENSE.match(name)
    if not m:
        return None
    port = m.group(3) if m.group(3) and m.group(3) != 'international' else ''
    return ' '.join(p for p in ('CC', m.group(1).upper(), m.group(2) or version, port.upper()) if p)


def orientation_ok(w, h, want):
    if not w or not h or want == 'any':
        return True
    return w >= h if want == 'landscape' else h >= w


def strip_html(text):
    return re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', '', text or ''))).strip()


def openverse(query, count, want):
    q = urllib.parse.urlencode({'q': query, 'page_size': min(count * 2, 20), 'license_type': 'commercial,modification', 'mature': 'false'})
    data = json.loads(get(f'https://api.openverse.org/v1/images/?{q}'))
    out = []
    for r in data.get('results') or []:
        lic = norm_license(r.get('license'), r.get('license_version') or '')
        if not lic or not r.get('url') or not orientation_ok(r.get('width'), r.get('height'), want):
            continue
        out.append({'source': 'openverse/' + (r.get('source') or ''), 'title': r.get('title') or '', 'creator': (r.get('creator') or '')[:120],
                    'license': lic, 'license_url': r.get('license_url') or '', 'width': r.get('width'), 'height': r.get('height'),
                    'page': r.get('foreign_landing_url') or '', 'url': r['url'],
                    'about': ', '.join(t.get('name', '') for t in (r.get('tags') or [])[:8])})
    return out


def commons(query, count, want):
    q = urllib.parse.urlencode({'action': 'query', 'format': 'json', 'generator': 'search', 'gsrsearch': f'{query} filetype:bitmap',
                                'gsrnamespace': 6, 'gsrlimit': min(count * 2, 20), 'prop': 'imageinfo',
                                'iiprop': 'url|size|mime|extmetadata', 'iiurlwidth': 1920})
    data = json.loads(get(f'https://commons.wikimedia.org/w/api.php?{q}'))
    out = []
    for page in sorted(((data.get('query') or {}).get('pages') or {}).values(), key=lambda p: p.get('index', 0)):
        info = (page.get('imageinfo') or [{}])[0]
        meta = info.get('extmetadata') or {}
        lic = norm_license(strip_html((meta.get('LicenseShortName') or {}).get('value')))
        url = info.get('thumburl') or info.get('url')
        if not lic or not url or info.get('mime') not in ('image/jpeg', 'image/png') or not orientation_ok(info.get('width'), info.get('height'), want):
            continue
        out.append({'source': 'wikimedia-commons', 'title': page.get('title', '').removeprefix('File:'),
                    'creator': strip_html((meta.get('Artist') or {}).get('value'))[:120], 'license': lic,
                    'license_url': strip_html((meta.get('LicenseUrl') or {}).get('value')), 'width': info.get('width'), 'height': info.get('height'),
                    'page': info.get('descriptionurl') or '', 'url': url,
                    'about': strip_html((meta.get('ImageDescription') or {}).get('value'))[:160]})
    return out


SEARCHERS = (('openverse', openverse), ('commons', commons))


def load(path, default):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def save(path, data):
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def commons_key(c):
    """Openverse also indexes Wikimedia Commons: the same file then comes back from both, under the same file title."""
    if c['source'] not in ('openverse/wikimedia', 'wikimedia-commons'):
        return None
    return re.sub(r'[\W_]+', '', re.sub(r'\.(jpe?g|png)$', '', c['title'] or '', flags=re.I).lower()) or None


def cmd_search(a):
    found, errors, urls, keys, calls = [], [], set(), set(), 0
    for query in a.query:
        for name, fn in SEARCHERS:
            calls += 1
            try:
                for c in fn(query, a.count, a.orientation):
                    key = commons_key(c)
                    if c['url'] in urls or (key and key in keys):
                        continue
                    urls.add(c['url']); keys.add(key); c['query'] = query; found.append(c)
            except Exception as e:  # one source down must not hide the other
                errors.append(f'{name}「{query}」：{e}')
    save(os.path.join(a.dir, 'candidates.json'), found)
    for i, c in enumerate(found, 1):
        print(f"{i}. {c['title'][:80]} | {c['creator'][:40] or '未署名'} | {c['license']} | {c['width']}x{c['height']} | {c['page']}")
        if c.get('about'):
            print(f"   {c['about'][:160]}")
    for e in errors:
        print('warning:', e, file=sys.stderr)
    if found:
        print('（编号只对这次 search 有效：下一次 search 会换掉候选。先 fetch 这次要的，再搜下一批。）')
        return 0
    if len(errors) == calls:
        print('图库都连不上（见上面的 warning），不用再换关键词：直接用 theme_art.py 生成配图。')
        return 1
    print('没有找到可商用、可修改授权的图片：换英文、更具体或更宽泛的关键词再搜；仍没有就用 theme_art.py 生成配图。')
    return 0


def slug(text):
    s = re.sub(r'[^\w-]+', '-', text, flags=re.U).strip('-')
    return s[:40] or 'image'


def free_path(directory, base, ext):
    path, i = os.path.join(directory, base + ext), 2
    while os.path.exists(path):
        path, i = os.path.join(directory, f'{base}-{i}{ext}'), i + 1
    return path


def image_module():
    try:
        from PIL import Image
        return Image
    except ImportError:
        raise SystemExit('需要 Pillow：用 load_workspace_dependencies 给的 python（带 python-pptx）运行本脚本')


def cmd_fetch(a):
    Image = image_module()
    cands = load(os.path.join(a.dir, 'candidates.json'), [])
    sources_path = os.path.join(a.dir, 'sources.json')
    sources = load(sources_path, [])
    names, results = a.name or [], []
    if len(names) > len(a.pick):
        raise SystemExit('--name 比 --pick 多；每个 --name 按顺序对应一个 --pick')
    for n, pick in enumerate(a.pick):
        if not 1 <= pick <= len(cands):
            raise SystemExit(f'--pick {pick} 不在候选 1–{len(cands)} 里；先运行 search')
        c = cands[pick - 1]
        try:
            raw = get(c['url'], accept='image/*')
            img = Image.open(io.BytesIO(raw)).convert('RGB')
        except Exception as e:
            results.append({'pick': pick, 'error': f'下载或解码失败：{e}；换一张'}); continue
        digest = hashlib.sha256(raw).hexdigest()
        dup = next((s for s in sources if s.get('download_sha256') == digest or (s.get('url') and s.get('url') == c['url'])), None)
        if dup:
            results.append({'pick': pick, 'error': f"与已下载的 {dup['file']} 是同一张图；同一张图整份只用一次"}); continue
        img.thumbnail((a.max_side, a.max_side))
        os.makedirs(a.dir, exist_ok=True)
        path = free_path(a.dir, slug(names[n] if n < len(names) else c['title'].rsplit('.', 1)[0]), '.jpg')
        img.save(path, 'JPEG', quality=88, optimize=True)
        if os.path.getsize(path) > MAX_BYTES:
            os.remove(path)
            results.append({'pick': pick, 'error': '图片超过 10 MB，换一张或调小 --max-side'}); continue
        rec = {'file': path, 'kind': 'photo', 'title': c['title'], 'creator': c['creator'], 'license': c['license'], 'license_url': c['license_url'],
               'page': c['page'], 'url': c['url'], 'source': c['source'], 'width': img.width, 'height': img.height,
               'sha256': file_sha256(path), 'download_sha256': digest}
        sources.append(rec); results.append(rec)
    save(sources_path, sources)
    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0 if all('error' not in r for r in results) else 1


def file_sha256(path):
    with open(path, 'rb') as f:
        return hashlib.sha256(f.read()).hexdigest()


def parse_box(text):
    m = re.fullmatch(r'\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*', text or '')
    if not m or float(m.group(1)) <= 0 or float(m.group(2)) <= 0:
        raise SystemExit('--box 写成 宽x高，例如 312x230（单位不限，只看比例）')
    return float(m.group(1)), float(m.group(2))


def same_file(a, b):
    try:
        return os.path.samefile(a, b)
    except OSError:
        return os.path.abspath(a) == os.path.abspath(b)


def crop_offset(extra, anchor, start, end):
    """Where to keep `extra` pixels when trimming to the box. `start`/`end` are the edges along that axis."""
    if extra <= 0:
        return 0
    if anchor == start:
        return 0
    if anchor == end:
        return extra
    return extra // 2


def cmd_fit(a):
    Image = image_module()
    bw, bh = parse_box(a.box)
    if not 0 <= a.darken < 1:
        raise SystemExit('--darken 取 0 到 1 之间，例如 0.45')
    anchor = a.anchor
    img = Image.open(a.input).convert('RGB')
    w, h = img.size
    if w / h > bw / bh:
        nw = round(h * bw / bh)
        left = crop_offset(w - nw, anchor, 'left', 'right')
        img = img.crop((left, 0, left + nw, h))
    else:
        nh = round(w * bh / bw)
        top = crop_offset(h - nh, anchor, 'top', 'bottom')
        img = img.crop((0, top, w, top + nh))
    if a.darken:
        img = Image.blend(img, Image.new('RGB', img.size, (0, 0, 0)), a.darken)
    base, _ = os.path.splitext(a.input)
    out = a.out or f'{base}-{round(bw)}x{round(bh)}{"-dark" if a.darken else ""}.jpg'
    sources_path = os.path.join(a.dir, 'sources.json')
    sources = load(sources_path, [])
    origin = next((s for s in sources if same_file(s['file'], a.input)), None)
    # One picture once per deck: a second crop of the same original goes elsewhere in the deck, so it is refused;
    # fitting again to the same output (a corrected box) replaces it.
    other = [f for f in (origin or {}).get('fitted', []) if os.path.abspath(f) != os.path.abspath(out)]
    if other:
        raise SystemExit(f'{a.input} 已裁成 {other[0]}；同一张图整份只用一次，换一张图（要改尺寸或 --anchor 就用同一个 --out 重裁）')
    img.save(out, 'JPEG', quality=88, optimize=True)
    if origin is not None:
        origin['fitted'] = [out]
        origin['fitted_sha256'] = [file_sha256(out)]
        save(sources_path, sources)
    print(json.dumps({'file': out, 'from': a.input, 'width': img.width, 'height': img.height, 'darken': a.darken,
                      'anchor': anchor, 'recorded': origin is not None}, ensure_ascii=False))
    if origin is None:
        print(f'warning: {a.input} 不在 {sources_path} 里，来源要自己写进交付说明', file=sys.stderr)
    return 0


def credit_line(s):
    name = os.path.basename(s['file'])
    if s.get('kind') == 'generated':
        return f"{name}：按主题生成的示意配图（{s.get('title') or '本地生成'}），无版权限制"
    return f"{name}：「{s.get('title')}」，作者 {s.get('creator') or '未署名'}，{s.get('license')}（{s.get('license_url') or '-'}），来源 {s.get('page')}"


def cmd_credits(a):
    sources = load(os.path.join(a.dir, 'sources.json'), [])
    lines = [credit_line(s) for s in sources]
    os.makedirs(a.dir, exist_ok=True)
    with open(os.path.join(a.dir, 'sources.md'), 'w', encoding='utf-8') as f:
        f.write('# 图片来源\n\n' + ''.join(f'- {line}\n' for line in lines))
    print('\n'.join(lines) or '（没有记录的图片）')
    return 0


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest='cmd', required=True)
    s = sub.add_parser('search')
    s.add_argument('--query', action='append', required=True)
    s.add_argument('--count', type=int, default=8)
    s.add_argument('--orientation', choices=['landscape', 'portrait', 'any'], default='any')
    s.add_argument('--dir', default='images')
    f = sub.add_parser('fetch')
    f.add_argument('--pick', type=int, action='append', required=True)
    f.add_argument('--name', action='append')
    f.add_argument('--dir', default='images')
    f.add_argument('--max-side', type=int, default=1920)
    t = sub.add_parser('fit')
    t.add_argument('--in', dest='input', required=True)
    t.add_argument('--box', required=True)
    t.add_argument('--anchor', choices=['center', 'top', 'bottom', 'left', 'right'], default='center')
    t.add_argument('--darken', type=float, default=0.0)
    t.add_argument('--out')
    t.add_argument('--dir', default='images')
    c = sub.add_parser('credits')
    c.add_argument('--dir', default='images')
    a = p.parse_args(argv)
    return {'search': cmd_search, 'fetch': cmd_fetch, 'fit': cmd_fit, 'credits': cmd_credits}[a.cmd](a)


if __name__ == '__main__':
    sys.exit(main())
