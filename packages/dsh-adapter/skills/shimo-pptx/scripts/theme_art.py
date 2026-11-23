#!/usr/bin/env python3
"""Draw a theme illustration for a slide when no suitable photo was found, in the style pack's colours.

  theme_art.py --motif skyline,mountains,river --out images/cover-art.png [--size 1920x1080] [--style default]
               [--palette <name>] [--mood day|dusk|night] [--seed 3] [--note "重庆：山城、两江、夜景"] [--dir images]

Motifs (combine with commas, drawn back to front): sky (always), sun, mountains, hills, skyline, towers, bridge, river,
waves, network, grid, circles, bars, leaves. The picture is a flat illustration, not a photo: it never shows text,
logos or real people. It is recorded in <dir>/sources.json as a generated image (no copyright restriction), so the
credits from find_images.py list it too. Needs Pillow (shipped with python-pptx).
"""
import argparse
import hashlib
import json
import math
import os
import random
import re
import sys

try:
    from PIL import Image, ImageDraw, ImageFilter
except ImportError:
    raise SystemExit('需要 Pillow：用 load_workspace_dependencies 给的 python（带 python-pptx）运行本脚本')

HERE = os.path.dirname(os.path.abspath(__file__))
STYLE_NAME = re.compile(r'^[a-z0-9]+(?:-[a-z0-9]+)*$')
MOTIFS = ['sky', 'sun', 'mountains', 'hills', 'skyline', 'towers', 'bridge', 'river', 'waves', 'network', 'grid', 'circles', 'bars', 'leaves']
ORDER = {m: i for i, m in enumerate(MOTIFS)}


def hex_rgb(value):
    value = value.lstrip('#')
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def style_names():
    root = os.path.join(HERE, '..', 'styles')
    return sorted(d for d in os.listdir(root) if STYLE_NAME.match(d) and os.path.isfile(os.path.join(root, d, 'style.json')))


def load_palette(style, palette):
    if style not in style_names():
        raise SystemExit(f'未知样式 {style}；可用样式：{", ".join(style_names())}')
    with open(os.path.join(HERE, '..', 'styles', style, 'style.json'), encoding='utf-8') as f:
        spec = json.load(f)
    palettes = spec.get('palettes') or {}
    key = palette or spec.get('default_palette')
    found = palettes.get(key) or next((p for p in palettes.values() if p.get('name') == key), None)
    if not found:
        raise SystemExit(f'样式 {style} 没有配色 {key}；可用：{", ".join(palettes)}')
    return {k: hex_rgb(v) for k, v in found.items() if isinstance(v, str) and v.startswith('#')}


class Art:
    def __init__(self, w, h, colors, mood, rnd):
        self.w, self.h, self.rnd = w, h, rnd
        primary = colors.get('primary', (40, 38, 201))
        secondary = colors.get('secondary', mix(primary, (255, 255, 255), 0.6))
        white, black = (255, 255, 255), (12, 14, 28)
        if mood == 'night':
            self.sky_top, self.sky_bottom = mix(primary, black, 0.75), mix(primary, black, 0.35)
            self.far, self.near, self.light = mix(primary, black, 0.5), mix(primary, black, 0.82), (255, 214, 140)
        elif mood == 'dusk':
            self.sky_top, self.sky_bottom = mix(primary, black, 0.25), mix(secondary, (255, 190, 150), 0.55)
            self.far, self.near, self.light = mix(primary, white, 0.25), mix(primary, black, 0.55), (255, 236, 200)
        else:
            self.sky_top, self.sky_bottom = mix(primary, white, 0.45), mix(secondary, white, 0.85)
            self.far, self.near, self.light = mix(primary, white, 0.55), mix(primary, black, 0.15), white
        self.primary, self.secondary, self.mood = primary, secondary, mood
        # Lines and dots: light on a dark sky, the primary colour on a pale one.
        self.ink = self.light if mood != 'day' else mix(primary, black, 0.1)
        self.img = Image.new('RGB', (w, h), self.sky_bottom)
        self.d = ImageDraw.Draw(self.img, 'RGBA')
        self.horizon = int(h * 0.68)

    def layer(self, t):  # 0 = far, 1 = near
        return mix(self.far, self.near, t)

    def sky(self):
        for y in range(self.h):
            self.d.line([(0, y), (self.w, y)], fill=mix(self.sky_top, self.sky_bottom, min(1, y / max(1, self.horizon))))
        if self.mood == 'night':
            for _ in range(int(self.w * self.h / 18000)):
                x, y = self.rnd.random() * self.w, self.rnd.random() * self.horizon * 0.7
                r = self.rnd.choice([1, 1, 1.5, 2])
                self.d.ellipse([x - r, y - r, x + r, y + r], fill=(255, 255, 255, self.rnd.randint(90, 200)))

    def sun(self):
        r = self.h * 0.11
        cx, cy = self.w * self.rnd.uniform(0.62, 0.82), self.h * self.rnd.uniform(0.2, 0.32)
        glow = Image.new('RGBA', self.img.size, (0, 0, 0, 0))
        g = ImageDraw.Draw(glow)
        g.ellipse([cx - r * 2.4, cy - r * 2.4, cx + r * 2.4, cy + r * 2.4], fill=self.light + (60,))
        self.img.paste(glow.filter(ImageFilter.GaussianBlur(r)), (0, 0), glow.filter(ImageFilter.GaussianBlur(r)))
        self.d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=self.light + (235,))

    def ridge(self, base, amp, color, waves=3, rough=0.25):
        pts, phase = [], [self.rnd.uniform(0, math.tau) for _ in range(waves)]
        for x in range(0, self.w + 9, 8):
            y = base
            for i in range(waves):
                y -= amp * math.sin(x / self.w * math.pi * (i + 1.3) + phase[i]) / (i + 1)
            y -= amp * rough * self.rnd.uniform(-1, 1) * 0.3
            pts.append((x, y))
        self.d.polygon([(0, self.h)] + pts + [(self.w, self.h)], fill=color)

    def mountains(self):
        for i, t in enumerate((0.0, 0.3, 0.55)):
            self.ridge(self.horizon - self.h * (0.18 - i * 0.05), self.h * (0.16 - i * 0.03), self.layer(t), waves=4)

    def hills(self):
        for i, t in enumerate((0.35, 0.65)):
            self.ridge(self.horizon + self.h * (0.02 + i * 0.05), self.h * 0.06, self.layer(t), waves=2, rough=0.05)

    def buildings(self, base, tall, t, density):
        x = -self.rnd.randint(0, 40)
        color = self.layer(t)
        while x < self.w:
            bw = self.rnd.randint(int(self.w * 0.018), int(self.w * 0.05))
            bh = self.rnd.uniform(0.25, 1.0) ** 1.6 * tall + tall * 0.12
            top = base - bh
            self.d.rectangle([x, top, x + bw, base + 2], fill=color)
            if self.rnd.random() < 0.18:
                self.d.rectangle([x + bw * 0.42, top - bh * 0.15, x + bw * 0.58, top], fill=color)
            if self.mood != 'day' and t > 0.5:
                for wy in range(int(top + 8), int(base - 6), 14):
                    for wx in range(int(x + 4), int(x + bw - 4), 9):
                        if self.rnd.random() < density:
                            self.d.rectangle([wx, wy, wx + 4, wy + 6], fill=self.light + (self.rnd.randint(120, 230),))
            x += bw + self.rnd.randint(0, int(self.w * 0.01))

    def skyline(self):
        self.buildings(self.horizon, self.h * 0.34, 0.45, 0.15)
        self.buildings(self.horizon + 4, self.h * 0.22, 0.8, 0.35)

    def towers(self):
        color = self.layer(0.7)
        for _ in range(self.rnd.randint(2, 3)):
            cx, bw = self.w * self.rnd.uniform(0.15, 0.85), self.w * 0.03
            top = self.horizon - self.h * self.rnd.uniform(0.42, 0.55)
            self.d.polygon([(cx - bw, self.horizon), (cx - bw * 0.6, top + 30), (cx, top), (cx + bw * 0.6, top + 30), (cx + bw, self.horizon)], fill=color)

    def bridge(self):
        y, color = self.horizon + self.h * 0.03, self.layer(0.9)
        x0, x1 = self.w * 0.05, self.w * 0.95
        self.d.rectangle([x0, y, x1, y + self.h * 0.012], fill=color)
        for px in (self.w * 0.33, self.w * 0.67):
            top = y - self.h * 0.16
            self.d.rectangle([px - 6, top, px + 6, y + self.h * 0.08], fill=color)
            for k in range(1, 9):
                for side in (-1, 1):
                    self.d.line([(px, top + 8), (px + side * k * self.w * 0.018, y)], fill=color + (200,), width=2)
        if self.mood != 'day':
            for lx in range(int(x0), int(x1), 26):
                self.d.ellipse([lx - 2, y - 3, lx + 2, y + 1], fill=self.light + (220,))

    def river(self):
        top = self.horizon + self.h * 0.06
        for y in range(int(top), self.h):
            t = (y - top) / max(1, self.h - top)
            self.d.line([(0, y), (self.w, y)], fill=mix(mix(self.sky_bottom, self.primary, 0.35), self.near, t * 0.6))
        for _ in range(int(self.w / 14)):
            x, y = self.rnd.random() * self.w, self.rnd.uniform(top + 4, self.h)
            ln = self.rnd.uniform(10, 60)
            self.d.line([(x, y), (x + ln, y)], fill=self.light + (self.rnd.randint(30, 110),), width=2)

    def waves(self):
        for i in range(5):
            base = self.h * (0.72 + i * 0.06)
            pts = [(x, base + math.sin(x / self.w * math.tau * 2 + i) * self.h * 0.015) for x in range(0, self.w + 9, 8)]
            self.d.line(pts, fill=mix(self.primary, (255, 255, 255), 0.3 + i * 0.12) + (180,), width=max(2, self.h // 180))

    def network(self):
        pts = [(self.rnd.random() * self.w, self.rnd.random() * self.h * 0.9) for _ in range(38)]
        for i, a in enumerate(pts):
            for b in pts[i + 1:]:
                dist = math.dist(a, b)
                if dist < self.w * 0.16:
                    self.d.line([a, b], fill=self.ink + (int(150 * (1 - dist / (self.w * 0.16))),), width=2)
        for x, y in pts:
            r = self.rnd.uniform(3, 8)
            self.d.ellipse([x - r, y - r, x + r, y + r], fill=self.ink + (220,))

    def grid(self):
        step = max(1, self.w // 24)
        for x in range(0, self.w, step):
            self.d.line([(x, 0), (x, self.h)], fill=self.ink + (30,), width=1)
        for y in range(0, self.h, step):
            self.d.line([(0, y), (self.w, y)], fill=self.ink + (30,), width=1)

    def circles(self):
        for _ in range(7):
            r = self.h * self.rnd.uniform(0.08, 0.35)
            cx, cy = self.rnd.random() * self.w, self.rnd.random() * self.h
            self.d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=mix(self.primary, self.secondary, self.rnd.random()) + (self.rnd.randint(50, 130),))

    def bars(self):
        n, gap = 9, self.w * 0.02
        bw = (self.w * 0.7 - gap * (n - 1)) / n
        for i in range(n):
            bh = self.h * (0.15 + 0.5 * ((i + 1) / n) ** 1.3 * self.rnd.uniform(0.8, 1.05))
            x = self.w * 0.15 + i * (bw + gap)
            self.d.rounded_rectangle([x, self.h * 0.85 - bh, x + bw, self.h * 0.85], radius=bw * 0.2, fill=mix(self.secondary, self.primary, i / n) + (220,))

    def leaves(self):
        for _ in range(26):
            cx, cy, r = self.rnd.random() * self.w, self.rnd.uniform(0.55, 1.0) * self.h, self.h * self.rnd.uniform(0.04, 0.1)
            a = self.rnd.uniform(0, math.pi)
            pts = [(cx + math.cos(a + s * math.pi) * r * (1 if s in (0, 1) else 0.35), cy + math.sin(a + s * math.pi) * r * (1 if s in (0, 1) else 0.35)) for s in (0, 0.5, 1, 1.5)]
            self.d.polygon(pts, fill=mix(self.primary, (60, 160, 90), 0.6) + (200,))


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--motif', required=True, help='逗号分隔：' + ', '.join(MOTIFS))
    p.add_argument('--out', required=True)
    p.add_argument('--size', default='1920x1080')
    p.add_argument('--style', default='default')
    p.add_argument('--palette')
    p.add_argument('--mood', choices=['day', 'dusk', 'night'], default='day')
    p.add_argument('--seed', type=int)
    p.add_argument('--note', default='', help='这张图表现什么，写进来源记录')
    p.add_argument('--dir', default='images', help='sources.json 所在目录（与 find_images.py 共用）')
    a = p.parse_args(argv)
    motifs = [m.strip() for m in a.motif.split(',') if m.strip()]
    unknown = [m for m in motifs if m not in ORDER]
    if unknown:
        raise SystemExit(f'未知母题 {", ".join(unknown)}；可用：{", ".join(MOTIFS)}')
    try:
        w, h = (int(v) for v in a.size.lower().split('x'))
    except ValueError:
        raise SystemExit('--size 写成 宽x高，例如 1920x1080')
    if not (16 <= w <= 4096 and 16 <= h <= 4096):
        raise SystemExit('--size 每边 16–4096 像素')
    if os.path.splitext(a.out)[1].lower() not in ('.png', '.jpg', '.jpeg'):
        raise SystemExit('--out 要以 .png 或 .jpg 结尾')
    seed = a.seed if a.seed is not None else int(hashlib.sha256((a.out + a.motif).encode()).hexdigest()[:8], 16)
    art = Art(w, h, load_palette(a.style, a.palette), a.mood, random.Random(seed))
    for m in sorted(set(['sky'] + motifs), key=ORDER.get):
        getattr(art, m)()
    os.makedirs(os.path.dirname(a.out) or '.', exist_ok=True)
    art.img.save(a.out, optimize=True)
    data = open(a.out, 'rb').read()
    rec = {'file': a.out, 'kind': 'generated', 'title': a.note or ', '.join(motifs), 'note': f"本地生成，母题 {','.join(motifs)}，{a.mood}，seed {seed}",
           'width': w, 'height': h, 'sha256': hashlib.sha256(data).hexdigest(), 'license': '生成图，无版权限制'}
    src = os.path.join(a.dir, 'sources.json')
    try:
        with open(src, encoding='utf-8') as f:
            sources = json.load(f)
    except FileNotFoundError:
        sources = []
    sources = [s for s in sources if s.get('file') != a.out] + [rec]
    os.makedirs(a.dir, exist_ok=True)
    with open(src, 'w', encoding='utf-8') as f:
        json.dump(sources, f, ensure_ascii=False, indent=2)
    print(json.dumps(rec, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
