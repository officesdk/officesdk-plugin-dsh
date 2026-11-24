#!/usr/bin/env python3
"""Build an editable brand-style PPTX from a deck JSON spec (shimo-pptx, offline mode).

Usage:
  python brand_deck.py build --spec deck.json --out deck.pptx [--style <name>]
  python brand_deck.py layouts [--style <name>]   # list layouts and their fields

Page types are not code: each style pack's styles/<name>/layouts.json lists, per page type, the objects to draw
(text boxes, shapes, pictures, charts, tables) with their geometry on a 960x540 px canvas (1 px = 0.75 pt), and this
script draws them as native PowerPoint objects, one text box per block of text. Format: tooling/style-distill/pack-format.md
in the source repository.
"""
import argparse
import ast
import datetime
import hashlib
import importlib.util
import json
import operator
import re
import sys
import unicodedata
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

ROOT = Path(__file__).resolve().parents[1]
STYLES = ROOT / 'styles'
STYLE_NAME = re.compile(r'^[a-z0-9]+(?:-[a-z0-9]+)*$')
DEFAULT_STYLE = 'default'
PACK_FILES = ('style.json', 'template.pptx', 'layouts.json')


def set_styles_dir(path):
    """Read packs from another folder (style-distill staging); the default is styles/ next to this script."""
    global STYLES
    STYLES = Path(path).resolve()
    if not STYLES.is_dir():
        raise ValueError(f'样式目录 {path} 不存在')


def style_names():
    """Style packs under styles/: each folder with style.json, template.pptx and layouts.json (the Host offers the same ones)."""
    return sorted(p.name for p in STYLES.iterdir() if STYLE_NAME.match(p.name) and all((p / f).is_file() for f in PACK_FILES))


def use_style(name):
    """Load styles/<name>/style.json into the module values the layouts read."""
    global STYLE, STYLE_ID, EMU, PTPX, CJK_FONT, LATIN, LATIN_LIGHT, DESC, LAYOUT_DEFS, LAYOUTS
    if name not in style_names():
        raise ValueError(f'未知样式 {name}，可选 {", ".join(style_names()) or "无"}')
    STYLE = json.loads((STYLES / name / 'style.json').read_text(encoding='utf-8'))
    STYLE_ID = name
    EMU = STYLE['canvas']['emu_per_px']
    PTPX = STYLE['canvas']['pt_per_px']
    CJK_FONT = STYLE['fonts']['cjk']
    LATIN = STYLE['fonts']['latin']
    LATIN_LIGHT = STYLE['fonts']['latin_light']
    DESC, LAYOUT_DEFS, LAYOUTS = load_layouts(name)


ARROW = '→'
# Chart text sizes in px (a layout's chart object may override any of them).
CHART_FONTS = {'base': 13, 'labels': 13, 'doughnut_labels': 10, 'value_axis': 11, 'category_axis': 12}


def _find_images():
    """find_images.py beside this script, loaded by path (python -I does not put the script folder on sys.path)."""
    spec = importlib.util.spec_from_file_location('shimo_find_images', Path(__file__).with_name('find_images.py'))
    module = importlib.util.module_from_spec(spec)
    # No __pycache__ inside the installed skill folder.
    was, sys.dont_write_bytecode = sys.dont_write_bytecode, True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = was
    return module


def image_credit(path):
    """The credit line for a picture recorded by find_images.py / theme_art.py in the sources.json beside it, else None."""
    try:
        records = json.loads((path.parent / 'sources.json').read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return None
    records = [rec for rec in records if isinstance(rec, dict)] if isinstance(records, list) else []
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    # The file itself (or a crop fit made of it), by content; a record without hashes by its file name.
    match = next((rec for rec in records if digest == rec.get('sha256') or digest in (rec.get('fitted_sha256') or [])), None)
    if match is None:
        match = next((rec for rec in records if not rec.get('sha256')
                      and path.name in [Path(str(rec.get('file', ''))).name] + [Path(str(f)).name for f in rec.get('fitted') or []]), None)
    return _find_images().credit_line({**match, 'file': str(path)}) if match else None

# ---------- layout description (styles/<name>/layouts.json) ----------
# Each page type is a list of objects in styles/<name>/layouts.json ("offline" part); this script draws them, so a new
# style is a new pack, not new code. Values may be expressions in a small, whitelisted subset of Python expression
# syntax (src/host/style-expr.mjs checks the same grammar on the Host side); they are interpreted node by node, never
# passed to eval().

_EXPR_CACHE = {}
_BINOPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv,
           ast.FloorDiv: operator.floordiv, ast.Mod: operator.mod}
_CMPOPS = {ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt, ast.LtE: operator.le, ast.Gt: operator.gt,
           ast.GtE: operator.ge, ast.In: lambda a, b: a in b, ast.NotIn: lambda a, b: a not in b}
EXPR_FUNCTIONS = ('len', 'str', 'int', 'float', 'min', 'max', 'abs', 'round', 'lines', 'line', 'tw', 'fit', 'pad2', 'fmt', 'year', 'pct')


def parse_expr(source):
    """The AST of an expression string, after the whitelist check (ValueError names what is not allowed)."""
    tree = _EXPR_CACHE.get(source)
    if tree is None:
        try:
            tree = ast.parse(source, mode='eval').body
        except SyntaxError as e:
            raise ValueError(f'表达式 {source!r} 语法错误：{e.msg}') from None
        _check_expr(tree, source)
        _EXPR_CACHE[source] = tree
    return tree


def _check_expr(node, source):
    ok = (ast.Constant, ast.Name, ast.List, ast.Tuple, ast.Dict, ast.BinOp, ast.UnaryOp, ast.BoolOp, ast.Compare,
          ast.IfExp, ast.Subscript, ast.Slice, ast.Call, ast.Load, ast.And, ast.Or, ast.Not, ast.USub, ast.UAdd,
          ast.Attribute, *_BINOPS, *_CMPOPS)
    for sub in ast.walk(node):
        if not isinstance(sub, ok):
            raise ValueError(f'表达式 {source!r} 含不允许的语法 {type(sub).__name__}')
        if isinstance(sub, ast.Constant) and not isinstance(sub.value, (str, int, float, bool, type(None))):
            raise ValueError(f'表达式 {source!r} 含不允许的常量')
        if isinstance(sub, ast.Dict) and any(k is None for k in sub.keys):
            raise ValueError(f'表达式 {source!r} 的字典不能用星号展开')
        if isinstance(sub, ast.Attribute) and sub.attr != 'get':
            raise ValueError(f'表达式 {source!r} 只允许 .get()，不允许 .{sub.attr}')
        if isinstance(sub, ast.Call):
            if sub.keywords or any(isinstance(a, ast.Starred) for a in sub.args):
                raise ValueError(f'表达式 {source!r} 的调用不能用关键字参数或星号展开')
            fn = sub.func
            if isinstance(fn, ast.Name) and fn.id not in EXPR_FUNCTIONS:
                raise ValueError(f'表达式 {source!r} 调用了不允许的函数 {fn.id}')
            if not isinstance(fn, (ast.Name, ast.Attribute)):
                raise ValueError(f'表达式 {source!r} 只能调用内置函数或 .get()')
            if isinstance(fn, ast.Attribute) and not 1 <= len(sub.args) <= 2:
                raise ValueError(f'表达式 {source!r} 的 .get() 带 1–2 个参数')
        if isinstance(sub, ast.Attribute):
            parent_ok = any(isinstance(c, ast.Call) and c.func is sub for c in ast.walk(node))
            if not parent_ok:
                raise ValueError(f'表达式 {source!r} 的 .get 必须直接调用')


class Scope:
    """Names an expression can see: fixed vars, then lazily evaluated `let` names (each computed once, at first use, so
    side effects such as fit() warnings happen in drawing order), then the parent scope."""

    def __init__(self, deck, vars=None, lets=None, parent=None):
        self.deck, self.vars, self.lets, self.parent = deck, vars or {}, lets or {}, parent
        self.memo, self.busy = {}, set()

    def get(self, name):
        scope = self
        while scope is not None:
            if name in scope.vars:
                return scope.vars[name]
            if name in scope.lets:
                if name not in scope.memo:
                    if name in scope.busy:
                        raise ValueError(f'let {name} 循环引用自己')
                    scope.busy.add(name)
                    scope.memo[name] = scope.deck.ev(scope.lets[name], scope)
                    scope.busy.discard(name)
                return scope.memo[name]
            scope = scope.parent
        raise ValueError(f'{self.deck.where}：表达式用到未定义的名字 {name}')


def _eval(node, scope, fns):
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.Name):
        return scope.get(node.id)
    if isinstance(node, ast.List):
        return [_eval(e, scope, fns) for e in node.elts]
    if isinstance(node, ast.Tuple):
        return tuple(_eval(e, scope, fns) for e in node.elts)
    if isinstance(node, ast.Dict):
        return {_eval(k, scope, fns): _eval(v, scope, fns) for k, v in zip(node.keys, node.values)}
    if isinstance(node, ast.BinOp):
        return _BINOPS[type(node.op)](_eval(node.left, scope, fns), _eval(node.right, scope, fns))
    if isinstance(node, ast.UnaryOp):
        v = _eval(node.operand, scope, fns)
        return not v if isinstance(node.op, ast.Not) else (-v if isinstance(node.op, ast.USub) else +v)
    if isinstance(node, ast.BoolOp):
        is_and = isinstance(node.op, ast.And)
        v = None
        for e in node.values:
            v = _eval(e, scope, fns)
            if (not v) if is_and else v:
                return v
        return v
    if isinstance(node, ast.Compare):
        left = _eval(node.left, scope, fns)
        for op, right_node in zip(node.ops, node.comparators):
            right = _eval(right_node, scope, fns)
            if not _CMPOPS[type(op)](left, right):
                return False
            left = right
        return True
    if isinstance(node, ast.IfExp):
        return _eval(node.body, scope, fns) if _eval(node.test, scope, fns) else _eval(node.orelse, scope, fns)
    if isinstance(node, ast.Subscript):
        value = _eval(node.value, scope, fns)
        if isinstance(node.slice, ast.Slice):
            part = [None if p is None else _eval(p, scope, fns) for p in (node.slice.lower, node.slice.upper, node.slice.step)]
            return value[slice(*part)]
        return value[_eval(node.slice, scope, fns)]
    if isinstance(node, ast.Call):
        args = [_eval(a, scope, fns) for a in node.args]
        if isinstance(node.func, ast.Attribute):
            owner = _eval(node.func.value, scope, fns)
            if not isinstance(owner, dict) or not 1 <= len(args) <= 2:
                raise ValueError(f'{scope.deck.where}：.get() 只能用在对象上，且带 1–2 个参数')
            return owner.get(*args)
        return fns[node.func.id](*args)
    raise ValueError(f'不支持的表达式节点 {type(node).__name__}')


def _resolve_layouts(desc):
    """{id: layout} with `extends` merged (vars merged; other keys from the child win)."""
    raw = desc['offline']['layouts']
    out = {}
    for lid, layout in raw.items():
        base = raw.get(layout['extends']) if layout.get('extends') else None
        if layout.get('extends') and (base is None or base.get('extends')):
            raise ValueError(f'版式 {lid} 的 extends={layout["extends"]!r} 不存在或也是变体')
        merged = {**(base or {}), **layout}
        merged['vars'] = {**(base or {}).get('vars', {}), **layout.get('vars', {})}
        merged['let'] = {**(base or {}).get('let', {}), **layout.get('let', {})}
        merged.setdefault('fields', {})
        out[lid] = merged
    return out


def load_layouts(name):
    """styles/<name>/layouts.json → (description, resolved layouts, {id: description line} in display order)."""
    path = STYLES / name / 'layouts.json'
    try:
        desc = json.loads(path.read_text(encoding='utf-8'))
    except OSError:
        raise ValueError(f'样式 {name} 缺少 layouts.json') from None
    if desc.get('version') != 1:
        raise ValueError(f'样式 {name} 的 layouts.json version 应为 1')
    layouts = _resolve_layouts(desc)
    order = desc['offline'].get('order') or list(layouts)
    return desc, layouts, {lid: layouts[lid].get('description', lid) for lid in order}


use_style(DEFAULT_STYLE)


def px(v):
    return Emu(int(round(v * EMU)))


def rgb(hex_color):
    return RGBColor.from_string(hex_color.lstrip('#').upper())


def is_cjk_line(text):
    return any(ord(c) > 0x7F for c in text)


def char_width(c, cjk_line):
    if unicodedata.east_asian_width(c) in ('W', 'F'):
        return 1.0
    if c == ' ':
        return 0.28
    if c in ARROW:
        return 1.0
    if c.isdigit():
        return 0.56 if cjk_line else 0.62
    if c.isupper():
        return 0.62 if cjk_line else 0.70
    if c.islower():
        return 0.52 if cjk_line else 0.58
    return 0.32


def text_width(text, size):
    """Approximate rendered width in px (no font files are bundled)."""
    cjk = is_cjk_line(text)
    return sum(char_width(c, cjk) for c in text) * size


def lines_of(text):
    if isinstance(text, list):
        return [str(x) for x in text]
    return str(text or '').split('\n')


def font_face(role, cjk, weight=400):
    """Typeface of a font role in style.json fonts.roles: {latin, cjk, latin_light?}; a Chinese line takes cjk."""
    roles = STYLE['fonts'].get('roles') or {}
    if role not in roles:
        raise ValueError(f'字体角色 {role} 不在 style.json fonts.roles 里（可选 {", ".join(roles) or "无"}）')
    spec = roles[role]
    if cjk:
        return spec['cjk']
    return spec.get('latin_light', spec['latin']) if weight == 300 else spec['latin']


class Deck:
    def __init__(self, spec, base_dir):
        self.spec = spec
        self.base = base_dir
        self.warnings = []
        pid = spec.get('palette', STYLE['default_palette'])
        if pid not in STYLE['palettes']:
            raise ValueError(f'未知配色 {pid}，可选 {", ".join(STYLE["palettes"])}')
        self.P = STYLE['palettes'][pid]
        self.palette_id = pid
        self.prs = Presentation()
        self.prs.slide_width = px(960)
        self.prs.slide_height = px(540)
        self.blank = self.prs.slide_layouts[6]
        self.used_images = {}
        self.slide_images = []
        self.image_credits = []
        self._color_warned = set()
        self.where = '封面'
        self._fns = self.fns()
        brand = spec.get('brand') or {}
        self.website = brand.get('website', '')
        self.logo = self.resolve(brand.get('logo'))
        self.logo_dark = self.resolve(brand.get('logo_on_dark')) or self.logo
        self.logo_reason = brand.get('logo_missing_reason', '')
        if not self.logo:
            self.warnings.append('未提供 Logo 文件：封面、内容页与结束页均不放 Logo（不用文字仿 Logo）。'
                                 + (f'原因：{self.logo_reason}' if self.logo_reason else '请在 brand.logo_missing_reason 记录原因。'))
        self.year = str(spec.get('year') or datetime.date.today().year)

    # ---------- resources ----------
    def resolve(self, path):
        if not path:
            return None
        p = Path(path)
        if not p.is_absolute():
            p = (self.base / p).resolve()
        if not p.exists():
            raise ValueError(f'找不到文件 {path}')
        if p.suffix.lower() == '.svg':
            raise ValueError(f'{path} 是 SVG；python-pptx 不能插入 SVG，请先提供 PNG/JPG 版本')
        return p

    def photo(self, path, where):
        p = self.resolve(path)
        if p is None:
            return None
        digest = hashlib.sha256(p.read_bytes()).hexdigest()
        if digest in self.used_images:
            raise ValueError(f'{where} 的图片与 {self.used_images[digest]} 是同一张图；同一张普通图片整份只能用一次')
        self.used_images[digest] = where
        self.slide_images.append(p)
        return p

    def take(self, lst, n, field):
        lst = list(lst or [])
        if len(lst) > n:
            dropped = '、'.join(str(x if not isinstance(x, dict) else (x.get('label') or x.get('heading') or x.get('value') or ''))[:10] for x in lst[n:])
            self.warnings.append(f'{self.where}：{field} 有 {len(lst)} 项，本页型最多 {n} 项，已截掉：{dropped}。请精简或拆页。')
        return lst[:n]

    # ---------- primitives ----------
    def new_slide(self, bg):
        s = self.prs.slides.add_slide(self.blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = rgb(bg)
        return s

    @staticmethod
    def _plain(shape):
        style = shape._element.find(qn('p:style'))
        if style is not None:
            shape._element.remove(style)

    def rect(self, s, x, y, w, h, fill, alpha=None, name=None, outline=None, gradient=None, shape=MSO_SHAPE.RECTANGLE):
        sh = s.shapes.add_shape(shape, px(x), px(y), px(w), px(h))
        self._plain(sh)
        if gradient:
            self._gradient(sh, gradient)
        elif fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(fill)
            if alpha is not None:
                clr = sh.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
                a = etree.SubElement(clr, qn('a:alpha'))
                a.set('val', str(int(alpha * 100000)))
        else:
            sh.fill.background()
        if outline:
            sh.line.color.rgb = rgb(outline)
            sh.line.width = Pt(0.75)
        else:
            sh.line.fill.background()
        if name:
            sh.name = name
        return sh

    @staticmethod
    def _gradient(sh, gradient):
        """Linear gradient fill: gradient = (angle in degrees, [(pos 0–1, '#RRGGBB', alpha 0–1), ...]); raw XML (python-pptx has no stop alpha)."""
        angle, stops = gradient
        sppr = sh.fill._xPr
        for tag in ('a:solidFill', 'a:noFill', 'a:gradFill'):
            for el in sppr.findall(qn(tag)):
                sppr.remove(el)
        grad = etree.Element(qn('a:gradFill'), rotWithShape='1')
        lst = etree.SubElement(grad, qn('a:gsLst'))
        for pos, color, alpha in stops:
            gs = etree.SubElement(lst, qn('a:gs'), pos=str(int(round(pos * 100000))))
            clr = etree.SubElement(gs, qn('a:srgbClr'), val=color.lstrip('#').upper())
            if alpha is not None and alpha < 1:
                etree.SubElement(clr, qn('a:alpha'), val=str(int(round(alpha * 100000))))
        # OOXML angles run clockwise from "left to right" in 1/60000 degree; CSS 180deg (top → bottom) is OOXML 90°.
        etree.SubElement(grad, qn('a:lin'), ang=str(int(round(((angle - 90) % 360) * 60000))), scaled='0')
        geom = sppr.find(qn('a:prstGeom'))
        geom.addnext(grad)

    def hline(self, s, x, y, w, color):
        return self.rect(s, x, y, w, 1, color)

    def vline(self, s, x, y, h, color):
        return self.rect(s, x, y, 1, h, color)

    @staticmethod
    def _fonts(run, face):
        rpr = run._r.get_or_add_rPr()
        for tag in ('a:latin', 'a:ea', 'a:cs'):
            el = rpr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rpr, qn(tag))
            el.set('typeface', face)

    def text(self, s, x, y, w, content, size, color, weight=400, lh=1.13, align='left',
             wrap=True, max_h=None, name=None, anchor='top', font=None, italic=False, spacing=None, bold=False, para_colors=None):
        """One text box; each '\n' line is a paragraph. Returns (shape, line_tops)."""
        ls = lines_of(content)
        h = max_h if max_h else size * lh * max(1, len(ls)) + 4
        tb = s.shapes.add_textbox(px(x), px(y), px(w), px(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE, 'bottom': MSO_ANCHOR.BOTTOM}[anchor]
        for i, line in enumerate(ls):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[align]
            p.line_spacing = Pt(size * lh * PTPX)
            p.space_before = p.space_after = Pt(0)
            r = p.add_run()
            r.text = line
            f = r.font
            f.size = Pt(size * PTPX)
            f.bold = bool(bold)
            f.color.rgb = rgb(para_colors[i] if para_colors and i < len(para_colors) else color)
            cjk = is_cjk_line(line)
            if font is None:
                face = CJK_FONT if cjk else (LATIN_LIGHT if weight == 300 else LATIN)
            else:
                face = font_face(font, cjk, weight)
            self._fonts(r, face)
            if italic and not cjk:   # no synthetic italic for Chinese (the source styles drop it too)
                f.italic = True
            if spacing:
                r._r.get_or_add_rPr().set('spc', str(int(round(spacing * 100))))   # 1/100 pt
        if name:
            tb.name = name
        if wrap:
            est = sum(max(1, -(-int(text_width(l, size)) // max(1, int(w)))) for l in ls)
            if max_h and est * size * lh > max_h + size * 0.3:
                self.warnings.append(f'{self.where}：文字「{ls[0][:12]}…」估计超出文本框高度，请精简。')
        return tb, [y + i * size * lh for i in range(len(ls))]

    def fit_size(self, content, size, max_w, min_ratio=0.62):
        widest = max((text_width(l, size) for l in lines_of(content)), default=0)
        if widest <= max_w:
            return size
        new = max(size * min_ratio, size * max_w / widest)
        if new == size * min_ratio:
            self.warnings.append(f'{self.where}：标题「{lines_of(content)[0][:10]}…」过长，已缩到下限仍可能溢出，请精简。')
        return new

    def highlight_title(self, s, x, y, content, size, fill, color, row_h=(70, 73), pad=None, weight=300,
                        max_w=520, text_right=900, base=60.0, pad_ratio=0.09, dy_ratio=0.02):
        """Big title where every line sits on its own colour block sized to the line (the defaults are default's)."""
        ls = lines_of(content)
        size = self.fit_size(ls, size, max_w)
        pad = size * pad_ratio if pad is None else pad
        scale = size / base
        top = y
        for i, line in enumerate(ls):
            hh = (row_h[0] if i == 0 else row_h[1]) * scale
            self.rect(s, x, top, text_width(line, size) + pad * 2, hh, fill)
            top += hh
        self.text(s, x + pad, y - size * dy_ratio, text_right - x, ls, size, color, weight=weight, lh=(row_h[0] * scale) / size, wrap=False)
        return y + top - y

    def tag(self, s, x, y, label, size, fill, color, h=None, pad_ratio=0.6154, h_ratio=1.3, text_extra=20):
        """Colour block as wide as the label plus even padding, the label centred in it."""
        pad = size * pad_ratio * 0.5
        h = h or size * h_ratio
        w = text_width(label, size) + pad * 2
        self.rect(s, x, y, w, h, fill)
        self.text(s, x + pad, y + (h - size * 1.13) / 2, w + text_extra, label, size, color, wrap=False)
        return w

    def arrow(self, s, x, y, size, color):
        self.text(s, x, y, size * 1.4, ARROW, size, color, lh=1.0, wrap=False)

    def picture(self, s, path, x, y, w, h, name=None):
        pic = s.shapes.add_picture(str(path), px(x), px(y), px(w), px(h))
        iw, ih = pic.image.size
        box, img = w / h, iw / ih
        if img > box:
            cut = (1 - box / img) / 2
            pic.crop_left = pic.crop_right = cut
        elif img < box:
            cut = (1 - img / box) / 2
            pic.crop_top = pic.crop_bottom = cut
        if name:
            pic.name = name
        return pic

    def image_or_panel(self, s, key, x, y, w, h, panel_color, mask=None):
        path = self.spec_slide.get(key)
        p = self.photo(path, f'{self.where}.{key}') if path else None
        if p:
            self.picture(s, p, x, y, w, h, name=key)
            if mask:
                self.rect(s, x, y, w, h, '#000000', alpha=mask, name=key + '-mask')
            return True
        self.rect(s, x, y, w, h, panel_color, name=key + '-panel')
        self.warnings.append(f'{self.where}：未提供 {key}，图位暂时改为色面。请按 shimo-pptx 补图：先用 scripts/find_images.py 上网找，'
                             f'找不到合适的用 scripts/theme_art.py 按主题生成配图，写进 {key} 后加 --overwrite 重跑。')
        return False

    def logo_pic(self, s, path, x, y, w, max_h, align='left'):
        """Horizontal logo sized by width (as the original master), height capped."""
        if not path:
            return
        pic = s.shapes.add_picture(str(path), px(x), px(y), width=px(w))
        if pic.height > px(max_h):
            ratio = px(max_h) / pic.height
            pic.height = px(max_h)
            pic.width = int(pic.width * ratio)
        if align == 'right':
            pic.left = px(x) - pic.width
        pic.name = 'brand-logo'

    def header(self, s, left=52, right=918, dark=False):
        color = '#FFFFFF' if dark else self.P['ink']
        if self.website:
            self.text(s, left, 24, 300, self.website, 12, color, wrap=False, name='header-website')
        self.logo_pic(s, self.logo_dark if dark else self.logo, right, 24, 76, 18, align='right')

    def notes(self, s):
        parts = []
        if self.spec_slide.get('notes'):
            parts.append(str(self.spec_slide['notes']))
        if self.spec_slide.get('sources'):
            parts.append('来源：' + '；'.join(map(str, self.spec_slide['sources'])))
        credits = [c for c in (image_credit(p) for p in self.slide_images) if c]
        if credits:
            parts.append('图片：' + '；'.join(credits))
            self.image_credits += [{'slide': self.where, 'credit': c} for c in credits]
        if parts:
            s.notes_slide.notes_text_frame.text = '\n'.join(parts)

    # ---------- charts ----------
    def _chart_fonts(self, chart, size_px=13, color=None):
        chart.font.size = Pt(size_px * PTPX)
        chart.font.color.rgb = rgb(color or self.P['ink'])
        txpr = chart._chartSpace.find(qn('c:txPr'))
        if txpr is not None:
            for d in txpr.iter(qn('a:defRPr')):
                for tag in ('a:latin', 'a:ea', 'a:cs'):
                    el = d.find(qn(tag))
                    if el is None:
                        el = etree.SubElement(d, qn(tag))
                    el.set('typeface', CJK_FONT)

    @staticmethod
    def _mix(a, b, t):
        a, b = a.lstrip('#'), b.lstrip('#')
        return '#' + ''.join(f'{round(int(a[i:i + 2], 16) * (1 - t) + int(b[i:i + 2], 16) * t):02X}' for i in (0, 2, 4))

    def series_colors(self, n):
        """Primary first, then grey steps; beyond 3 add primary tints so no two series share a colour."""
        P = self.P
        base = [P['primary'], P['chartSecondary'], P['chartTertiary'],
                self._mix(P['primary'], '#FFFFFF', 0.45), self._mix(P['primary'], '#FFFFFF', 0.7),
                self._mix(P['chartSecondary'], '#000000', 0.35)]
        if n > 3 and self.where not in self._color_warned:
            self._color_warned.add(self.where)
            self.warnings.append(f'{self.where}：图表有 {n} 个系列/类别，超过 3 个时用主色浅色和深灰区分，建议精简或改图。')
        if n > len(base):
            raise ValueError(f'{self.where}：图表系列/类别不能超过 {len(base)} 个')
        return base[:n]

    def native_chart(self, s, spec, x, y, w, h, bg, fonts=None):
        fonts = {**CHART_FONTS, **(fonts or {})}
        kind = spec.get('type', 'column')
        cats = [str(c) for c in spec['categories']]
        series = spec.get('series') or [{'name': spec.get('name', '数值'), 'values': spec['values']}]
        data = CategoryChartData()
        data.categories = cats
        for se in series:
            if len(se['values']) != len(cats):
                raise ValueError(f'{self.where}：系列 {se.get("name")} 的数值个数与类别不一致')
            try:
                vals = [float(v) for v in se['values']]
            except (TypeError, ValueError):
                raise ValueError(f'{self.where}：系列 {se.get("name")} 含非数字值 {se["values"]}')
            data.add_series(se.get('name', ''), vals)
        ctype = {'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'bar': XL_CHART_TYPE.BAR_CLUSTERED,
                 'line': XL_CHART_TYPE.LINE_MARKERS, 'doughnut': XL_CHART_TYPE.DOUGHNUT}[kind]
        gf = s.shapes.add_chart(ctype, px(x), px(y), px(w), px(h), data)
        gf.name = 'chart'
        chart = gf.chart
        chart.has_title = False
        self._chart_fonts(chart, fonts['base'])
        chart.has_legend = len(series) > 1 and kind != 'doughnut'
        if chart.has_legend:
            from pptx.enum.chart import XL_LEGEND_POSITION
            chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.include_in_layout = False
        plot = chart.plots[0]
        fmt = spec.get('number_format', 'General')
        if kind == 'doughnut':
            colors = self.series_colors(len(cats))
            pts = plot.series[0].points
            for i, c in enumerate(colors):
                pts[i].format.fill.solid()
                pts[i].format.fill.fore_color.rgb = rgb(c)
                pts[i].format.line.color.rgb = rgb(bg)
            hole = plot._element.find(qn('c:holeSize'))
            if hole is None:
                hole = etree.SubElement(plot._element, qn('c:holeSize'))
            hole.set('val', '58')
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_percentage = True
            dl.show_value = False
            dl.show_category_name = False
            dl.number_format = '0%'
            dl.number_format_is_linked = False
            dl.font.size = Pt(fonts['doughnut_labels'] * PTPX * 1.3)
            dl.font.color.rgb = rgb(self.P['ink'])
            return chart
        colors = self.series_colors(len(series))
        for se, c in zip(plot.series, colors):
            if kind == 'line':
                se.format.line.color.rgb = rgb(c)
                se.format.line.width = Pt(2.25)
                se.smooth = False
                se.marker.format.fill.solid()
                se.marker.format.fill.fore_color.rgb = rgb(c)
                se.marker.format.line.color.rgb = rgb(c)
            else:
                se.format.fill.solid()
                se.format.fill.fore_color.rgb = rgb(c)
        if kind != 'line':
            plot.gap_width = 90
            plot.overlap = -10 if len(series) > 1 else 0
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_value = True
        dl.number_format = fmt
        dl.number_format_is_linked = fmt == 'General'
        dl.font.size = Pt(fonts['labels'] * PTPX)
        dl.font.color.rgb = rgb(self.P['ink'])
        dl.position = XL_LABEL_POSITION.OUTSIDE_END if kind != 'line' else XL_LABEL_POSITION.ABOVE
        va, ca = chart.value_axis, chart.category_axis
        va.has_major_gridlines = False
        va.minimum_scale = 0 if kind != 'line' or spec.get('zero_base', True) else None
        top = max(float(v) for se in series for v in se['values'])
        if top > 0:
            raw = top / 4
            mag = 10 ** int(f'{raw:e}'.split('e')[1])
            step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
            va.major_unit = step
            va.maximum_scale = step * (int(top / step) + 1)
        va.format.line.color.rgb = rgb(self.P['lineOnPaper'])
        ca.format.line.color.rgb = rgb(self.P['lineOnPaper'])
        va.tick_labels.font.size = Pt(fonts['value_axis'] * PTPX)
        ca.tick_labels.font.size = Pt(fonts['category_axis'] * PTPX)
        if spec.get('hide_value_axis'):
            va.visible = False
        return chart

    # ---------- layouts: objects from layouts.json ----------
    def fns(self):
        def line(text, k):
            ls = lines_of(text)
            return ls[k] if -len(ls) <= k < len(ls) else ''

        def pct(values, k):
            nums = [float(v) for v in values]
            return nums[k] / (sum(nums) or 1) * 100

        return {'len': len, 'str': str, 'int': int, 'float': float, 'min': min, 'max': max, 'abs': abs, 'round': round,
                'lines': lines_of, 'line': line, 'tw': text_width, 'pad2': lambda n: f'{int(n):02d}', 'fmt': format,
                'fit': lambda text, size, max_w, min_ratio=0.62: self.fit_size(text, size, max_w, min_ratio),
                'year': lambda: self.year, 'pct': pct}

    def ev(self, source, scope):
        try:
            return _eval(parse_expr(source), scope, self._fns)
        except (KeyError, IndexError, TypeError, ZeroDivisionError) as e:
            raise ValueError(f'{self.where}：表达式 {source!r} 求值失败（{type(e).__name__}: {e}）') from None

    def val(self, obj, key, scope, default=None):
        """A numeric or other property: literal as written, or an expression string."""
        v = obj.get(key, default)
        return self.ev(v, scope) if isinstance(v, str) else v

    def color(self, value, scope):
        """'#RRGGBB', a palette role ('primary'), roles with fallbacks ('dividerOnPrimary|lineOnPrimary') or an expression."""
        if value is None:
            return None
        if value.startswith('#'):
            return value
        roles = value.split('|')
        if all(r.isidentifier() for r in roles):
            for role in roles:
                if role in self.P:
                    return self.P[role]
            if len(roles) > 1:
                raise ValueError(f'{self.where}：配色 {self.palette_id} 没有颜色 {value}')
        try:
            found = self.ev(value, scope)
        except ValueError:
            found = None
        if isinstance(found, str) and found in self.P:   # an expression may pick a role: 'primary' if closing else 'ink'
            return self.P[found]
        if not (isinstance(found, str) and re.fullmatch(r'#[0-9A-Fa-f]{6}', found)):
            raise ValueError(f'{self.where}：颜色 {value!r} 既不是配色 {self.palette_id} 的颜色，也不是结果为 #RRGGBB 的表达式')
        return found

    def render(self, layout_id, d):
        layout = LAYOUT_DEFS[layout_id]
        for field, spec in layout['fields'].items():
            if spec.get('required') and field not in d:
                raise ValueError(f'{self.where}：缺少字段 {field}')
        scope = Scope(self, vars={'d': d, 'deck': self.spec, 'style': STYLE, 'P': self.P, **layout['vars']}, lets=layout['let'])
        s = self.new_slide(self.color(layout.get('background', 'paper'), scope))
        self.draw(s, layout.get('objects', []), scope)
        return s

    def draw(self, s, objects, scope):
        for obj in objects:
            kind = obj.get('type')
            if kind == 'repeat':
                if 'when' in obj and not self.ev(obj['when'], scope):
                    continue
                self.O_repeat(s, obj, scope)
                continue
            here = Scope(self, lets=obj['let'], parent=scope) if obj.get('let') else scope
            if 'when' in obj and not self.ev(obj['when'], here):
                continue
            method = getattr(self, 'O_' + str(kind), None)
            if method is None:
                raise ValueError(f'{self.where}：未知对象类型 {kind!r}')
            method(s, obj, here)

    def O_group(self, s, o, sc):
        self.draw(s, o.get('objects', []), sc)

    def O_repeat(self, s, o, sc):
        if not isinstance(o.get('max'), int) or o['max'] < 1:
            raise ValueError(f'{self.where}：repeat 要写 max（正整数）')
        items = self.take(self.val(o, 'items', sc), o['max'], o.get('label', 'items'))
        for i, item in enumerate(items):
            here = Scope(self, vars={'item': item, 'i': i, 'n': len(items)}, lets=o.get('let'), parent=sc)
            self.draw(s, o.get('objects', []), here)

    def O_text(self, s, o, sc):
        content = self.ev(o['text'], sc)
        x, y, w = self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc)
        size = self.val(o, 'size', sc)
        self.text(s, x, y, w, content, size, self.color(o['color'], sc), weight=self.val(o, 'weight', sc, 400),
                  lh=self.val(o, 'lh', sc, 1.13), align=o.get('align', 'left'), wrap=o.get('wrap', True),
                  max_h=self.val(o, 'max_h', sc), name=o.get('name'), anchor=o.get('anchor', 'top'),
                  font=o.get('font'), italic=bool(o.get('italic')), spacing=self.val(o, 'spacing', sc), bold=bool(o.get('bold')),
                  para_colors=[self.color(c, sc) for c in o['para_colors']] if o.get('para_colors') else None)

    def gradient_of(self, o, sc):
        g = o.get('gradient')
        if not g:
            return None
        stops = g.get('stops') or []
        if len(stops) < 2:
            raise ValueError(f'{self.where}：gradient 至少要 2 个 stops')
        return (float(g.get('angle', 180)), [(float(p), self.color(c, sc), None if a is None else float(a)) for p, c, a in stops])

    def O_rect(self, s, o, sc):
        self.rect(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc), self.val(o, 'h', sc),
                  self.color(o.get('fill'), sc), alpha=self.val(o, 'alpha', sc), name=o.get('name'), outline=self.color(o.get('outline'), sc),
                  gradient=self.gradient_of(o, sc))

    SHAPES = {'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE, 'rtTriangle': MSO_SHAPE.RIGHT_TRIANGLE,
              'parallelogram': MSO_SHAPE.PARALLELOGRAM, 'ellipse': MSO_SHAPE.OVAL}

    def O_shape(self, s, o, sc):
        """A preset shape (triangle, rtTriangle, parallelogram, ellipse), optionally flipped; filled like rect."""
        kind = o.get('kind')
        if kind not in self.SHAPES:
            raise ValueError(f'{self.where}：shape 的 kind {kind!r} 不支持，可选 {", ".join(self.SHAPES)}')
        sh = self.rect(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc), self.val(o, 'h', sc),
                       self.color(o.get('fill'), sc), alpha=self.val(o, 'alpha', sc), name=o.get('name'),
                       outline=self.color(o.get('outline'), sc), gradient=self.gradient_of(o, sc), shape=self.SHAPES[kind])
        xfrm = sh._element.spPr.find(qn('a:xfrm'))
        if o.get('flip_h'):
            xfrm.set('flipH', '1')
        if o.get('flip_v'):
            xfrm.set('flipV', '1')

    def O_hline(self, s, o, sc):
        self.hline(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc), self.color(o['color'], sc))

    def O_vline(self, s, o, sc):
        self.vline(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'h', sc), self.color(o['color'], sc))

    def O_arrow(self, s, o, sc):
        self.arrow(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'size', sc), self.color(o['color'], sc))

    def O_tag(self, s, o, sc):
        extra = {k: o[k] for k in ('pad_ratio', 'h_ratio', 'text_extra') if k in o}
        self.tag(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.ev(o['text'], sc), self.val(o, 'size', sc),
                 self.color(o['fill'], sc), self.color(o['color'], sc), h=self.val(o, 'h', sc), **extra)

    def O_highlight(self, s, o, sc):
        extra = {k: o[k] for k in ('max_w', 'text_right', 'base', 'pad_ratio', 'dy_ratio', 'weight') if k in o}
        self.highlight_title(s, self.val(o, 'x', sc), self.val(o, 'y', sc), self.ev(o['text'], sc), self.val(o, 'size', sc),
                             self.color(o['fill'], sc), self.color(o['color'], sc), row_h=tuple(o.get('row_h', (70, 73))),
                             pad=self.val(o, 'pad', sc), **extra)

    def O_image(self, s, o, sc):
        self.image_or_panel(s, o['slot'], self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc), self.val(o, 'h', sc),
                            self.color(o.get('panel', 'secondary'), sc), mask=self.val(o, 'mask', sc))

    def O_logo(self, s, o, sc):
        self.logo_pic(s, self.logo_dark if o.get('dark') else self.logo, self.val(o, 'x', sc), self.val(o, 'y', sc),
                      self.val(o, 'w', sc), self.val(o, 'max_h', sc), align=o.get('align', 'left'))

    def O_header(self, s, o, sc):
        self.header(s, self.val(o, 'left', sc), self.val(o, 'right', sc), dark=bool(o.get('dark')))

    def O_chart(self, s, o, sc):
        spec = self.ev(o['chart'], sc)
        if o.get('kind'):
            spec = {**spec, 'type': o['kind']}
        self.native_chart(s, spec, self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc), self.val(o, 'h', sc),
                          self.color(o.get('bg', 'paper'), sc), fonts=o.get('fonts'))

    def O_legend(self, s, o, sc):
        """Doughnut legend: colour swatch, label (with share by default) and a rule per category."""
        ch = self.ev(o['chart'], sc)
        x, y, w = self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc)
        swatch, size, step = o.get('swatch', 12), o.get('size', 13), o.get('step', 30.5)
        colors = self.series_colors(len(ch['categories']))
        try:
            total = sum(float(v) for v in ch['values']) or 1
        except (TypeError, ValueError):
            raise ValueError(f'{self.where}：环图 values 必须全是数字')
        ink, rule = self.color(o.get('color', 'ink'), sc), self.color(o.get('rule', 'lineOnPaper'), sc)
        for cat, val, col in zip(ch['categories'], ch['values'], colors):
            self.rect(s, x, y + o.get('swatch_dy', 2), swatch, swatch, col)
            label = f'{cat} {float(val) / total * 100:.0f}%' if ch.get('legend_percent', True) else str(cat)
            self.text(s, self.val(o, 'label_x', sc), y, self.val(o, 'label_w', sc), label, size, ink, wrap=False)
            self.hline(s, x, y + o.get('rule_dy', 27), w, rule)
            y += step

    def O_table(self, s, o, sc):
        """Native table: header row (and first column unless row_labels is false) on head_fill, borders, widths."""
        t = self.ev(o['table'], sc)
        x, top, width = self.val(o, 'x', sc), self.val(o, 'y', sc), self.val(o, 'w', sc)
        cols, rows = t['columns'], t['rows']
        nr, nc = len(rows) + 1, len(cols)
        avail = self.val(o, 'bottom', sc) - top
        row_h = min(o.get('max_row_h', 46), avail / nr)
        gf = s.shapes.add_table(nr, nc, px(x), px(top), px(width), px(row_h * nr))
        gf.name = 'table'
        tbl = gf.table
        sid = tbl._tbl.tblPr.find(qn('a:tableStyleId'))
        if sid is None:
            sid = etree.SubElement(tbl._tbl.tblPr, qn('a:tableStyleId'))
        sid.text = o.get('style_id', '{2D5ABB26-0587-4C30-8999-92F81FD0307C}')
        for i, row in enumerate(rows, 1):
            if not isinstance(row, list) or len(row) != nc:
                raise ValueError(f'{self.where}：表格第 {i} 行有 {len(row) if isinstance(row, list) else "?"} 个单元格，应为 {nc} 个')
        widths = t.get('column_widths') or ([1.3] + [1] * (nc - 1))
        if len(widths) != nc or not all(isinstance(v, (int, float)) and v > 0 for v in widths):
            raise ValueError(f'{self.where}：column_widths 必须是 {nc} 个正数')
        tot = sum(widths)
        for i, wv in enumerate(widths[:nc]):
            tbl.columns[i].width = px(width * wv / tot)
        for r in range(nr):
            tbl.rows[r].height = px(row_h)
        fs = t.get('font_size', o.get('font', 15) if nr <= o.get('small_after_rows', 7) else o.get('font_small', 13))
        head_fill, head_color = self.color(o.get('head_fill', 'primary'), sc), self.color(o.get('head_color', 'onPrimary'), sc)
        body_fill, body_color = self.color(o.get('body_fill', 'paper'), sc), self.color(o.get('body_color', 'ink'), sc)
        border = self.color(o.get('border', 'chartTertiary'), sc).lstrip('#')
        for r in range(nr):
            for c in range(nc):
                cell = tbl.cell(r, c)
                val = cols[c] if r == 0 else rows[r - 1][c]
                val = '' if val is None else val
                tcpr = cell._tc.get_or_add_tcPr()
                for side in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                    ln = etree.SubElement(tcpr, qn(side), w='9525')
                    sf = etree.SubElement(ln, qn('a:solidFill'))
                    etree.SubElement(sf, qn('a:srgbClr'), val=border)
                head = r == 0 or (c == 0 and t.get('row_labels', True))
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(head_fill if head else body_fill)
                cell.margin_left = cell.margin_right = px(o.get('pad_x', 14))
                cell.margin_top = cell.margin_bottom = px(o.get('pad_y', 4))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
                run = p.add_run()
                run.text = str(val)
                run.font.size = Pt(fs * PTPX)
                run.font.bold = False
                run.font.color.rgb = rgb(head_color if head else body_color)
                self._fonts(run, CJK_FONT)

    # ---------- build ----------
    def build(self, out):
        slides = self.spec.get('slides') or []
        if not slides:
            raise ValueError('slides 不能为空')
        report = []
        for i, d in enumerate(slides, 1):
            if not isinstance(d, dict):
                raise ValueError(f'第 {i} 页必须是对象，例如 {{"layout": "cover"}}')
            layout = d.get('layout')
            if layout not in LAYOUTS:
                raise ValueError(f'第 {i} 页 layout={layout!r} 不存在，可选：{", ".join(LAYOUTS)}')
            self.where = f'第{i}页({layout})'
            self.spec_slide = d
            self.slide_images = []
            s = self.render(layout, d)
            self.notes(s)
            report.append({'slide': i, 'layout': layout})
        rules = DESC.get('deck', {})
        for end, page in (('first', slides[0]), ('last', slides[-1])):
            if rules.get(end) and page.get('layout') != rules[end]:
                self.warnings.append(rules.get('warnings', {}).get(end) or f'{"第 1 页" if end == "first" else "最后一页"}不是 {rules[end]}。')
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out))
        check = Presentation(str(out))
        return {'output': str(out.resolve()), 'slides': len(check.slides), 'style': STYLE_ID, 'palette': self.palette_id,
                'layouts': report, 'logo': 'included' if self.logo else 'omitted',
                'images': list(self.used_images.values()), 'image_credits': self.image_credits, 'warnings': self.warnings}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    b = sub.add_parser('build')
    b.add_argument('--spec', required=True, type=Path)
    b.add_argument('--out', required=True, type=Path)
    b.add_argument('--overwrite', action='store_true')
    b.add_argument('--style', default=DEFAULT_STYLE, help='style pack name (a folder under styles/); default: default')
    lp = sub.add_parser('layouts')
    lp.add_argument('--style', default=DEFAULT_STYLE)
    for p in (b, lp):
        p.add_argument('--styles-dir', type=Path, help='folder holding the packs (style-distill staging); default: styles/')
    a = ap.parse_args()
    if a.styles_dir:
        try:
            set_styles_dir(a.styles_dir)
        except ValueError as e:
            print(f'读取失败：{e}', file=sys.stderr)
            return 1
    if a.cmd == 'layouts':
        try:
            use_style(a.style)
        except (ValueError, KeyError, OSError, json.JSONDecodeError) as e:
            print(f'读取失败：{type(e).__name__}: {e}', file=sys.stderr)
            return 1
        print(json.dumps(LAYOUTS, ensure_ascii=False, indent=2))
        return 0
    if a.out.suffix.lower() != '.pptx':
        ap.error('--out 必须是 .pptx')
    if a.out.exists() and not a.overwrite:
        ap.error(f'{a.out} 已存在；换个文件名或加 --overwrite')
    try:
        use_style(a.style)
        spec = json.loads(a.spec.read_text(encoding='utf-8'))
        result = Deck(spec, a.spec.resolve().parent).build(a.out)
    except (ValueError, KeyError, TypeError, AttributeError, OSError, json.JSONDecodeError) as e:
        print(f'生成失败：{type(e).__name__}: {e}', file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    sys.exit(main())
