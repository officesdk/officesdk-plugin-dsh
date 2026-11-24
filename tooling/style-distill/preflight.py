#!/usr/bin/env python3
"""Preflight a source package before distilling it into a shimo-pptx style pack (style-distill tooling).

    python3 -I tooling/style-distill/preflight.py <package.zip | package dir> --out report.md [--json report.json]
    python3 -I tooling/style-distill/preflight.py <参考图片 | x.pdf | x.pptx | 参考目录> --out report.md [--json report.json] [--ref | --package]

Writes a one-page Markdown report: output format (PPTX / HTML), runtime dependencies, fonts and images with their
licence and provenance (marking what cannot be redistributed), where the design comes from, how many photos a deck
needs, and the smallest font size. Standard library only; Python 3.9+.

The package is untrusted data. Nothing in it is executed or imported: Python files are only parsed with ast, the rest
is read as text or bytes. A zip is listed first (absolute paths, '..', symlinks, more than 5000 entries or 500 MB
unpacked are refused), then unpacked file by file into a new temporary directory that is made read-only and removed
afterwards. A directory is walked in place without following symlinks and without writing to it.

Reference mode (a single image / PDF / PPTX, or a directory holding only such files and fonts): reports pages, pixel
sizes, PDF text layer and fonts (poppler's pdfinfo / pdftotext / pdffonts when installed, with timeouts), a PPTX
inventory (layouts, font names, media; XML read with the standard library after the zip checks, DOCTYPE refused),
font files with their OS/2 fsType, and always asks who owns the design. --ref / --package force a mode.
The PPTX inventory, image header reader and font parsers here are also loaded by ref_intake.py and measure_image.py.
"""
import argparse
import ast
import hashlib
import io
import json
import os
import re
import shutil
import stat
import struct
import subprocess
import sys
import sysconfig
import tempfile
import zipfile
import zlib
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath

MAX_ENTRIES = 5000
MAX_BYTES = 500 * 1024 * 1024
MAX_TEXT = 4 * 1024 * 1024
MAX_FONT_TABLE = 64 * 1024 * 1024
TEXT_EXT = {'.md', '.txt', '.json', '.py', '.mjs', '.js', '.cjs', '.ts', '.html', '.htm', '.css', '.yaml', '.yml', '.svg', '.xml', '.csv'}
FONT_EXT = {'.ttf', '.otf', '.ttc', '.woff', '.woff2'}
IMAGE_EXT = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp', '.tif', '.tiff'}
FREE_LICENSE = re.compile(r'SIL Open Font License|OFL|Apache License|Ubuntu Font Licen[cs]e|MIT License', re.I)
NONFREE_LICENSE = re.compile(r'all rights reserved|may not be (?:re)?distributed|not for redistribution|proprietary|commercial license', re.I)
SYSTEM_FONT_PATH = re.compile(r'(/System/Library/Fonts|/Library/Fonts|~/Library/Fonts|C:\\\\Windows\\\\Fonts|C:/Windows/Fonts)', re.I)
# Families that ship with an OS or under a commercial licence: never redistributable inside a style pack.
PROPRIETARY = ['Times', 'Helvetica', 'Arial', 'PingFang', 'Hiragino', 'Microsoft YaHei', 'SimSun', 'SimHei', 'Songti', 'Heiti', 'STHeiti',
               'Kaiti', 'Georgia', 'Verdana', 'Segoe', 'SF Pro', 'San Francisco', 'Menlo', 'Monaco', 'Apple Color Emoji', 'Calibri',
               'Cambria', 'Avenir', 'Futura', 'Gill Sans', 'Optima', 'Palatino', 'Baskerville', 'Didot', 'Lucida Grande', 'Courier New']
CODEX_MARKERS = ['@oai/artifact-tool', 'skia-canvas', 'container_tools', '/runtime/node_modules', 'codex-primary-runtime']
NODE_BUILTINS = {'fs', 'path', 'url', 'os', 'module', 'crypto', 'child_process', 'util', 'stream', 'events', 'zlib', 'http', 'https',
                 'net', 'process', 'buffer', 'assert', 'readline', 'worker_threads', 'fs/promises', 'timers', 'string_decoder', 'tty', 'v8', 'vm'}
DESIGN_WORDS = re.compile(r'Kimi|Genspark|Gamma|Canva|Beautiful\.ai|Tome|重建|复刻|仿照|照着|参考|reference|官方|Download|based on|inspired|template|模板|来源|出处|设计', re.I)
PPTX_CODE = re.compile(r'python-pptx|from pptx\b|import pptx\b|pptxgenjs|@oai/artifact-tool|\.pptx[\'"`]|Presentation\(')
HTML_WRITE = re.compile(r'\.html[\'"`]|text/html|<!doctype html', re.I)


class Refused(Exception):
    pass


# ---------- reading the package safely ----------
def _is_symlink(info):
    return stat.S_ISLNK(info.external_attr >> 16)


def check_zip(path):
    """Refuse dangerous or oversized archives before anything is written."""
    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        if len(infos) > MAX_ENTRIES:
            raise Refused(f'压缩包有 {len(infos)} 个条目，超过 {MAX_ENTRIES}')
        total = 0
        for info in infos:
            name = info.filename
            parts = PurePosixPath(name.replace('\\', '/')).parts
            if name.startswith(('/', '\\')) or re.match(r'^[A-Za-z]:', name):
                raise Refused(f'压缩包含绝对路径：{name}')
            if '..' in parts:
                raise Refused(f'压缩包含 ..：{name}')
            if _is_symlink(info):
                raise Refused(f'压缩包含符号链接：{name}')
            total += info.file_size
            if total > MAX_BYTES:
                raise Refused(f'解压后超过 {MAX_BYTES // (1024 * 1024)} MB')
        return infos


def unpack(path, dest):
    infos = check_zip(path)
    with zipfile.ZipFile(path) as zf:
        for info in infos:
            if info.is_dir():
                continue
            target = dest.joinpath(*PurePosixPath(info.filename.replace('\\', '/')).parts)
            if dest.resolve() not in target.resolve().parents:
                raise Refused(f'条目会写到解压目录外：{info.filename}')
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, 'wb') as out:
                shutil.copyfileobj(src, out)
    for root, dirs, files in os.walk(dest):
        for name in files:
            os.chmod(os.path.join(root, name), 0o444)
        for name in dirs:
            os.chmod(os.path.join(root, name), 0o555)


def make_writable(dest):
    for root, dirs, files in os.walk(dest):
        os.chmod(root, 0o755)
        for name in files:
            os.chmod(os.path.join(root, name), 0o644)


class Package:
    def __init__(self, root):
        self.root = root
        self.files = []      # relative POSIX paths of regular files
        self.symlinks = []
        for dirpath, dirs, files in os.walk(root, followlinks=False):
            for name in list(dirs):
                if os.path.islink(os.path.join(dirpath, name)):
                    self.symlinks.append(os.path.relpath(os.path.join(dirpath, name), root))
                    dirs.remove(name)
            for name in files:
                full = os.path.join(dirpath, name)
                rel = Path(os.path.relpath(full, root)).as_posix()
                if os.path.islink(full):
                    self.symlinks.append(rel)
                elif os.path.isfile(full):
                    self.files.append(rel)
        self.files.sort()
        self._text = {}

    def size(self, rel):
        return os.path.getsize(os.path.join(self.root, rel))

    def bytes(self, rel):
        with open(os.path.join(self.root, rel), 'rb') as f:
            return f.read()

    def text(self, rel):
        if rel not in self._text:
            data = b''
            if Path(rel).suffix.lower() in TEXT_EXT and self.size(rel) <= MAX_TEXT:
                data = self.bytes(rel)
            self._text[rel] = data.decode('utf-8', 'replace')
        return self._text[rel]

    def texts(self, *exts):
        return [(rel, self.text(rel)) for rel in self.files if Path(rel).suffix.lower() in exts]

    def json(self, rel):
        try:
            return json.loads(self.text(rel))
        except ValueError:
            return None


# ---------- 1. format and 2. runtime ----------
def _stdlib_names():
    names = set(getattr(sys, 'stdlib_module_names', ())) | set(sys.builtin_module_names)
    stdlib = sysconfig.get_paths().get('stdlib')
    if stdlib and os.path.isdir(stdlib):
        for entry in os.listdir(stdlib):
            names.add(entry[:-3] if entry.endswith('.py') else entry)
    return names


def python_imports(source):
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return None
    found = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            found.update(alias.name.split('.')[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            found.add(node.module.split('.')[0])
    return found


NODE_IMPORT = re.compile(r'''(?:import\s[^'"]*?from\s*|import\s*\(\s*|require\s*\(\s*|import\s+)['"]([^'"]+)['"]''')


def node_imports(source):
    return set(NODE_IMPORT.findall(source))


def runtime(pkg):
    stdlib = _stdlib_names()
    py_third, py_local, unparsed = set(), set(), []
    local_modules = {Path(rel).stem for rel in pkg.files if rel.endswith('.py')}
    for rel, src in pkg.texts('.py'):
        mods = python_imports(src)
        if mods is None:
            unparsed.append(rel)
            continue
        for mod in mods:
            if mod in stdlib or mod == '__future__':
                continue
            (py_local if mod in local_modules else py_third).add(mod)
    node_third, node_builtin = set(), set()
    for rel, src in pkg.texts('.mjs', '.js', '.cjs', '.ts'):
        for spec in node_imports(src):
            if spec.startswith(('.', '/')):
                continue
            bare = spec[5:] if spec.startswith('node:') else spec
            if spec.startswith('node:') or bare in NODE_BUILTINS:
                node_builtin.add(bare)
            else:
                parts = spec.split('/')
                node_third.add('/'.join(parts[:2]) if spec.startswith('@') else parts[0])
    for rel in pkg.files:
        if PurePosixPath(rel).name == 'package.json':
            data = pkg.json(rel) or {}
            for key in ('dependencies', 'peerDependencies', 'optionalDependencies'):
                node_third.update((data.get(key) or {}).keys())
    markers, paths, network, processes = {}, {}, set(), set()
    for rel in pkg.files:
        src = pkg.text(rel)
        if not src:
            continue
        for marker in CODEX_MARKERS:
            if marker in src:
                markers.setdefault(marker, []).append(rel)
        code = Path(rel).suffix.lower() in {'.py', '.mjs', '.js', '.cjs', '.ts'}
        for m in re.finditer(r'(?:/Users/[^\s\'"`)]+|~/Library/[^\s\'"`)]+|/System/Library/[^\s\'"`)]+|Library/Fonts/[^\s\'"`)]+|C:\\\\[^\s\'"`]+)', src):
            paths.setdefault(m.group(0)[:90], set()).add(rel)
        if code:
            if re.search(r'\bchild_process\b|\bsubprocess\b|\bspawnSync\b|\bexecSync\b|os\.system\(', src):
                processes.add(rel)
            if re.search(r'\bfetch\(|urllib\.request|\brequests\.|https?\.get\(|XMLHttpRequest', src):
                network.add(rel)
    has_python = any(rel.endswith('.py') for rel in pkg.files)
    has_node = any(Path(rel).suffix in {'.mjs', '.js', '.cjs'} for rel in pkg.files)
    if markers:
        verdict = '需要 OpenAI Codex 运行时（DSH 和客户环境没有）'
    elif py_third or node_third:
        verdict = '需要第三方依赖：' + '、'.join(sorted(py_third | node_third))
    elif has_python and not has_node:
        verdict = '只要 Python 标准库'
    elif has_node and not has_python:
        verdict = '只要 Node 内置模块'
    elif has_python or has_node:
        verdict = '只要 Python 标准库与 Node 内置模块'
    else:
        verdict = '没有脚本'
    return {'verdict': verdict, 'python_third_party': sorted(py_third), 'node_packages': sorted(node_third),
            'node_builtins': sorted(node_builtin), 'codex_markers': {k: sorted(set(v)) for k, v in markers.items()},
            'hardcoded_paths': {k: sorted(v) for k, v in sorted(paths.items())}, 'spawns_processes': sorted(processes),
            'network': sorted(network), 'unparsed_python': unparsed}


def output_format(pkg, deps):
    pptx_code = sorted({rel for rel, src in pkg.texts('.py', '.mjs', '.js', '.cjs', '.ts') if PPTX_CODE.search(src)})
    if deps['codex_markers'].get('@oai/artifact-tool'):
        pptx_code = sorted(set(pptx_code) | set(deps['codex_markers']['@oai/artifact-tool']))
    html_templates = [rel for rel in pkg.files if rel.lower().endswith(('.html', '.htm')) and '/slides' in '/' + rel.lower()]
    html_any = [rel for rel in pkg.files if rel.lower().endswith(('.html', '.htm'))]
    html_code = sorted({rel for rel, src in pkg.texts('.py', '.mjs', '.js', '.cjs') if HTML_WRITE.search(src)})
    pptx_files = [rel for rel in pkg.files if rel.lower().endswith('.pptx')]
    skill_lines = []
    for rel in pkg.files:
        if PurePosixPath(rel).name.upper() == 'SKILL.MD':
            skill_lines += [line.strip() for line in pkg.text(rel).splitlines() if re.search(r'pptx|html|幻灯|演示', line, re.I)][:6]
    is_pptx = bool(pptx_code)
    is_html = bool(html_templates) or (bool(html_code) and len(html_any) >= 3)
    fmt = 'PPTX 与 HTML' if is_pptx and is_html else 'PPTX' if is_pptx else 'HTML' if is_html else '不明'
    return {'format': fmt, 'pptx_code': pptx_code, 'html_templates': len(html_templates), 'html_files': len(html_any),
            'html_writers': html_code, 'pptx_files': pptx_files, 'skill_says': skill_lines}


# ---------- 3. fonts ----------
def _sfnt_tables(data, offset=0):
    num = struct.unpack('>H', data[offset + 4:offset + 6])[0]
    tables = {}
    for i in range(num):
        rec = offset + 12 + i * 16
        tag, _, off, length = struct.unpack('>4sLLL', data[rec:rec + 16])
        if off + length > len(data):
            raise IndexError('table outside the font')
        tables[tag.decode('latin-1')] = data[off:off + length]
    return tables


def _woff_tables(data):
    num = struct.unpack('>H', data[12:14])[0]
    tables = {}
    for i in range(num):
        rec = 44 + i * 20
        tag, off, comp, orig, _ = struct.unpack('>4sLLLL', data[rec:rec + 20])
        if off + comp > len(data):
            raise IndexError('table outside the font')
        raw = data[off:off + comp]
        if orig > MAX_FONT_TABLE:
            raise zlib.error('table too large')
        # Bounded: a crafted table cannot inflate past its declared size (itself capped).
        tables[tag.decode('latin-1')] = zlib.decompressobj().decompress(raw, orig) if comp < orig else raw
    return tables


def _names(table):
    out = {}
    if not table or len(table) < 6:
        return out
    _, count, base = struct.unpack('>HHH', table[:6])
    for i in range(count):
        pid, eid, lid, nid, length, off = struct.unpack('>HHHHHH', table[6 + i * 12:18 + i * 12])
        raw = table[base + off:base + off + length]
        try:
            value = raw.decode('utf-16-be') if pid in (0, 3) else raw.decode('latin-1')
        except UnicodeDecodeError:
            continue
        english = (pid == 3 and lid == 0x409) or pid == 1 or pid == 0
        if nid not in out or english:
            out[nid] = value
    return out


FSTYPE_BITS = [(0x0002, '限制嵌入'), (0x0004, '仅预览打印'), (0x0008, '可编辑嵌入'), (0x0100, '不可子集'), (0x0200, '仅位图')]


def fstype_label(value):
    """OS/2 fsType in words: one embedding level (the most permissive set bit wins, as OpenType says), then the flags."""
    if value is None:
        return '读不出'
    level = '可编辑嵌入' if value & 0x0008 else '仅预览打印' if value & 0x0004 else '限制嵌入' if value & 0x0002 else '可安装嵌入'
    flags = [label for bit, label in FSTYPE_BITS if bit >= 0x0100 and value & bit]
    return '、'.join([level] + flags)


def _fs_type(table):
    if not table or len(table) < 10:
        return None
    return struct.unpack('>H', table[8:10])[0]


def font_names(data):
    """{family, subfamily, copyright, license, license_url, fs_type} from a TTF/OTF/TTC/WOFF; {} for WOFF2/unknown."""
    try:
        if data[:4] == b'wOFF':
            tables = _woff_tables(data)
        elif data[:4] == b'ttcf':
            first = struct.unpack('>L', data[12:16])[0]
            tables = _sfnt_tables(data, first)
        elif data[:4] in (b'\x00\x01\x00\x00', b'OTTO', b'true'):
            tables = _sfnt_tables(data)
        else:
            return {}
        names = _names(tables.get('name'))
        fs_type = _fs_type(tables.get('OS/2'))
    except (struct.error, zlib.error, IndexError):
        return {}
    return {'family': names.get(16) or names.get(1), 'subfamily': names.get(17) or names.get(2), 'copyright': names.get(0),
            'license': names.get(13), 'license_url': names.get(14), 'fs_type': fs_type}


def _norm(text):
    return re.sub(r'[^a-z0-9]', '', (text or '').lower())


def fonts(pkg):
    files = [rel for rel in pkg.files if Path(rel).suffix.lower() in FONT_EXT]
    aliases = {rel: {PurePosixPath(rel).name} for rel in files}
    families = {rel: set() for rel in files}
    locals_ = {rel: set() for rel in files}
    # CSS @font-face: family and local() names for each url().
    for rel, css in pkg.texts('.css', '.html'):
        for block in re.findall(r'@font-face\s*\{([^}]*)\}', css, re.I):
            fam = re.search(r'font-family\s*:\s*[\'"]?([^;\'"]+)', block)
            for url in re.findall(r'url\(\s*[\'"]?([^\'")]+)', block):
                base = PurePosixPath(url).name
                for f in files:
                    if PurePosixPath(f).name == base:
                        if fam:
                            families[f].add(fam.group(1).strip())
                        locals_[f].update(re.findall(r'local\(\s*[\'"]?([^\'")]+)', block))
    # JSON manifests and source records: provenance per file (by any alias).
    provenance = {}
    def walk(node, where):
        if isinstance(node, dict):
            names = [v for k, v in node.items() if isinstance(v, str) and Path(v).suffix.lower() in FONT_EXT]
            for k in ('originalFiles', 'files'):
                if isinstance(node.get(k), list):
                    names += [v for v in node[k] if isinstance(v, str)]
            if names:
                record = {k: v for k, v in node.items() if isinstance(v, (str, int))}
                for name in names:
                    provenance.setdefault(PurePosixPath(name).name, []).append({'record': record, 'from': where, 'all': names})
            for v in node.values():
                walk(v, where)
        elif isinstance(node, list):
            for v in node:
                walk(v, where)
    for rel in pkg.files:
        if rel.endswith('.json') and ('font' in rel.lower() or 'source' in rel.lower()):
            walk(pkg.json(rel), rel)
    for f in files:
        for entry in provenance.get(PurePosixPath(f).name, []):
            aliases[f].update(PurePosixPath(n).name for n in entry['all'])
    licence_files = [rel for rel in pkg.files if re.search(r'(OFL|LICEN[SC]E|COPYING)', PurePosixPath(rel).name, re.I)]
    licence_heads = {rel: '\n'.join(pkg.text(rel).splitlines()[:12]) for rel in licence_files}

    def covered_by_licence(family):
        """A free licence file for this family: its file name names the family (longest match), else its header does."""
        n = _norm(family)
        free = [rel for rel in licence_heads if FREE_LICENSE.search(pkg.text(rel))]
        by_name = [(len(stem), rel) for rel in free
                   for stem in [_norm(re.sub(r'(?i)(ofl|licen[sc]e|copying|\.txt|\.md)', '', PurePosixPath(rel).name))] if stem and n.startswith(stem)]
        if by_name:
            return max(by_name)[1]
        for rel in free:
            if family and re.search(re.escape(family), licence_heads[rel], re.I):
                return rel
        first = (family or '').split()[0] if family else ''
        hits = [rel for rel in free if first and re.search(r'(?<![A-Za-z])' + re.escape(first) + r'(?![A-Za-z])', licence_heads[rel], re.I)]
        return hits[0] if len(hits) == 1 else None

    by_family = {}
    for f in files:
        info = font_names(pkg.bytes(f))
        names = sorted(families[f]) or [info.get('family') or PurePosixPath(sorted(aliases[f])[0]).stem]
        sources = [e['record'] for a in aliases[f] for e in provenance.get(a, [])]
        source_text = json.dumps(sources, ensure_ascii=False)
        for fam in names:
            entry = by_family.setdefault(fam, {'family': fam, 'files': [], 'aliases': set(), 'locals': set(), 'name_table': [], 'sources': [], 'bytes': 0})
            entry['files'].append(f)
            entry['bytes'] += pkg.size(f)
            entry['aliases'].update(a for a in aliases[f] if a != PurePosixPath(f).name)
            entry['locals'].update(locals_[f])
            if info:
                entry['name_table'].append(info)
            if sources:
                entry['sources'].append(source_text[:300])
    result = []
    for fam, e in sorted(by_family.items()):
        hay = ' '.join([fam, *e['aliases'], *e['locals'], *(i.get('family') or '' for i in e['name_table'])])
        system = [s for s in e['sources'] if SYSTEM_FONT_PATH.search(s)]
        proprietary = [p for p in PROPRIETARY if re.search(r'(?<![A-Za-z])' + re.escape(p) + r'(?![A-Za-z])', hay, re.I)]
        name_lic = ' '.join((i.get('license') or '') + ' ' + (i.get('license_url') or '') for i in e['name_table'])
        lic_file = covered_by_licence(fam)
        if system or proprietary or NONFREE_LICENSE.search(name_lic):
            verdict = '不能分发'
            why = ('来源是系统字体 ' + SYSTEM_FONT_PATH.search(system[0]).group(1) if system else '') or \
                  (f'商业/系统字体家族 {proprietary[0]}' if proprietary else '字体自带的授权说明不允许再分发')
        elif FREE_LICENSE.search(name_lic) or lic_file:
            verdict = '可分发'
            why = f'许可文件 {lic_file}' if lic_file else '字体 name 表写明开源授权'
        else:
            verdict = '待查'
            why = '没有找到许可文件或授权说明'
        friendly = sorted(a for a in e['aliases'] | e['locals'] if not re.search(r'[A-Za-z0-9_-]{20,}', a))
        fs_types = sorted({i.get('fs_type') for i in e['name_table'] if i.get('fs_type') is not None})
        result.append({'family': fam, 'verdict': verdict, 'why': why, 'files': len(e['files']), 'bytes': e['bytes'], 'aliases': friendly[:6],
                       'fs_type': fs_types, 'fs_type_label': '；'.join(fstype_label(v) for v in fs_types) or '读不出'})
    # Fonts the code or docs point at but the package does not contain.
    # Provenance records (the font JSON read above) name the files a bundled font was made from; those are not runtime
    # references. A reference to the same font under another extension (NotoSansSC-VF.ttf → .woff2) is bundled.
    bundled = {PurePosixPath(f).name.lower() for f in files} | {a.lower() for f in files for a in aliases[f]}
    bundled_stems = {Path(name).stem.lower() for name in bundled}
    provenance_files = {e['from'] for entries in provenance.values() for e in entries}
    referenced = {}
    for rel in pkg.files:
        if Path(rel).suffix.lower() in FONT_EXT or rel in provenance_files:
            continue
        for m in re.finditer(r'([\w./~@-]*?([\w@-]+\.(?:ttf|otf|ttc)))\b', pkg.text(rel), re.I):
            if m.group(2).lower() not in bundled and Path(m.group(2)).stem.lower() not in bundled_stems:
                referenced.setdefault(m.group(2), set()).add(rel)
    missing = [{'file': name, 'verdict': '待查（引用了但没打包，需另行取得授权）', 'where': sorted(where)[:4]} for name, where in sorted(referenced.items())]
    return {'families': result, 'referenced_not_bundled': missing, 'licence_files': licence_files}


# ---------- 4. images ----------
def images(pkg):
    imgs = [rel for rel in pkg.files if Path(rel).suffix.lower() in IMAGE_EXT and '/fonts/' not in '/' + rel.lower()]
    recorded = set()
    record_files = [rel for rel in pkg.files if rel.endswith('.json') and re.search(r'source|credit|licen', PurePosixPath(rel).name, re.I)
                    and 'font' not in rel.lower()]
    for rel in record_files:
        text = pkg.text(rel)
        for img in imgs:
            if PurePosixPath(img).name in text:
                recorded.add(img)
    reference = [rel for rel in imgs if re.search(r'(^|/)(reference|references|ref|preview|screenshots?)(/|$)', rel, re.I)]
    rows = {}
    for rel in imgs:
        folder = str(PurePosixPath(rel).parent)
        row = rows.setdefault(folder, {'folder': folder, 'count': 0, 'recorded': 0, 'reference': 0})
        row['count'] += 1
        row['recorded'] += rel in recorded
        row['reference'] += rel in reference
    return {'total': len(imgs), 'recorded': len(recorded), 'reference': len(reference), 'record_files': record_files,
            'folders': sorted(rows.values(), key=lambda r: r['folder'])}


# ---------- 5. design source ----------
def design_source(pkg):
    lines = []
    docs = [rel for rel in pkg.files if re.search(r'(^|/)(SKILL\.md|README[^/]*|references/[^/]+\.md|agents/[^/]+\.ya?ml|[^/]*manifest[^/]*\.json)$', rel, re.I)]
    for rel in docs:
        for n, line in enumerate(pkg.text(rel).splitlines(), 1):
            text = line.strip()
            if text and DESIGN_WORDS.search(text) and re.search(r'Kimi|Genspark|Gamma|Canva|重建|复刻|仿照|照着|Download|based on|inspired|官方|参考', text, re.I):
                lines.append({'where': f'{rel}:{n}', 'text': text[:180]})
    names = sorted({m.group(0) for item in lines for m in re.finditer(r'Kimi(?: Slides)?|Genspark|Gamma|Canva|Beautiful\.ai|Tome', item['text'])})
    return {'mentions': lines[:12], 'more': max(0, len(lines) - 12), 'named_sources': names}


# ---------- 6. photos and 7. smallest font ----------
def masters(pkg):
    """(rel, canvas width, pages) for JSON files shaped like an object-level master: pages[{objects[]}]."""
    out = []
    for rel in pkg.files:
        if not rel.endswith('.json'):
            continue
        data = pkg.json(rel)
        if isinstance(data, dict) and isinstance(data.get('pages'), list) and any(isinstance(p, dict) and isinstance(p.get('objects'), list) for p in data['pages']):
            out.append((rel, data.get('width') or (data.get('canvas') or {}).get('width'), data['pages']))
    return out


def html_layouts(pkg):
    out = []
    for rel in pkg.files:
        if not rel.endswith('.json'):
            continue
        data = pkg.json(rel)
        if isinstance(data, dict) and isinstance(data.get('layouts'), dict) and any(isinstance(v, dict) and 'visuals' in v for v in data['layouts'].values()):
            out.append((rel, data.get('width'), data['layouts']))
    return out


def photos(pkg):
    per_page, method = [], None
    for rel, _, pages in masters(pkg):
        method = f'母版 {rel}：type=image 的对象'
        for i, page in enumerate(pages, 1):
            objs = [o for o in page.get('objects', []) if isinstance(o, dict) and o.get('type') == 'image' and o.get('slotKind', 'photo') == 'photo']
            per_page.append({'page': page.get('name') or page.get('number') or i, 'photos': len(objs)})
    if not per_page:
        for rel, _, layouts in html_layouts(pkg):
            method = f'版式目录 {rel}：visuals 里 kind=image 的槽'
            for lid, layout in layouts.items():
                vis = layout.get('visuals') or {}
                per_page.append({'page': f'{lid} {layout.get("name", "")}'.strip(), 'photos': sum(1 for v in vis.values() if isinstance(v, dict) and v.get('kind') == 'image')})
    total = sum(p['photos'] for p in per_page)
    top = max(per_page, key=lambda p: p['photos']) if per_page else None
    return {'total': total, 'pages': len(per_page), 'max_page': top, 'per_page': per_page, 'method': method or '没有找到母版或版式目录，未统计'}


def smallest_font(pkg):
    samples = []
    for rel, width, pages in masters(pkg):
        width = width or 960
        for page in pages:
            for o in page.get('objects', []):
                if isinstance(o, dict) and isinstance(o.get('fontSize'), (int, float)):
                    samples.append((o['fontSize'] * 720 / width, f'{rel} {page.get("name", "")} {o.get("id", "")}'.strip(), o['fontSize'], width))
    html_width = next((w for _, w, _ in html_layouts(pkg) if w), None)
    for rel, src in pkg.texts('.html', '.htm', '.css'):
        if 'viewer' in PurePosixPath(rel).name.lower() or '/fonts/' in '/' + rel:
            continue
        width = html_width or (1920 if re.search(r'1920px|width:\s*1920', src) else None)
        if not width:
            continue
        for m in re.finditer(r'font-size\s*:\s*([\d.]+)px', src):
            px = float(m.group(1))
            samples.append((px * 720 / width, rel, px, width))
    if not samples:
        return {'min_pt': None, 'note': '没有找到可折算的字号'}
    samples.sort(key=lambda s: s[0])
    low = samples[0]
    below = [s for s in samples if s[0] < 9]
    by_folder = {}
    for pt, where, px, width in samples:
        folder = str(PurePosixPath(where.split(' ')[0]).parent)
        by_folder.setdefault(folder, (pt, px))
    return {'min_pt': round(low[0], 2), 'min_px': low[2], 'canvas_width': low[3], 'where': low[1], 'below_9pt': len(below),
            'samples': len(samples), 'by_folder': {k: {'min_pt': round(v[0], 2), 'min_px': v[1]} for k, v in sorted(by_folder.items())}}


# ---------- report ----------
def analyse(pkg, source):
    deps = runtime(pkg)
    report = {'source': source, 'files': len(pkg.files), 'bytes': sum(pkg.size(f) for f in pkg.files), 'symlinks_skipped': pkg.symlinks,
              'format': output_format(pkg, deps), 'runtime': deps, 'fonts': fonts(pkg), 'images': images(pkg),
              'design': design_source(pkg), 'photos': photos(pkg), 'min_font': smallest_font(pkg)}
    report['verdict'] = verdict(report)
    return report


def verdict(r):
    blockers, decide = [], []
    if r['format']['format'] in ('HTML', '不明'):
        blockers.append(f'产物是 {r["format"]["format"]}，不是 PPTX：要先从页面反推对象几何，工作量大、失真多')
    if r['runtime']['codex_markers']:
        blockers.append('依赖 Codex 运行时：生成脚本不能直接用，只能把母版数据改写成 layouts.json，由 brand_deck.py 渲染')
    bad = [f['family'] for f in r['fonts']['families'] if f['verdict'] == '不能分发']
    if bad:
        blockers.append('字体不能分发：' + '、'.join(bad) + '（样式包只写字体名、不打包，并换成可分发字体）')
    unknown = [f['family'] for f in r['fonts']['families'] if f['verdict'] == '待查'] + [m['file'] for m in r['fonts']['referenced_not_bundled']]
    if unknown:
        decide.append('字体授权待查：' + '、'.join(unknown))
    if r['design']['named_sources'] or r['design']['mentions']:
        decide.append('设计来源：' + ('、'.join(r['design']['named_sources']) or '见摘录') + '，能否给客户用由用户或法务决定')
    if r['images']['reference']:
        decide.append(f'有 {r["images"]["reference"]} 张参考截图，只能对照用，不能进样式包')
    if r['photos']['total'] >= 10:
        blockers.append(f'照片需求大：一份 {r["photos"]["total"]} 个照片槽，找图能力不够时成稿质量会掉')
    mf = r['min_font']
    if mf.get('min_pt') is not None and mf['min_pt'] < 9:
        blockers.append(f'最小字号约 {mf["min_pt"]}pt（{mf["below_9pt"]} 处小于 9pt），要整体重定字号')
    can = not any('HTML' in b or '不明' in b for b in blockers)
    return {'can_extract': can, 'blockers': blockers, 'decide': decide}


def _mb(n):
    return f'{n / 1024 / 1024:.1f} MB'


def markdown(r):
    f, d, fo, im, ds, ph, mf, v = r['format'], r['runtime'], r['fonts'], r['images'], r['design'], r['photos'], r['min_font'], r['verdict']
    out = [f'# 预检报告：{r["source"]}', '',
           f'{r["files"]} 个文件，{_mb(r["bytes"])}' + (f'；跳过符号链接 {len(r["symlinks_skipped"])} 个' if r['symlinks_skipped'] else '') + '。脚本只读不执行。', '',
           '| 项 | 结论 | 依据 |', '|---|---|---|']
    out.append(f'| 产物格式 | **{f["format"]}** | 生成 PPTX 的代码 {len(f["pptx_code"])} 个文件；HTML 页模板 {f["html_templates"]} 个（HTML 文件共 {f["html_files"]}） |')
    deps_note = '；'.join(filter(None, [
        ('Codex 标记 ' + '、'.join(d['codex_markers'])) if d['codex_markers'] else '',
        ('Python 第三方 ' + '、'.join(d['python_third_party'])) if d['python_third_party'] else '',
        ('Node 包 ' + '、'.join(d['node_packages'])) if d['node_packages'] else '',
        ('起子进程 ' + '、'.join(d['spawns_processes'])) if d['spawns_processes'] else '',
        ('联网 ' + '、'.join(d['network'])) if d['network'] else '',
    ])) or '无第三方依赖'
    out.append(f'| 运行时依赖 | **{d["verdict"]}** | {deps_note} |')
    bad = [x for x in fo['families'] if x['verdict'] != '可分发']
    out.append(f'| 字体 | {len(fo["families"])} 个家族：可分发 {sum(x["verdict"] == "可分发" for x in fo["families"])}，'
               f'**不能分发 {sum(x["verdict"] == "不能分发" for x in fo["families"])}**，待查 {sum(x["verdict"] == "待查" for x in fo["families"])}；'
               f'引用未打包 {len(fo["referenced_not_bundled"])} | 见下表 |')
    out.append(f'| 图片 | {im["total"]} 张，有来源记录 {im["recorded"]}，参考截图 {im["reference"]} | 来源记录：{"、".join(im["record_files"]) or "无"} |')
    out.append(f'| 设计来源 | {"、".join(ds["named_sources"]) or ("见摘录" if ds["mentions"] else "未写明")} | 需用户/法务确认 |')
    mp = ph['max_page']
    out.append(f'| 照片需求 | **{ph["total"]}** 个照片槽 / {ph["pages"]} 页' + (f'，最多一页 {mp["photos"]}（{mp["page"]}）' if mp else '') + f' | {ph["method"]} |')
    folders = '；'.join(f'{k} {v["min_pt"]}pt' for k, v in (mf.get('by_folder') or {}).items())
    out.append(f'| 最小字号 | **{mf.get("min_pt")}pt**' + (f'（{mf["min_px"]}px @ {mf["canvas_width"]}px 宽），9pt 以下 {mf["below_9pt"]} 处' if mf.get('min_pt') is not None else '') + f' | 最小在 {mf.get("where", mf.get("note", ""))}' + (f'；各目录最小：{folders}' if folders else '') + ' |')
    out += ['', '## 字体', '', '| 家族 | 判定 | 原因 | 文件 | fsType | 别名/本机名 |', '|---|---|---|---|---|---|']
    for x in fo['families']:
        fs = '、'.join(f'0x{v:04X}' for v in x['fs_type'])
        out.append(f'| {x["family"]} | {"**" + x["verdict"] + "**" if x["verdict"] != "可分发" else x["verdict"]} | {x["why"]} | {x["files"]} 个，{_mb(x["bytes"])} | {(fs + " " + x["fs_type_label"]) if fs else "读不出"} | {"、".join(x["aliases"]) or "-"} |')
    for m in fo['referenced_not_bundled']:
        out.append(f'| {m["file"]} | **{m["verdict"]}** | 出现在 {"、".join(m["where"])} | 未打包 | - | - |')
    if d['hardcoded_paths']:
        out += ['', '写死的本机路径：' + '；'.join(f'`{k}`（{"、".join(v[:2])}）' for k, v in list(d['hardcoded_paths'].items())[:6])]
    out += ['', '## 设计来源摘录', '']
    out += [f'- `{m["where"]}` {m["text"]}' for m in ds['mentions']] or ['- 没有找到出处说明（同样要问清楚）']
    if ds['more']:
        out.append(f'- 另有 {ds["more"]} 处')
    if f['skill_says']:
        out += ['', '## SKILL.md 对产物的自述', ''] + [f'- {line[:180]}' for line in f['skill_says']]
    out += ['', '## 结论', '', f'- 能否进入提取阶段：**{"能" if v["can_extract"] else "不建议"}**']
    out += [f'- 阻碍：{b}' for b in v['blockers']] or ['- 阻碍：无']
    out += [f'- 要人决定：{x}' for x in v['decide']] or ['- 要人决定：无']
    return '\n'.join(out) + '\n'


# ---------- reference mode: images, PDF, PPTX handed over as the design to copy ----------
REF_IMAGE_EXT = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.tif', '.tiff'}
REF_DOC_EXT = {'.pdf', '.pptx'}
MACRO_PPT_EXT = {'.pptm', '.ppsm', '.potm', '.ppam', '.ppt', '.pps', '.pot'}
REF_SIDE_FILES = re.compile(r'^(README[^/]*|[^/]*\.(txt|md))$', re.I)
POPPLER_TIMEOUT = 120
PDF_FONT_TYPE = re.compile(r'\s(Type 1C?(?: \(OT\))?|Type 3|TrueType(?: \(OT\))?|CID Type 0C?(?: \(OT\))?|CID TrueType(?: \(OT\))?|unknown)\s*$')
MAX_PDF_PAGES = 300
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
EMU_PER_PX = 9525


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block)
    return h.hexdigest()


def image_size(head):
    """(width, height) from the first bytes of a PNG / JPEG / GIF / WebP / BMP / TIFF; None when unknown."""
    try:
        if head[:8] == b'\x89PNG\r\n\x1a\n':
            return struct.unpack('>LL', head[16:24])
        if head[:6] in (b'GIF87a', b'GIF89a'):
            return struct.unpack('<HH', head[6:10])
        if head[:2] == b'BM':
            w, h = struct.unpack('<ll', head[18:26])
            return (w, abs(h))
        if head[:4] == b'RIFF' and head[8:12] == b'WEBP':
            chunk = head[12:16]
            if chunk == b'VP8 ':
                w, h = struct.unpack('<HH', head[26:30])
                return (w & 0x3FFF, h & 0x3FFF)
            if chunk == b'VP8L':
                b = head[21:25]
                return (1 + (((b[1] & 0x3F) << 8) | b[0]), 1 + (((b[3] & 0xF) << 10) | (b[2] << 2) | ((b[1] & 0xC0) >> 6)))
            if chunk == b'VP8X':
                return (1 + int.from_bytes(head[24:27], 'little'), 1 + int.from_bytes(head[27:30], 'little'))
        if head[:2] == b'\xff\xd8':
            i = 2
            while i + 9 < len(head):
                if head[i] != 0xFF:
                    i += 1
                    continue
                marker = head[i + 1]
                if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                length = struct.unpack('>H', head[i + 2:i + 4])[0]
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    h, w = struct.unpack('>HH', head[i + 5:i + 9])
                    return (w, h)
                i += 2 + length
        if head[:4] in (b'II*\x00', b'MM\x00*'):
            end = '<' if head[:2] == b'II' else '>'
            ifd = struct.unpack(end + 'L', head[4:8])[0]
            count = struct.unpack(end + 'H', head[ifd:ifd + 2])[0]
            dims = {}
            for k in range(count):
                tag, typ, _, val = struct.unpack(end + 'HHLL', head[ifd + 2 + k * 12:ifd + 14 + k * 12])
                if typ == 3:
                    val = struct.unpack(end + 'H', head[ifd + 10 + k * 12:ifd + 12 + k * 12])[0]
                if tag in (256, 257):
                    dims[tag] = val
            if 256 in dims and 257 in dims:
                return (dims[256], dims[257])
    except (struct.error, IndexError):
        return None
    return None


def aspect_name(w, h):
    if not w or not h:
        return '-'
    r = w / h
    for name, value in (('16:9', 16 / 9), ('4:3', 4 / 3), ('16:10', 16 / 10), ('A4 横', 297 / 210), ('A4 竖', 210 / 297), ('9:16', 9 / 16)):
        if abs(r - value) / value <= 0.02:
            return name
    return f'{r:.3f}'


def _poppler(name):
    return shutil.which(name) or next((p for p in (f'/opt/homebrew/bin/{name}', f'/usr/local/bin/{name}') if os.path.exists(p)), None)


def _run_tool(args):
    try:
        run = subprocess.run(args, capture_output=True, timeout=POPPLER_TIMEOUT)
    except subprocess.TimeoutExpired:
        raise Refused(f'{Path(args[0]).name} 超过 {POPPLER_TIMEOUT} 秒')
    if run.returncode != 0:
        raise Refused(f'{Path(args[0]).name} 失败：{run.stderr.decode("utf-8", "replace").strip()[:200]}')
    return run.stdout.decode('utf-8', 'replace')


def pdf_fonts(path, last=None):
    tool = _poppler('pdffonts')
    if not tool:
        return None
    args = [tool] + (['-l', str(last)] if last else []) + [str(path)]
    lines = _run_tool(args).splitlines()
    sep = next((i for i, line in enumerate(lines) if line.startswith('---')), None)
    if sep is None:
        return []
    fonts = []
    for line in lines[sep + 1:]:
        # pdffonts does not truncate long names, it pushes the columns right: read from the right instead.
        # name type encoding emb sub uni object-number generation; only name and type may contain spaces.
        parts = line.rsplit(None, 6)
        if len(parts) != 7 or parts[2] not in ('yes', 'no'):
            continue
        head, emb = parts[0], parts[2]
        m = PDF_FONT_TYPE.search(head)
        name, kind = (head[:m.start()].strip(), m.group(0).strip()) if m else (head.strip(), '?')
        if not name:
            continue
        name = name.split('+', 1)[1] if re.match(r'^[A-Z]{6}\+', name) else name
        fonts.append({'name': name, 'type': kind, 'embedded': emb == 'yes'})
    unique = {}
    for f in fonts:
        unique.setdefault((f['name'], f['type']), f)
    return list(unique.values())


def pdf_inventory(path):
    """Pages, page sizes (pt), text characters per page and fonts, via poppler; approximate without it."""
    info = _poppler('pdfinfo')
    if not info:
        data = Path(path).read_bytes()
        count = len(re.findall(rb'/Type\s*/Page(?![a-zA-Z])', data))
        return {'pages': count or None, 'approximate': True, 'page_size_pt': [], 'text_layer': None, 'text_chars': None, 'chars_per_page': [],
                'fonts': None, 'note': '没有 poppler：页数按 /Type /Page 计数，' + ('可能不准' if count else '页数未知（压缩对象流），需 brew install poppler')}
    head = _run_tool([info, str(path)])
    m = re.search(r'^Pages:\s+(\d+)', head, re.M)
    pages = int(m.group(1)) if m else 0
    if pages > MAX_PDF_PAGES:
        raise Refused(f'PDF 有 {pages} 页，超过上限 {MAX_PDF_PAGES}')
    sizes = []
    if pages:
        detail = _run_tool([info, '-f', '1', '-l', str(pages), str(path)])
        sizes = [[float(w), float(h)] for w, h in re.findall(r'^Page\s+\d+\s+size:\s+([\d.]+)\s+x\s+([\d.]+)', detail, re.M)]
    chars, text_layer = [], None
    text_tool = _poppler('pdftotext')
    if text_tool and pages:
        text = _run_tool([text_tool, '-layout', '-l', str(pages), str(path), '-'])
        chars = [len(re.sub(r'\s', '', page)) for page in text.split('\f')[:pages]]
        chars += [0] * (pages - len(chars))
        text_layer = any(chars)
    return {'pages': pages, 'approximate': False, 'page_size_pt': sizes, 'text_layer': text_layer, 'text_chars': sum(chars) if chars else None,
            'chars_per_page': chars, 'fonts': pdf_fonts(path, pages)}


def _xml(zf, name):
    data = zf.read(name)
    if re.search(rb'<!DOCTYPE', data[:4096], re.I) or b'<!ENTITY' in data:
        raise Refused(f'{name} 含 DOCTYPE/ENTITY，拒绝解析')
    return ET.fromstring(data)


def _rels(zf, part):
    folder, name = part.rsplit('/', 1)
    rels = f'{folder}/_rels/{name}.rels'
    if rels not in zf.namelist():
        return {}
    out = {}
    for rel in _xml(zf, rels).findall('rel:Relationship', NS):
        target = rel.get('Target', '')
        if rel.get('TargetMode') == 'External':
            continue
        path = PurePosixPath(folder, target).as_posix() if not target.startswith('/') else target.lstrip('/')
        parts = []
        for seg in path.split('/'):
            if seg == '..':
                parts and parts.pop()
            elif seg not in ('', '.'):
                parts.append(seg)
        out[rel.get('Id')] = '/'.join(parts)
    return out


def pptx_inventory(path):
    """Slides, slide size, masters, layouts in use, theme and run fonts, media and embedded fonts of a .pptx (standard library)."""
    path = Path(path)
    if path.suffix.lower() in MACRO_PPT_EXT:
        raise Refused(f'只收 .pptx，{path.suffix} 不读（可能带宏或是旧格式）')
    check_zip(path)
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        if 'ppt/presentation.xml' not in names:
            raise Refused('不是 PPTX：没有 ppt/presentation.xml')
        pres = _xml(zf, 'ppt/presentation.xml')
        rels = _rels(zf, 'ppt/presentation.xml')
        size = pres.find('p:sldSz', NS)
        try:
            cx, cy = (int(size.get('cx')), int(size.get('cy'))) if size is not None else (None, None)
        except (TypeError, ValueError):
            raise Refused('presentation.xml 的页面尺寸 sldSz 不完整')
        if cx is not None and (cx <= 0 or cy <= 0):
            raise Refused('presentation.xml 的页面尺寸不是正数')
        slides = [rels.get(s.get(f'{{{NS["r"]}}}id')) for s in pres.findall('p:sldIdLst/p:sldId', NS)]
        slides = [s for s in slides if s and s in names]
        masters = [rels.get(s.get(f'{{{NS["r"]}}}id')) for s in pres.findall('p:sldMasterIdLst/p:sldMasterId', NS)]
        embedded_named = [f.get('typeface') for f in pres.findall('p:embeddedFontLst/p:embeddedFont/p:font', NS)]
        layouts = {}
        for slide in slides:
            for target in _rels(zf, slide).values():
                if '/slideLayouts/' in target and target in names:
                    c = _xml(zf, target).find('p:cSld', NS)
                    name = (c.get('name') if c is not None else None) or PurePosixPath(target).stem
                    layouts[name] = layouts.get(name, 0) + 1
        theme = {}
        for name in sorted(n for n in names if re.match(r'ppt/theme/theme\d+\.xml$', n)):
            scheme = _xml(zf, name).find('.//a:fontScheme', NS)
            if scheme is None:
                continue
            for kind in ('majorFont', 'minorFont'):
                node = scheme.find(f'a:{kind}', NS)
                if node is None:
                    continue
                for script in ('latin', 'ea', 'cs'):
                    face = node.find(f'a:{script}', NS)
                    if face is not None and face.get('typeface'):
                        theme.setdefault(f'{kind}.{script}', face.get('typeface'))
        used = {}
        parts = [n for n in names if re.match(r'ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$', n)]
        for part in parts:
            root = _xml(zf, part)
            for tag in ('latin', 'ea', 'cs', 'sym'):
                for node in root.iter(f'{{{NS["a"]}}}{tag}'):
                    face = node.get('typeface')
                    if face and not face.startswith('+'):
                        used[face] = used.get(face, 0) + 1
        media = {}
        for n in names:
            if n.startswith('ppt/media/') and not n.endswith('/'):
                ext = PurePosixPath(n).suffix.lower().lstrip('.') or '?'
                media[ext] = media.get(ext, 0) + 1
        embedded = []
        for n in sorted(x for x in names if x.startswith('ppt/fonts/') and not x.endswith('/')):
            info = font_names(zf.read(n))
            embedded.append({'part': n, 'family': info.get('family'), 'fs_type': info.get('fs_type'), 'fs_type_label': fstype_label(info.get('fs_type')) if info else '读不出'})
    return {'slides': len(slides), 'slide_size_emu': [cx, cy], 'slide_size_px': [round(cx / EMU_PER_PX), round(cy / EMU_PER_PX)] if cx else None,
            'aspect': round(cx / cy, 3) if cx and cy else None, 'masters': len(masters),
            'layouts': [{'name': k, 'used_by': v} for k, v in sorted(layouts.items(), key=lambda kv: -kv[1])],
            'fonts': {'theme': theme, 'used': [{'name': k, 'runs': v} for k, v in sorted(used.items(), key=lambda kv: -kv[1])], 'embedded_named': embedded_named},
            'media': dict(sorted(media.items())), 'embedded_fonts': embedded}


def classify(source, force=None):
    """('reference' | 'package', why, files) for a path; files are the reference files when in reference mode."""
    source = Path(source)
    if force == 'package':
        return 'package', '按 --package 指定', []
    if source.is_file():
        ext = source.suffix.lower()
        if ext in REF_IMAGE_EXT | REF_DOC_EXT or ext in MACRO_PPT_EXT:
            kind = '图片' if ext in REF_IMAGE_EXT else ext.lstrip('.').upper()
            return 'reference', f'单个 {kind}', [source]
        if force == 'ref':
            return 'reference', '按 --ref 指定', [source]
        return 'package', '单个 zip' if zipfile.is_zipfile(source) else '单个文件', []
    if not source.is_dir():
        return 'package', '不是目录', []
    refs, others = [], []
    for dirpath, dirs, files in os.walk(source, followlinks=False):
        dirs[:] = sorted(d for d in dirs if not d.startswith('.') and not os.path.islink(os.path.join(dirpath, d)))
        for name in sorted(files):
            full = Path(dirpath, name)
            if name.startswith('.') or full.is_symlink():
                continue
            ext = full.suffix.lower()
            if ext in REF_IMAGE_EXT | REF_DOC_EXT | FONT_EXT or ext in MACRO_PPT_EXT:
                refs.append(full)
            elif name.upper() == 'SKILL.MD' or not REF_SIDE_FILES.match(name):
                others.append(full)
    if force == 'ref':
        return 'reference', '按 --ref 指定', refs
    if others:
        return 'package', f'目录含 {len(others)} 个非参考文件（如 {others[0].relative_to(source)}）', []
    if not any(p.suffix.lower() not in FONT_EXT for p in refs):
        return 'package', '目录里没有图片、PDF 或 PPTX', []
    return 'reference', f'目录只含参考文件（{len(refs)} 个）', refs


def _named_font_verdict(name):
    hit = next((p for p in PROPRIETARY if re.search(r'(?<![A-Za-z])' + re.escape(p) + r'(?![A-Za-z])', name, re.I)), None)
    return ('不能分发', f'商业/系统字体家族 {hit}') if hit else ('待查', '只有名字，授权待查')


def analyse_reference(source, files, mode_reason):
    source = Path(source)
    root = source if source.is_dir() else source.parent
    rows, font_rows, named, pages_known, unknown, sizes = [], [], {}, 0, 0, {}
    licences = [p for p in files if p.suffix.lower() == '.txt']

    def add_named(name, where):
        entry = named.setdefault(name, {'name': name, 'from': set()})
        entry['from'].add(where)

    for path in files:
        rel = path.relative_to(root).as_posix()
        ext = path.suffix.lower()
        row = {'file': rel, 'bytes': path.stat().st_size, 'sha256': sha256_file(path)}
        try:
            if ext in REF_IMAGE_EXT:
                with open(path, 'rb') as f:
                    head = f.read(256 * 1024)
                wh = image_size(head)
                row.update({'kind': 'image', 'pages': 1, 'width': wh[0] if wh else None, 'height': wh[1] if wh else None, 'aspect': aspect_name(*wh) if wh else '-'})
                pages_known += 1
                key = f'{wh[0]}×{wh[1]}' if wh else '读不出'
                sizes[key] = sizes.get(key, 0) + 1
            elif ext == '.pdf':
                inv = pdf_inventory(path)
                row.update({'kind': 'pdf', 'pages': inv['pages'], 'pdf': inv})
                if inv['pages']:
                    pages_known += inv['pages']
                else:
                    unknown += 1
                for w, h in inv['page_size_pt']:
                    key = f'{w:g}×{h:g}pt'
                    sizes[key] = sizes.get(key, 0) + 1
                for f in inv['fonts'] or []:
                    add_named(f['name'], rel + ('（嵌入）' if f['embedded'] else '（未嵌入）'))
            elif ext == '.pptx' or ext in MACRO_PPT_EXT:
                inv = pptx_inventory(path)
                row.update({'kind': 'pptx', 'pages': inv['slides'], 'pptx': inv})
                pages_known += inv['slides']
                if inv['slide_size_px'] and inv['slides']:
                    key = f'{inv["slide_size_px"][0]}×{inv["slide_size_px"][1]}px（{aspect_name(*inv["slide_size_px"])}）'
                    sizes[key] = sizes.get(key, 0) + inv['slides']
                for k, v in inv['fonts']['theme'].items():
                    add_named(v, f'{rel} 主题 {k}')
                for f in inv['fonts']['used']:
                    add_named(f['name'], f'{rel} 正文')
                for f in inv['embedded_fonts']:
                    font_rows.append({'file': f'{rel}:{f["part"]}', 'family': f['family'] or '读不出', 'verdict': _named_font_verdict(f['family'] or '')[0] if f['family'] else '待查',
                                      'why': '嵌在 PPTX 里', 'fs_type': f['fs_type'], 'fs_type_label': f['fs_type_label']})
            elif ext in FONT_EXT:
                info = font_names(path.read_bytes())
                fam = info.get('family') or path.stem
                verdict, why = _named_font_verdict(fam)
                lic = ' '.join(filter(None, [info.get('license'), info.get('license_url')]))
                if verdict != '不能分发':
                    if NONFREE_LICENSE.search(lic):
                        verdict, why = '不能分发', '字体自带的授权说明不允许再分发'
                    elif FREE_LICENSE.search(lic):
                        verdict, why = '可分发', '字体 name 表写明开源授权'
                    elif any(FREE_LICENSE.search(p.read_text('utf-8', 'replace')[:4000]) and _norm(fam)[:4] in _norm(p.name) for p in licences):
                        verdict, why = '可分发', '同目录许可文件'
                font_rows.append({'file': rel, 'family': fam, 'verdict': verdict, 'why': why, 'fs_type': info.get('fs_type'),
                                  'fs_type_label': fstype_label(info.get('fs_type')) if info else '读不出'})
                row.update({'kind': 'font'})
        except (Refused, zipfile.BadZipFile, OSError, ET.ParseError, ValueError, TypeError, KeyError, IndexError, struct.error) as e:
            row.update({'kind': ext.lstrip('.'), 'error': str(e)})
        rows.append(row)
    named_rows = []
    for name, entry in sorted(named.items()):
        verdict, why = _named_font_verdict(name)
        named_rows.append({'name': name, 'verdict': verdict, 'why': why, 'from': sorted(entry['from'])[:4]})
    report = {'mode': 'reference', 'mode_reason': mode_reason, 'source': source.name, 'files': rows,
              'pages_total': pages_known, 'pages_unknown_files': unknown, 'sizes': sizes,
              'fonts': {'files': font_rows, 'named': named_rows}}
    report['verdict'] = reference_verdict(report)
    return report


def reference_verdict(r):
    usable = [f for f in r['files'] if f.get('kind') in ('image', 'pdf', 'pptx') and not f.get('error') and f.get('pages') != 0]
    blockers, decide = [], []
    for f in r['files']:
        if f.get('kind') in ('pdf', 'pptx') and f.get('pages') == 0 and not f.get('error'):
            blockers.append(f'{f["file"]} 没有页面（PPTX 只有母版/版式时，先在 PowerPoint 里用各版式各做一页）')
    errors = [f for f in r['files'] if f.get('error')]
    for f in errors:
        blockers.append(f'{f["file"]} 读不了：{f["error"]}')
    if len(r['sizes']) > 1:
        decide.append('页面尺寸不一（' + '、'.join(f'{k} ×{v}' for k, v in r['sizes'].items()) + '）：确认哪些是要蒸馏的页')
    bad = [f['family'] for f in r['fonts']['files'] if f['verdict'] == '不能分发']
    if bad:
        blockers.append('字体不能分发：' + '、'.join(dict.fromkeys(bad)) + '（样式包只写字体名、不打包，并换成可分发字体）')
    named_bad = [f['name'] for f in r['fonts']['named'] if f['verdict'] == '不能分发']
    if named_bad:
        decide.append('引用了商业/系统字体：' + '、'.join(named_bad) + '（样式包要换成可分发字体，或确认客户机都有）')
    unknown = [f['family'] for f in r['fonts']['files'] if f['verdict'] == '待查'] + [f['name'] for f in r['fonts']['named'] if f['verdict'] == '待查']
    if unknown:
        decide.append('字体授权待查：' + '、'.join(dict.fromkeys(unknown)))
    if any(f.get('kind') == 'pdf' and f.get('pdf', {}).get('text_layer') is False for f in usable):
        decide.append('PDF 没有文字层：文字只能 agent 看图识读，逐条标来源')
    decide.insert(0, '设计归属待用户确认：这套设计归谁、能否给客户用（确认前 meta.design_source.cleared 保持 false）')
    return {'can_extract': bool(usable), 'blockers': blockers, 'decide': decide}


def markdown_reference(r):
    v = r['verdict']
    out = [f'# 预检报告（参考稿）：{r["source"]}', '', f'模式：参考稿模式，{r["mode_reason"]}。只读文件头、XML 与 poppler 输出，不执行任何内容。', '',
           f'合计 **{r["pages_total"]}** 页' + (f'，另有 {r["pages_unknown_files"]} 个文件页数未知' if r['pages_unknown_files'] else '') + '；尺寸：'
           + ('、'.join(f'{k} ×{n}' for k, n in r['sizes'].items()) or '-'), '',
           '| 文件 | 类型 | 页数 | 尺寸 / 画幅 | 说明 | sha256 |', '|---|---|---|---|---|---|']
    for f in r['files']:
        kind, note, size = f.get('kind', '?'), '', '-'
        if f.get('error'):
            note = f'**读不了**：{f["error"]}'
        elif kind == 'image':
            size = f'{f["width"]}×{f["height"]}（{f["aspect"]}）' if f.get('width') else '读不出'
        elif kind == 'pdf':
            p = f['pdf']
            sz = sorted({f'{w:g}×{h:g}pt' for w, h in p['page_size_pt']})
            size = '、'.join(sz[:3]) or '-'
            note = ('文字层：' + ('有' if p['text_layer'] else '无' if p['text_layer'] is False else '未知') + (f'（{p["text_chars"]} 字）' if p['text_chars'] else '')
                    + (f'；字体 {len(p["fonts"])} 个' if p['fonts'] is not None else '') + (f'；{p["note"]}' if p.get('note') else ''))
        elif kind == 'pptx':
            p = f['pptx']
            size = f'{p["slide_size_px"][0]}×{p["slide_size_px"][1]}px（{aspect_name(*p["slide_size_px"])}）' if p['slide_size_px'] else '-'
            note = (f'母版 {p["masters"]}，用到版式 {len(p["layouts"])} 个（' + '、'.join(f'{x["name"]}×{x["used_by"]}' for x in p['layouts'][:6]) + '）；媒体 '
                    + ('、'.join(f'{k} {n}' for k, n in p['media'].items()) or '无') + f'；嵌入字体 {len(p["embedded_fonts"])}')
        out.append(f'| {f["file"]} | {kind} | {f.get("pages", "-") if f.get("pages") is not None else "未知"} | {size} | {note or "-"} | `{f["sha256"][:12]}` |')
    out += ['', '## 字体', '']
    if r['fonts']['files']:
        out += ['| 字体文件 | 家族 | 判定 | 原因 | OS/2 fsType |', '|---|---|---|---|---|']
        for f in r['fonts']['files']:
            fs = f'0x{f["fs_type"]:04X} {f["fs_type_label"]}' if f['fs_type'] is not None else f['fs_type_label']
            out.append(f'| {f["file"]} | {f["family"]} | {"**" + f["verdict"] + "**" if f["verdict"] != "可分发" else f["verdict"]} | {f["why"]} | {fs} |')
        out.append('')
    if r['fonts']['named']:
        out += ['只有名字的字体（PDF / PPTX 里引用，没有字体文件）：', '', '| 字体名 | 判定 | 出现在 |', '|---|---|---|']
        for f in r['fonts']['named']:
            out.append(f'| {f["name"]} | **{f["verdict"]}**（{f["why"]}） | {"；".join(f["from"])} |')
    if not r['fonts']['files'] and not r['fonts']['named']:
        out.append('- 没有字体文件，也没有读到字体名（图片只能看图推断字体，写进 meta.fonts 时注明「看图推断」）')
    out += ['', '## 结论', '', f'- 能否进入提取阶段：**{"能" if v["can_extract"] else "不建议"}**']
    out += [f'- 阻碍：{b}' for b in v['blockers']] or ['- 阻碍：无']
    out += [f'- 要人决定：{x}' for x in v['decide']]
    return '\n'.join(out) + '\n'


def run(source):
    source = Path(source)
    if source.is_dir():
        return analyse(Package(str(source)), source.name)
    if not zipfile.is_zipfile(source):
        raise Refused(f'{source} 既不是目录也不是 zip')
    tmp = Path(tempfile.mkdtemp(prefix='style-preflight-'))
    try:
        unpack(source, tmp)
        entries = [p for p in tmp.iterdir()]
        root = entries[0] if len(entries) == 1 and entries[0].is_dir() else tmp
        return analyse(Package(str(root)), source.name)
    finally:
        make_writable(tmp)
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('source', type=Path)
    ap.add_argument('--out', type=Path, required=True, help='Markdown 报告路径')
    ap.add_argument('--json', type=Path, help='同时写一份 JSON')
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument('--ref', action='store_true', help='强制按参考稿（图片 / PDF / PPTX）预检')
    mode.add_argument('--package', action='store_true', help='强制按源包预检')
    a = ap.parse_args()
    try:
        kind, why, files = classify(a.source, 'ref' if a.ref else 'package' if a.package else None)
        if kind == 'reference':
            if not files:
                raise Refused(f'{a.source} 里没有图片、PDF、PPTX 或字体文件')
            report = analyse_reference(a.source, files, why)
        else:
            report = run(a.source)
            report['mode'], report['mode_reason'] = 'package', why
    except (Refused, zipfile.BadZipFile, OSError) as e:
        print(f'预检拒绝：{e}', file=sys.stderr)
        return 2
    except Exception as e:  # noqa: BLE001 - malformed untrusted input must end in a clear message, not a traceback
        print(f'预检失败（源包内容无法解析）：{type(e).__name__}: {e}', file=sys.stderr)
        return 2
    if kind == 'reference':
        a.out.write_text(markdown_reference(report), encoding='utf-8')
        summary = {'report': str(a.out), 'mode': 'reference', 'mode_reason': why, 'pages': report['pages_total'],
                   'sizes': report['sizes'], 'can_extract': report['verdict']['can_extract']}
    else:
        a.out.write_text(markdown(report).replace('\n', f'\n\n模式：源包模式，{why}。\n', 1), encoding='utf-8')
        summary = {'report': str(a.out), 'mode': 'package', 'mode_reason': why, 'format': report['format']['format'], 'runtime': report['runtime']['verdict'],
                   'photos': report['photos']['total'], 'min_pt': report['min_font'].get('min_pt'), 'can_extract': report['verdict']['can_extract']}
    if a.json:
        a.json.write_text(json.dumps(report, ensure_ascii=False, indent=1, default=sorted), encoding='utf-8')
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
