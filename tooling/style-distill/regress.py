#!/usr/bin/env python3
"""Offline renderer regression against a golden (style-distill tooling; developer machines only).

    <python with python-pptx> -I regress.py golden [--renderer <brand_deck.py>] [--style default] [--out <golden.json>]
    <python with python-pptx> -I regress.py check  [--renderer <brand_deck.py>] [--style default] [--golden <golden.json>]

Builds the style's example.json (with its year pinned) and every fixtures/<style>-regression*.json with the
renderer, then records (golden) or compares (check) for each deck:
  report  the build JSON without its absolute `output` path (warnings and their order included);
  parts   sha256 of every normalised part (inspect_pptx.parts: C14N XML, embedded workbooks opened, no docProps);
  dump    a readable shape list, only used to point at what changed.
check exits 0 when report and parts all match, else prints the differences and exits 1.
"""
import argparse
import importlib.util
import json
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
SKILL = REPO / 'packages' / 'dsh-adapter' / 'skills' / 'shimo-pptx'
PINNED_YEAR = 2026


def _inspect():
    spec = importlib.util.spec_from_file_location('style_distill_inspect', HERE / 'inspect_pptx.py')
    module = importlib.util.module_from_spec(spec)
    sys.dont_write_bytecode = True
    spec.loader.exec_module(module)
    return module


def decks(style):
    """(name, spec path) pairs: the pack's example and this style's regression fixtures."""
    found = [('example', SKILL / 'styles' / style / 'example.json')]
    found += [(p.stem, p) for p in sorted((HERE / 'fixtures').glob(f'{style}-regression*.json'))]
    return found


def _absolute(node, base):
    """The spec with every relative path that names a file beside it made absolute (so the copy can live elsewhere)."""
    if isinstance(node, dict):
        return {k: _absolute(v, base) for k, v in node.items()}
    if isinstance(node, list):
        return [_absolute(v, base) for v in node]
    if isinstance(node, str) and node and not Path(node).is_absolute() and '\n' not in node and (base / node).is_file():
        return str((base / node).resolve())
    return node


def build(renderer, style, name, spec_path, work):
    spec = _absolute(json.loads(spec_path.read_text(encoding='utf-8')), spec_path.parent)
    spec.setdefault('year', PINNED_YEAR)
    # The pinned copy lives in the temporary work folder, never inside the (shipped) skill or fixtures folders.
    pinned = work / f'{name}.json'
    pinned.write_text(json.dumps(spec, ensure_ascii=False), encoding='utf-8')
    out = work / f'{name}.pptx'
    run = subprocess.run([sys.executable, '-I', str(renderer), 'build', '--spec', str(pinned), '--out', str(out), '--style', style, '--overwrite'],
                         capture_output=True, text=True)
    if run.returncode != 0:
        raise SystemExit(f'{name}: renderer failed ({run.returncode}): {run.stderr.strip()}')
    report = json.loads(run.stdout)
    report.pop('output', None)
    return out, report


def record(renderer, style):
    inspect = _inspect()
    result = {}
    with tempfile.TemporaryDirectory(prefix='style-regress-') as tmp:
        for name, path in decks(style):
            out, report = build(renderer, style, name, path, Path(tmp))
            result[name] = {'report': report, 'parts': inspect.parts(out), 'dump': inspect.dump(out)}
    return result


def env():
    import lxml
    import PIL
    import pptx
    import xlsxwriter
    return {'python-pptx': pptx.__version__, 'lxml': lxml.__version__, 'Pillow': PIL.__version__, 'XlsxWriter': xlsxwriter.__version__,
            'python': sys.version.split()[0]}


def _first_dump_diff(old, new):
    for a, b in zip(old, new):
        if a != b:
            for i, (sa, sb) in enumerate(zip(a['shapes'], b['shapes'])):
                if sa != sb:
                    return f'slide {a["slide"]} shape {i}: golden {json.dumps(sa, ensure_ascii=False)} / now {json.dumps(sb, ensure_ascii=False)}'
            return f'slide {a["slide"]}: golden {len(a["shapes"])} shapes, notes {a["notes"]!r} / now {len(b["shapes"])} shapes, notes {b["notes"]!r}'
    return f'slide count golden {len(old)} / now {len(new)}' if len(old) != len(new) else 'dump identical (difference is below the dump level)'


def compare(golden, now):
    problems = []
    for name in sorted(set(golden) | set(now)):
        if name not in golden or name not in now:
            problems.append(f'{name}: deck only in {"golden" if name in golden else "current run"}')
            continue
        g, n = golden[name], now[name]
        if g['report'] != n['report']:
            problems.append(f'{name}: report differs\n  golden {json.dumps(g["report"], ensure_ascii=False)}\n  now    {json.dumps(n["report"], ensure_ascii=False)}')
        changed = sorted(k for k in set(g['parts']) | set(n['parts']) if g['parts'].get(k) != n['parts'].get(k))
        if changed:
            problems.append(f'{name}: {len(changed)} part(s) differ: {", ".join(changed[:12])}{" …" if len(changed) > 12 else ""}\n  {_first_dump_diff(g["dump"], n["dump"])}')
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('cmd', choices=['golden', 'check'])
    ap.add_argument('--renderer', type=Path, default=SKILL / 'scripts' / 'brand_deck.py')
    ap.add_argument('--style', default='default')
    ap.add_argument('--golden', '--out', dest='golden', type=Path)
    a = ap.parse_args()
    golden_path = a.golden or HERE / 'fixtures' / f'{a.style}-golden.json'
    now = record(a.renderer.resolve(), a.style)
    if a.cmd == 'golden':
        golden_path.write_text(json.dumps({'env': env(), 'renderer': str(a.renderer), 'year': PINNED_YEAR, 'decks': now},
                                          ensure_ascii=False, indent=1) + '\n', encoding='utf-8')
        print(f'wrote {golden_path} ({len(now)} decks)')
        return 0
    golden = json.loads(golden_path.read_text(encoding='utf-8'))
    problems = compare(golden['decks'], now)
    if golden['env'] != env():
        print(f'note: library versions differ from the golden: golden {golden["env"]} / now {env()}', file=sys.stderr)
    if problems:
        print('\n'.join(problems))
        return 1
    print(f'identical: {", ".join(sorted(now))} (report + {sum(len(d["parts"]) for d in now.values())} parts)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
