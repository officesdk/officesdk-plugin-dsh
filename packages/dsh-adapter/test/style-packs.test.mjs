// Style packs against tooling/style-distill/validate.mjs, the expression grammar on both sides, and the offline
// renderer against its golden. Runs in CI (no Python there: the Python parts skip); imports nothing that needs DSH's
// peer packages, unlike styles.test.mjs.
import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { cpSync, symlinkSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { crc32, deflateSync } from 'node:zlib';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listZip, readZipEntry } from '../../core/src/zip.mjs';
import { parseExpr } from '../src/host/style-expr.mjs';
import { liveLayouts, readLayouts } from '../src/host/style-layouts.mjs';
import { DEFAULT_SKILL, STAGING_STYLES, findPython, validatePacks } from '../../../tooling/style-distill/validate.mjs';

const TOOLING = fileURLToPath(new URL('../../../tooling/style-distill/', import.meta.url));
const CASES = JSON.parse(readFileSync(join(TOOLING, 'fixtures', 'expr-cases.json'), 'utf8'));
const PYTHON = findPython();

test('every style pack passes the validator (renderer check when a Python with python-pptx is here)', () => {
  const result = validatePacks({ python: PYTHON ?? false, allowNoPython: true });
  for (const pack of result.packs) assert.deepEqual(pack.problems, [], pack.id);
  const pack = result.packs.find(item => item.id === 'default');
  assert.equal(pack.renderer, PYTHON ? 'checked' : 'skipped');
  assert.deepEqual(pack.stats.photos, { slots: 7, required: 1, per_slide_max: 3 });
  assert.equal(pack.stats.previews, 3, 'default shows three sample pages in the template gallery');
  assert.ok(pack.stats.minTextPt >= 9 && pack.stats.minChartPt >= 8);
  // The repository ships the default pack alone (the blue house style).
  assert.deepEqual(result.packs.map(item => item.id), ['default']);
  assert.equal(JSON.parse(readFileSync(join(DEFAULT_SKILL, 'styles', 'default', 'style.json'), 'utf8')).meta.design_source.cleared, true);
});

test('the validator names each broken part of a pack', () => {
  const skill = mkdtempSync(join(tmpdir(), 'style-validate-test-'));
  const pack = join(skill, 'styles', 'default');
  const json = file => JSON.parse(readFileSync(join(pack, file), 'utf8'));
  const write = (file, value) => writeFileSync(join(pack, file), typeof value === 'string' ? value : JSON.stringify(value));
  const cases = [
    ['style.json', style => { delete style.meta; }, /缺少 meta/],
    ['style.json', style => { style.meta.name = { zh: '蓝' }; }, /meta\.name 要有 zh 和 en/],
    ['style.json', style => { style.meta.thumb = '../x.png'; }, /meta\.thumb/],
    ['style.json', style => { style.meta.requires.photos.slots = 9; }, /meta\.requires\.photos\.slots=9，按 layouts\.json 算是 7/],
    ['style.json', style => { style.meta.design_source.cleared = false; }, /cleared 为 false/],
    ['style.json', style => { style.meta.fonts[0] = { ...style.meta.fonts[0], license: 'Proprietary', bundled: true }; }, /不能随包分发/],
    ['style.json', style => { style.palettes['color-1'].primary = '#123456'; }, /模板主题 dk1=2826C9，style\.json 是 123456/],
    ['style.json', style => { style.template.sizes_pt.label = 8; }, /实时字号 template\.sizes_pt\.label=8pt/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[5].size = 10; }, /cover\.objects\[5\] 字号最小 7\.50pt/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[3].size = 'fit(title, 14, 700)'; }, /cover\.objects\[3\] 字号最小 6\.51pt/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[1].text = 'missing_name'; }, /cover\.objects\[1\]\.text 用到未定义的名字 missing_name/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[1].text = "d.__class__"; }, /只允许 \.get\(\)/],
    ['layouts.json', desc => { desc.offline.layouts.overview.objects[6].max = 40; }, /overview 最多 \d+ 个文本框，超过 30/],
    ['layouts.json', desc => { desc.offline.layouts.chart.objects[2].fonts = { value_axis: 9 }; }, /图表 value_axis 字号 6\.75pt/],
    ['layouts.json', desc => { desc.live.layouts[0].boxes.title = [0, 0, 10, 10]; }, /模板版式「封面」title 的位置/],
    ['layouts.json', desc => { desc.live.slots[0].box = [1, 2, 3, 4]; }, /layouts\.md 没有写实时槽位 body 的位置 1, 2, 3, 4/],
    ['layouts.md', doc => doc.replace('36, 100, 312, 270', ''), /layouts\.md 没有写实时槽位 left-image/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[1].color = "__import__('os')"; }, /Host 读不了 layouts\.json：.*__import__/],
    ['layouts.json', desc => { desc.offline.layouts.cover.background = "open('x')"; }, /Host 读不了 layouts\.json：.*不允许的函数 open/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[2].color = 'noSuchRole'; }, /cover\.objects\[2\]\.color 的颜色 noSuchRole 在配色 color-1 里没有/],
    ['layouts.json', desc => { desc.offline.layouts.overview.objects[1].text = "d['titel']"; }, /overview\.objects\[1\]\.text 读了字段 d\['titel'\]/],
    ['layouts.json', desc => { desc.offline.layouts.overview.objects[2].when = "d.get('leed')"; }, /读了字段 d\['leed'\]/],
    ['layouts.json', desc => { desc.offline.layouts.overview.objects[1].size = 14; }, /overview\.objects\[1\] 字号最小 6\.51pt/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[5].size = 'max(4, 1)'; }, /cover\.objects\[5\] 字号最小 3\.00pt/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[5].size = "len(d.get('year', ''))"; }, /推算不出下限/],
    ['layouts.json', desc => { delete desc.offline.layouts.overview.objects[6].max; }, /repeat 要写 max/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[3].size = '40 - (10 if closing else 30)'; }, /cover\.objects\[3\] 的字号 .* 推算不出下限/],
    ['layouts.json', desc => { desc.offline.layouts.cover.objects[3].size = '1600 / fit(title, 110, 700)'; }, /推算不出下限/],
    ['layouts.json', desc => { desc.offline.layouts.cover.background = 5; }, /background 必须是颜色字符串/],
    ['thumb.png', () => 'not a png', /thumb\.png 不是 PNG/],
    ['style.json', style => { style.meta.previews = ['preview-1.png']; }, /meta\.previews 可不写；写了要是 2–4 个/],
    ['style.json', style => { style.meta.previews = ['../a.png', 'b.png']; }, /meta\.previews 可不写/],
    ['style.json', style => { style.meta.previews = ['p-a.png', 'p-b.png']; }, /meta\.previews 里的 p-a\.png 不存在/],
  ];
  try {
    cpSync(DEFAULT_SKILL, skill, { recursive: true });
    // Only default: the cases read packs[0], and the other registered packs are checked by the test above.
    for (const id of readdirSync(join(skill, 'styles'))) if (id !== 'default') rmSync(join(skill, 'styles', id), { recursive: true, force: true });
    for (const [file, edit, message] of cases) {
      const original = readFileSync(join(pack, file));
      const value = file.endsWith('.json') ? json(file) : original.toString('utf8');
      const changed = edit(value);
      write(file, changed ?? value);
      const result = validatePacks({ skill, python: false, allowNoPython: true });
      assert.equal(result.ok, false, `${file}: ${message}`);
      assert.ok(result.packs[0].problems.some(problem => message.test(problem)), `${message} not in ${JSON.stringify(result.packs[0].problems)}`);
      writeFileSync(join(pack, file), original);
    }
    assert.ok(validatePacks({ skill, python: false, allowNoPython: true }).ok, 'restored pack passes again');
    // A second pack may not take the first one's alias or id.
    cpSync(pack, join(skill, 'styles', 'second'), { recursive: true });
    const second = JSON.parse(readFileSync(join(skill, 'styles', 'second', 'style.json'), 'utf8'));
    second.name = second.meta.id = 'second';
    second.meta.aliases = ['默认'];
    writeFileSync(join(skill, 'styles', 'second', 'style.json'), JSON.stringify(second));
    const clash = validatePacks({ skill, python: false, allowNoPython: true });
    assert.ok(clash.packs.flatMap(item => item.problems).some(problem => /别名或 id「默认」同时属于 default 和 second/.test(problem)));
    // Without a Python and without --allow-no-python, the renderer check fails instead of passing silently.
    assert.ok(validatePacks({ skill, style: 'default', python: false }).packs[0].problems.some(problem => /找不到带 python-pptx 的 Python/.test(problem)));
    assert.equal(findPython('/no/such/python'), null, 'an explicit python that does not work is not replaced by another');
  } finally { rmSync(skill, { recursive: true, force: true }); }
});

test('every staging pack passes the validator with the staging rules (cleared=false is only a note there)', () => {
  const result = validatePacks({ styles: STAGING_STYLES, python: PYTHON ?? false, allowNoPython: true });
  assert.equal(result.staging, true);
  // Staging may hold no pack at all (every distilled pack so far went straight to styles/): that is still a pass.
  assert.ok(Array.isArray(result.packs) && result.ok);
  for (const pack of result.packs) {
    assert.deepEqual(pack.problems, [], pack.id);
    assert.equal(pack.renderer, PYTHON ? 'checked' : 'skipped', pack.id);
  }
});

test('staging: cleared=false is a note, the other rules stay, and names may not clash with registered packs', () => {
  const work = mkdtempSync(join(tmpdir(), 'style-staging-test-'));
  const staging = join(work, 'staging');
  const pack = join(staging, 'trial');
  try {
    cpSync(join(DEFAULT_SKILL, 'styles', 'default'), pack, { recursive: true });
    const style = JSON.parse(readFileSync(join(pack, 'style.json'), 'utf8'));
    const write = value => writeFileSync(join(pack, 'style.json'), JSON.stringify(value));
    Object.assign(style, { name: 'trial' });
    Object.assign(style.meta, { id: 'trial', aliases: ['试用样式'], design_source: { ...style.meta.design_source, cleared: false } });
    write(style);
    const ok = validatePacks({ styles: staging, python: PYTHON ?? false, allowNoPython: true });
    assert.equal(ok.ok, true, JSON.stringify(ok.packs[0].problems));
    assert.ok(ok.packs[0].notes.some(note => /待登记/.test(note)));
    if (PYTHON) assert.equal(ok.packs[0].renderer, 'checked', 'the renderer reads the pack from the staging folder');
    // Other rules are the same as under styles/.
    write({ ...style, template: { ...style.template, sizes_pt: { ...style.template.sizes_pt, label: 8 } } });
    assert.ok(validatePacks({ styles: staging, python: false, allowNoPython: true }).packs[0].problems.some(problem => /实时字号 template\.sizes_pt\.label=8pt/.test(problem)));
    // An alias or id of a registered pack is taken.
    write({ ...style, meta: { ...style.meta, aliases: ['默认'] } });
    assert.ok(validatePacks({ styles: staging, python: false, allowNoPython: true }).packs[0].problems.some(problem => /别名或 id「默认」同时属于 default（已登记） 和 trial/.test(problem)));
    write(style);
    cpSync(pack, join(staging, 'default'), { recursive: true });
    const named = validatePacks({ styles: staging, python: false, allowNoPython: true });
    assert.ok(named.packs.find(item => item.id === 'default').problems.some(problem => /「default」同时属于 default（已登记） 和 default/.test(problem)));
    rmSync(join(staging, 'default'), { recursive: true });
    // The same pack under a skill's styles/ is refused: uncleared packs do not get registered.
    const skill = join(work, 'skill');
    cpSync(DEFAULT_SKILL, skill, { recursive: true });
    cpSync(pack, join(skill, 'styles', 'trial'), { recursive: true });
    const registered = validatePacks({ skill, style: 'trial', python: false, allowNoPython: true });
    assert.equal(registered.staging, false);
    assert.ok(registered.packs[0].problems.some(problem => /cleared 为 false/.test(problem)));
  } finally { rmSync(work, { recursive: true, force: true }); }
});

/** A staging copy of default whose cover also draws the objects and keys new styles need (shape, gradient, font roles…). */
function extendedPack(work) {
  const staging = join(work, 'staging');
  const pack = join(staging, 'extended');
  cpSync(join(DEFAULT_SKILL, 'styles', 'default'), pack, { recursive: true });
  const style = JSON.parse(readFileSync(join(pack, 'style.json'), 'utf8'));
  style.name = style.meta.id = 'extended';
  style.meta.aliases = ['扩展测试'];
  style.meta.design_source.cleared = false;
  style.fonts.roles = { serif: { latin: 'Source Serif 4', latin_light: 'Source Serif 4 Light', cjk: 'Noto Serif SC' } };
  style.meta.fonts.push(...['Source Serif 4', 'Source Serif 4 Light', 'Noto Serif SC'].map(family => ({ family, role: 'serif', license: 'OFL-1.1', source: 'test', bundled: false })));
  style.template.colors = { content_bg: 'ink', title: 'paper', body: 'paper' };
  style.template.dark = true;
  const desc = JSON.parse(readFileSync(join(pack, 'layouts.json'), 'utf8'));
  desc.offline.layouts.cover.objects.push(
    { type: 'shape', kind: 'triangle', x: 10, y: 10, w: 100, h: 200, fill: 'ink', alpha: 0.5, flip_v: true, name: 'tri' },
    { type: 'rect', x: 0, y: 300, w: 960, h: 240, fill: null, name: 'fade', gradient: { angle: 180, stops: [[0, 'ink', 0], [1, '#000000', 0.9]] } },
    { type: 'text', text: "'Serif Title\\n中文标题'", x: 40, y: 40, w: 600, size: 24, color: 'paper', font: 'serif', italic: true, bold: true, spacing: 1.5, name: 'serif-text' });
  writeFileSync(join(pack, 'style.json'), JSON.stringify(style, null, 2));
  writeFileSync(join(pack, 'layouts.json'), JSON.stringify(desc, null, 2));
  return { staging, pack, style, desc };
}

test('new object keys: the validator accepts them and names each misuse', () => {
  const work = mkdtempSync(join(tmpdir(), 'style-extended-test-'));
  try {
    const { staging, pack, style, desc } = extendedPack(work);
    const run = () => validatePacks({ styles: staging, python: false, allowNoPython: true }).packs[0].problems;
    // template.pptx is still default's, so only template-theme facts may differ; nothing about the new keys.
    assert.deepEqual(run().filter(problem => !/模板/.test(problem)), []);
    const cover = desc.offline.layouts.cover.objects;
    const at = cover.length;
    const cases = [
      [d => { d.offline.layouts.cover.objects[at - 3].kind = 'star'; }, null, /kind="star"，可选 triangle/],
      [d => { d.offline.layouts.cover.objects[at - 2].gradient.stops = [[0, 'ink', 0]]; }, null, /gradient 要写/],
      [d => { d.offline.layouts.cover.objects[at - 2].gradient.stops[1][1] = 'noSuchRole'; }, null, /gradient\.stops\[1\] 的颜色 noSuchRole/],
      [d => { d.offline.layouts.cover.objects[at - 1].font = 'mono'; }, null, /font="mono" 不在 style\.json fonts\.roles 里/],
      [d => { d.offline.layouts.cover.objects[at - 1].spacing = 'nope'; }, null, /用到未定义的名字 nope/],
      [null, s => { s.template.colors.title = 'noSuchRole'; }, /template\.colors\.title=noSuchRole/],
      [null, s => { s.template.colors.footer = 'ink'; }, /template\.colors\.footer 不认识/],
      [null, s => { s.fonts.roles.serif = { latin: 'X' }; }, /fonts\.roles\.serif 要有 latin 和 cjk/],
      [null, s => { s.template.colors.content_bg = 'paper'; }, /模板版式「标题和内容」背景 无，按 template\.colors 应为 paper/],
      [null, s => { s.template.colors.title = 'secondary'; }, /模板版式「标题和内容」title 文字颜色 [0-9A-F]{6}，按 template\.colors 应为 secondary/],
      [null, s => { s.meta.fonts = s.meta.fonts.filter(font => font.family !== 'Noto Serif SC'); }, /meta\.fonts 没有列出 style\.json 用到的字体 Noto Serif SC/],
    ];
    for (const [editDesc, editStyle, message] of cases) {
      const d = structuredClone(desc), s = structuredClone(style);
      editDesc?.(d); editStyle?.(s);
      writeFileSync(join(pack, 'layouts.json'), JSON.stringify(d));
      writeFileSync(join(pack, 'style.json'), JSON.stringify(s));
      assert.ok(run().some(problem => message.test(problem)), `${message} not in ${JSON.stringify(run())}`);
    }
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('new object keys: the renderer and the template builder write them as OOXML', { skip: !PYTHON && 'no python with python-pptx' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'style-extended-render-'));
  try {
    const { staging, pack } = extendedPack(work);
    const scripts = join(DEFAULT_SKILL, 'scripts');
    const tpl = spawnSync(PYTHON, ['-I', join(scripts, 'build_template.py'), '--style', 'extended', '--styles-dir', staging], { encoding: 'utf8' });
    assert.equal(tpl.status, 0, tpl.stderr);
    const result = validatePacks({ styles: staging, python: PYTHON, allowNoPython: true });
    assert.deepEqual(result.packs[0].problems, [], 'the rebuilt template matches; the renderer reads the staging pack');
    const out = join(work, 'deck.pptx');
    const built = spawnSync(PYTHON, ['-I', join(scripts, 'brand_deck.py'), 'build', '--spec', join(pack, 'example.json'), '--out', out, '--style', 'extended', '--styles-dir', staging], { encoding: 'utf8' });
    assert.equal(built.status, 0, built.stderr);
    const archive = listZip(readFileSync(out));
    const xml = name => readZipEntry(archive, archive.entries.find(entry => entry.name === name), 8 * 1024 * 1024).toString('utf8');
    const slide = xml('ppt/slides/slide1.xml');
    const shape = name => new RegExp(`<p:sp>(?:(?!</p:sp>).)*?name="${name}"(?:(?!</p:sp>).)*</p:sp>`, 's').exec(slide)?.[0] ?? '';
    const tri = shape('tri');
    assert.match(tri, /<a:prstGeom prst="triangle"/);
    assert.match(tri, /<a:xfrm flipV="1"/);
    assert.match(tri, /<a:alpha val="50000"\/>/);
    const fade = shape('fade');
    assert.match(fade, /<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="[0-9A-F]{6}"><a:alpha val="0"\/><\/a:srgbClr><\/a:gs><a:gs pos="100000"><a:srgbClr val="000000"><a:alpha val="90000"\/>/);
    assert.match(fade, /<a:lin ang="5400000" scaled="0"\/>/, 'CSS 180deg (top to bottom) is OOXML 90°');
    const text = shape('serif-text');
    const runs = [...text.matchAll(/<a:r>(.*?)<\/a:r>/gs)].map(m => m[1]);
    assert.equal(runs.length, 2, 'one box, two paragraphs');
    assert.match(runs[0], /spc="150"/, '1.5pt letter spacing is 150 hundredths');
    assert.match(runs[0], / b="1"/);
    assert.match(runs[0], / i="1"/);
    assert.match(runs[0], /<a:latin typeface="Source Serif 4"\/>/);
    assert.doesNotMatch(runs[1], / i="1"/, 'no synthetic italic on the Chinese line');
    assert.match(runs[1], /<a:latin typeface="Noto Serif SC"\/>/);
    // Template: dark colour mapping, content backgrounds and title colour from template.colors.
    const tarchive = listZip(readFileSync(join(pack, 'template.pptx')));
    const txml = name => readZipEntry(tarchive, tarchive.entries.find(entry => entry.name === name), 4 * 1024 * 1024).toString('utf8');
    assert.match(txml('ppt/slideMasters/slideMaster1.xml'), /<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/);
    const style = JSON.parse(readFileSync(join(pack, 'style.json'), 'utf8'));
    const P = style.palettes[style.default_palette];
    const content = tarchive.entries.map(entry => entry.name).filter(name => /slideLayouts\/slideLayout\d+\.xml$/.test(name)).map(txml).find(x => /name="标题和内容"/.test(x));
    assert.match(content, new RegExp(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${P.ink.slice(1).toUpperCase()}"`));
    assert.ok(content.includes(`<a:srgbClr val="${P.paper.slice(1).toUpperCase()}"`), 'title and body take paper');
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('presentation_create\'s style.layouts is the live part of layouts.json, read without DSH packages (CI)', () => {
  const live = JSON.parse(readFileSync(join(DEFAULT_SKILL, 'styles', 'default', 'layouts.json'), 'utf8')).live;
  assert.deepEqual(liveLayouts(readLayouts(join(DEFAULT_SKILL, 'styles'), 'default')), { canvas: live.canvas, layouts: live.layouts, slots: live.slots });
});

test('the JS expression checker accepts and rejects the shared cases', () => {
  for (const source of CASES.valid) assert.doesNotThrow(() => parseExpr(source), source);
  for (const source of CASES.invalid) assert.throws(() => parseExpr(source), undefined, JSON.stringify(source));
});

test('the Python expression checker judges the shared cases the same way', { skip: !PYTHON && 'no python with python-pptx' }, () => {
  const script = `
import importlib.util, json, sys
sys.dont_write_bytecode = True  # no __pycache__ inside the shipped skill folder
spec = importlib.util.spec_from_file_location('bd', sys.argv[1]); m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
cases = json.load(open(sys.argv[2], encoding='utf-8'))
def ok(e):
    try: m.parse_expr(e); return True
    except ValueError: return False
print(json.dumps({'valid': [e for e in cases['valid'] if not ok(e)], 'invalid': [e for e in cases['invalid'] if ok(e)]}))`;
  const run = spawnSync(PYTHON, ['-I', '-c', script, join(DEFAULT_SKILL, 'scripts', 'brand_deck.py'), join(TOOLING, 'fixtures', 'expr-cases.json')], { encoding: 'utf8' });
  assert.equal(run.status, 0, run.stderr);
  assert.deepEqual(JSON.parse(run.stdout), { valid: [], invalid: [] });
});

test('the offline renderer reproduces the golden of default (example and regression fixtures, part by part)', { skip: !PYTHON && 'no python with python-pptx' }, () => {
  const run = spawnSync(PYTHON, ['-I', join(TOOLING, 'regress.py'), 'check'], { encoding: 'utf8' });
  assert.equal(run.status, 0, run.stdout + run.stderr);
  assert.match(run.stdout, /^identical: default-regression, default-regression-2, example/);
});

const SYSTEM_PYTHON = ['python3'].find(candidate => spawnSync(candidate, ['-c', 'import zipfile, ast'], { stdio: 'ignore' }).status === 0);
test('preflight reads a source package without running it, and refuses unsafe archives', { skip: !SYSTEM_PYTHON && 'no python3' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'style-preflight-test-'));
  try {
    // A small package: PPTX via a Codex runtime, an OFL font licence, a system font record, a master with photos.
    const pkg = join(work, 'pkg');
    const files = {
      'SKILL.md': '---\nname: demo\ndescription: 生成可编辑 PPTX\n---\n基于 Kimi Slides 公开模板重建的版式。\n',
      'scripts/build.mjs': "import fs from 'node:fs';\nimport { Presentation } from '@oai/artifact-tool';\nconst x = require('skia-canvas');\nconst out = 'deck.pptx';\nconst f = '~/Library/Fonts/Secret-Brand.ttf';\n",
      'scripts/helper.py': 'import json\nimport requests\n',
      'assets/master.json': JSON.stringify({ width: 960, pages: [{ name: 'p1', objects: [{ type: 'image', slotKind: 'photo' }, { type: 'image' }, { type: 'text', fontSize: 10 }] }, { name: 'p2', objects: [{ type: 'image' }, { type: 'text', fontSize: 30 }] }] }),
      'assets/fonts/OFL-Demo.txt': 'Copyright 2020 The Demo Sans Project Authors\nThis Font Software is licensed under the SIL Open Font License, Version 1.1.\n',
      'assets/fonts/fonts.css': "@font-face{font-family:'Demo Sans';src:url(DemoSans.woff2)}@font-face{font-family:'Ref Times';src:local('Times-Roman'),url(t.woff2)}",
      'assets/fonts/DemoSans.woff2': 'wOF2',
      'assets/fonts/t.woff2': 'wOF2',
      'assets/fonts/times-sources.json': JSON.stringify([{ name: 'Times-Roman', source: '/System/Library/Fonts/Times.ttc', file: 't.woff2' }]),
      'assets/reference/shot-1.png': 'png',
    };
    for (const [rel, body] of Object.entries(files)) { mkdirSync(dirname(join(pkg, rel)), { recursive: true }); writeFileSync(join(pkg, rel), body); }
    const run = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), pkg, '--out', join(work, 'r.md'), '--json', join(work, 'r.json')], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    const report = JSON.parse(readFileSync(join(work, 'r.json'), 'utf8'));
    assert.equal(report.format.format, 'PPTX');
    assert.match(report.runtime.verdict, /Codex 运行时/);
    assert.deepEqual(report.runtime.python_third_party, ['requests']);
    assert.deepEqual(Object.fromEntries(report.fonts.families.map(f => [f.family, f.verdict])), { 'Demo Sans': '可分发', 'Ref Times': '不能分发' });
    assert.deepEqual(report.fonts.referenced_not_bundled.map(f => f.file), ['Secret-Brand.ttf']);
    assert.equal(report.photos.total, 3);
    assert.equal(report.photos.max_page.photos, 2);
    assert.equal(report.min_font.min_pt, 7.5);
    assert.deepEqual(report.design.named_sources, ['Kimi Slides']);
    assert.equal(report.images.reference, 1);
    assert.equal(report.verdict.can_extract, true);
    assert.match(readFileSync(join(work, 'r.md'), 'utf8'), /\| 最小字号 \| \*\*7\.5pt\*\*/);
    // Unsafe archives are refused before anything is unpacked.
    const make = `
import sys, zipfile, stat
kind, path = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(path, 'w') as z:
    if kind == 'dotdot': z.writestr('pkg/../../escape.txt', 'x')
    if kind == 'abs': z.writestr('/abs.txt', 'x')
    if kind == 'link':
        i = zipfile.ZipInfo('pkg/link'); i.external_attr = (stat.S_IFLNK | 0o777) << 16; z.writestr(i, '/etc/passwd')`;
    for (const [kind, message] of [['dotdot', /包含 \.\./], ['abs', /绝对路径/], ['link', /符号链接/]]) {
      const zip = join(work, `${kind}.zip`);
      assert.equal(spawnSync(SYSTEM_PYTHON, ['-I', '-c', make, kind, zip]).status, 0);
      const refused = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), zip, '--out', join(work, `${kind}.md`)], { encoding: 'utf8' });
      assert.equal(refused.status, 2);
      assert.match(refused.stderr, message);
    }
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('html_survey reads an HTML style library without running it: blockers, photos, smallest font, licences, shortlist', { skip: !SYSTEM_PYTHON && 'no python3' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'style-survey-test-'));
  try {
    const pkg = join(work, 'lib');
    const style = (id, html, visuals) => ({
      [`assets/styles/${id}/layouts.json`]: JSON.stringify({ layouts: { '01': { texts: { t1: { style: 'font-size:24px' } }, visuals, regions: { r1: { style: 'position:absolute; left:0' } }, reference: '../../reference/x.png' } } }),
      [`assets/styles/${id}/slides/01.html`]: `<div class="slide-container" style="background:#FFFFFF">${html}<p style="font-size:30px;color:#C81D25">x</p><script>alert(1)</script></div>`,
      [`assets/styles/${id}/fonts.zh-CN.css`]: "@font-face{font-family:'Inter'}@font-face{font-family:'Reference Times'}",
    });
    const files = {
      'assets/catalog.json': JSON.stringify({ styles: [{ id: '001', title: '干净', category: '企业战略' }, { id: '002', title: '难转', category: '生活' }] }),
      'assets/fonts/licenses/inter-OFL.txt': 'OFL',
      'assets/reference/x.png': 'png',
      ...style('001', '<span>ok</span>', { v1: { kind: 'image', required: true } }),
      ...style('002', '<img style="filter: brightness(0.6)"><table></table><svg></svg>', {}),
    };
    for (const [rel, body] of Object.entries(files)) { mkdirSync(dirname(join(pkg, rel)), { recursive: true }); writeFileSync(join(pkg, rel), body); }
    const run = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'html_survey.py'), pkg, '--out', join(work, 's.md'), '--json', join(work, 's.json'), '--pick', '1'], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    const out = JSON.parse(readFileSync(join(work, 's.json'), 'utf8'));
    const [a, b] = out.styles;
    assert.deepEqual(a.photos, { slots: 1, required: 1, per_slide_max: 1 });
    assert.equal(a.min_pt, 9, '24px on the 1920 canvas is 9pt on ours');
    assert.deepEqual(Object.keys(a.blockers), []);
    assert.deepEqual(Object.keys(b.blockers).sort(), ['filter', 'svg', 'table']);
    assert.deepEqual(Object.fromEntries(a.fonts.map(f => [f.family, f.verdict])), { Inter: '可分发', 'Reference Times': '不能分发' });
    assert.ok(a.scores.total > b.scores.total);
    assert.deepEqual(out.shortlist, ['001']);
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('live_preview builds the live mode the model would: one slide per template layout, one per slot', { skip: !PYTHON && 'no python with python-pptx' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'style-live-test-'));
  try {
    const run = spawnSync(PYTHON, ['-I', join(TOOLING, 'live_preview.py'), '--style', 'default', '--out', work, '--no-png'], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    const live = JSON.parse(readFileSync(join(DEFAULT_SKILL, 'styles', 'default', 'layouts.json'), 'utf8')).live;
    const archive = listZip(readFileSync(join(work, 'live.pptx')));
    const slides = archive.entries.filter(entry => /^ppt\/slides\/slide\d+\.xml$/.test(entry.name));
    assert.equal(slides.length, live.layouts.length + live.slots.length);
    const pictures = slides.map(entry => (readZipEntry(archive, entry, 4 * 1024 * 1024).toString('utf8').match(/<p:pic>/g) ?? []).length).reduce((a, b) => a + b, 0);
    assert.equal(pictures, live.slots.filter(slot => slot.kind === 'image').length);
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('gs2layouts turns measured HTML into a layouts draft: halves geometry, one box per column, 9pt floor, CSS substitutes', { skip: !SYSTEM_PYTHON && 'no python3' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'style-gs2-test-'));
  try {
    const st = (over = {}) => ({ 'font-size': '40px', 'font-family': '"Source Serif 4", serif', 'font-weight': '400', 'font-style': 'normal', color: 'rgb(245, 235, 211)', 'letter-spacing': '2px',
      'line-height': '48px', 'text-align': 'start', 'text-transform': 'none', opacity: '1', 'background-color': 'rgba(0, 0, 0, 0)', 'background-image': 'none', 'border-top-width': '0px',
      'border-top-color': 'rgb(0, 0, 0)', filter: 'none', 'mix-blend-mode': 'normal', transform: 'none', 'z-index': '1', ...over });
    const t = (box, over, text) => ({ box, ink: box, lines: 1, tag: 'P', style: st(over), text, slots: [] });
    const measure = { canvas: { width: 1920, height: 1080 }, pages: { '01': { name: '封面', background: 'rgb(22, 20, 19)',
      regions: {
        r1: { kind: 'image', box: [0, 0, 1920, 1080], style: st() },
        r2: { kind: 'shape', box: [0, 0, 1920, 1080], style: st({ 'background-image': 'linear-gradient(rgba(22, 20, 19, 0.5) 0%, rgba(22, 20, 19, 0.9) 100%)' }) },
        r3: { kind: 'shape', box: [96, 120, 120, 2], style: st({ 'background-color': 'rgb(199, 165, 92)', opacity: '0.5' }) },
        r4: { kind: 'textbox', box: [96, 300, 1600, 400], style: st({ transform: 'matrix(1, 0, 0, 1, 0, 0) rotate(3deg)' }) },
      },
      texts: {
        a: t([96, 300, 800, 48], {}, 'First line'), b: t([96, 350, 800, 48], {}, 'Second line'),
        c: t([1000, 300, 400, 20], { 'font-size': '10px', 'font-family': '"JetBrains Mono"', 'text-transform': 'uppercase' }, 'label'),
      },
      visuals: { v1: { box: [0, 0, 1920, 1080], style: st({ filter: 'brightness(0.6)' }) } } } } };
    const source = { layouts: { '01': { regions: { r1: { texts: [], visuals: ['v1'] }, r2: { texts: [], visuals: [] }, r3: { texts: [], visuals: [] }, r4: { texts: ['p.a', 'p.b', 'p.c'], visuals: [] } },
      visuals: { v1: { kind: 'image', required: true, style: 'filter:brightness(0.6)' } } } } };
    measure.pages['01'].texts = Object.fromEntries(Object.entries(measure.pages['01'].texts).map(([k, v]) => [`p.${k}`, v]));
    writeFileSync(join(work, 'm.json'), JSON.stringify(measure));
    writeFileSync(join(work, 'l.json'), JSON.stringify(source));
    const run = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'gs2layouts.py'), join(work, 'm.json'), '--layouts', join(work, 'l.json'), '--ids', 'cover', '--out', join(work, 'd.json'), '--notes', join(work, 'n.md')], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    const draft = JSON.parse(readFileSync(join(work, 'd.json'), 'utf8'));
    const cover = draft.offline.layouts.cover;
    assert.equal(cover.background, '#161413');
    const [img, grad, line, ...texts] = cover.objects;
    assert.deepEqual([img.type, img.mask, cover.fields.photo.needed], ['image', 0.4, true], 'brightness(0.6) becomes a 40% black mask');
    assert.deepEqual(grad.gradient, { angle: 180, stops: [[0, '#161413', 0.5], [1, '#161413', 0.9]] });
    assert.deepEqual([line.type, line.x, line.y, line.w, line.color], ['hline', 48, 60, 60, '#6E5C38'], 'halved, and opacity premixed with the page');
    assert.equal(texts.length, 2, 'two stacked lines are one box, the label beside them another');
    assert.deepEqual([texts[0].size, texts[0].font, texts[0].spacing, texts[0].x], [20, 'serif', 0.75, 48]);
    assert.match(cover.fields[texts[0].name].example, /First line\nSecond line/);
    assert.deepEqual([texts[1].size, texts[1].font], [12, 'mono'], '5px on our canvas is raised to 12px (9pt)');
    const notes = readFileSync(join(work, 'n.md'), 'utf8');
    for (const re of [/brightness\(0\.6\) → 黑色遮罩/, /抬到 12px/, /text-transform/, /旋转\/变形/, /合并了 2 个槽/]) assert.match(notes, re);
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('measure_html refuses a source with symbolic links before opening anything', () => {
  const work = mkdtempSync(join(tmpdir(), 'style-measure-link-'));
  try {
    mkdirSync(join(work, 'src', 'slides'), { recursive: true });
    writeFileSync(join(work, 'outside.css'), 'body{font-size:77px}');
    writeFileSync(join(work, 'src', 'slides', '01.html'), '<link rel="stylesheet" href="link.css">');
    symlinkSync(join(work, 'outside.css'), join(work, 'src', 'slides', 'link.css'));
    writeFileSync(join(work, 'layouts.json'), JSON.stringify({ width: 1920, height: 1080, layouts: { '01': { file: '01.html', texts: {} } } }));
    const run = spawnSync(process.execPath, [join(TOOLING, 'measure_html.mjs'), '--root', join(work, 'src'), '--layouts', join(work, 'layouts.json'), '--slides', 'slides', '--out', join(work, 'm.json')], { encoding: 'utf8' });
    assert.equal(run.status, 1);
    assert.match(run.stderr, /符号链接/);
  } finally { rmSync(work, { recursive: true, force: true }); }
});

// ---------- reference intake: images / overview images / PDF / PPTX (ref_intake.py, measure_image.py, preflight reference mode) ----------
// Every fixture is drawn here: synthetic pages from Pillow, a hand-written PDF, a hand-built sfnt, a registered pack's template.
const POPPLER = ['pdfinfo', 'pdftoppm', 'pdftotext'].every(tool => spawnSync(tool, ['-v'], { stdio: 'ignore' }).error === undefined);
const PILLOW = PYTHON && spawnSync(PYTHON, ['-c', 'import PIL'], { stdio: 'ignore' }).status === 0 ? PYTHON : null;

/** A one-page PDF with real text (Helvetica 24pt at x=100, baseline y=400 from the bottom) on a 960×540pt page. */
function textPdf(font = 'Helvetica') {
  const content = 'BT /F1 24 Tf 100 400 Td (Hello Distill) Tj ET';
  const objects = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 960 540] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`, `<< /Type /Font /Subtype /Type1 /BaseFont /${font} >>`];
  let out = '%PDF-1.4\n';
  const offsets = objects.map((body, i) => { const at = out.length; out += `${i + 1} 0 obj\n${body}\nendobj\n`; return at; });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` + offsets.map(o => `${String(o).padStart(10, '0')} 00000 n \n`).join('');
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/** A minimal TrueType-shaped font: just a name table (family) and an OS/2 table carrying fsType. */
function sfnt(family, fsType) {
  const str = Buffer.from(family, 'utf16le').swap16();
  const name = Buffer.alloc(18 + str.length);
  name.writeUInt16BE(0, 0); name.writeUInt16BE(1, 2); name.writeUInt16BE(18, 4);
  [3, 1, 0x409, 1, str.length, 0].forEach((v, i) => name.writeUInt16BE(v, 6 + i * 2));
  str.copy(name, 18);
  const os2 = Buffer.alloc(78); os2.writeUInt16BE(fsType, 8);
  const tables = [['OS/2', os2], ['name', name]];
  const head = Buffer.alloc(12 + 16 * tables.length);
  head.writeUInt32BE(0x00010000, 0); head.writeUInt16BE(tables.length, 4);
  let offset = head.length;
  tables.forEach(([tag, data], i) => { head.write(tag, 12 + i * 16, 'latin1'); head.writeUInt32BE(offset, 20 + i * 16); head.writeUInt32BE(data.length, 24 + i * 16); offset += data.length; });
  return Buffer.concat([head, ...tables.map(([, data]) => data)]);
}

/** A tiny solid-colour PNG written without Pillow. */
function png(width, height, [r, g, b]) {
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const body = Buffer.concat([Buffer.from(type, 'latin1'), data]); const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(body) >>> 0); return Buffer.concat([len, body, crc]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = 8; ihdr[9] = 2;
  const row = Buffer.concat([Buffer.from([0]), Buffer.from(Array.from({ length: width }, () => [r, g, b]).flat())]);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(Buffer.concat(Array(height).fill(row)))), chunk('IEND', Buffer.alloc(0))]);
}

const iou = (a, b) => {
  const ix = Math.max(0, Math.min(a[2], b[2]) - Math.max(a[0], b[0])), iy = Math.max(0, Math.min(a[3], b[3]) - Math.max(a[1], b[1]));
  const i = ix * iy; return i / ((a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - i);
};

test('ref_intake splits overview images by seams, keeps single slides whole, and asks for boxes when unsure', { skip: !PILLOW && 'no python with Pillow' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'ref-intake-test-'));
  try {
    // Ten synthetic 480×270 pages (own background, a bar, a block of "text" lines), laid out like trial B.
    const make = String.raw`
import json, random, sys
from pathlib import Path
from PIL import Image, ImageDraw
out = Path(sys.argv[1]); random.seed(3)
W, H = 480, 270
def page(i, bg):
    im = Image.new('RGB', (W, H), bg); d = ImageDraw.Draw(im)
    d.rectangle([24, 24, 24 + 40 * (i % 5 + 2), 40], fill=(255, 90, 95))
    for k in range(4):
        d.rectangle([24, 70 + k * 30, 24 + random.randint(200, 420), 84 + k * 30], fill=(200, 200, 205) if sum(bg) < 300 else (40, 40, 50))
    return im
bgs = [(42, 47, 54), (52, 58, 66), (58, 42, 44), (30, 60, 90), (90, 30, 60), (20, 80, 40), (60, 60, 20), (40, 40, 40), (70, 50, 90), (35, 55, 65)]
truth = {}
def grid(name, back, gap, special=None):
    pages = [page(i, bg) for i, bg in enumerate(bgs)]
    if special is not None:
        pages[7] = page(7, special)
    m = Image.new('RGB', (gap + 5 * (W + gap), gap + 2 * (H + gap)), back); boxes = []
    for i, im in enumerate(pages):
        x, y = gap + (i % 5) * (W + gap), gap + (i // 5) * (H + gap); m.paste(im, (x, y)); boxes.append([x, y, x + W, y + H])
    m.save(out / f'{name}.png'); truth[name] = boxes
grid('light', (245, 245, 245), 24)
grid('dark', (17, 17, 17), 24, special=(15, 15, 15))            # one page almost the overview's own background
grid('widegap', (255, 255, 255), 120, special=(252, 250, 247))  # an off-white page on white, below the tolerance
m = Image.new('RGB', (2200, 1300), (245, 245, 245)); boxes = []
for i, (x, y, s) in enumerate([(30, 40, 1.4), (760, 30, 0.8), (1180, 60, 1.0), (1700, 20, 0.9), (40, 520, 0.7), (420, 560, 1.2), (1080, 480, 0.85), (1520, 520, 1.1), (60, 930, 0.9), (560, 1000, 0.6)]):
    w, h = int(W * s), int(H * s); m.paste(page(i, bgs[i]).resize((w, h)), (x, y)); boxes.append([x, y, x + w, y + h])
m.save(out / 'irregular.png'); truth['irregular'] = boxes
# Same background as the overview: no seam can find these pages.
m = Image.new('RGB', (1600, 700), (245, 245, 245)); boxes = []
for i, (x, y) in enumerate([(40, 40), (600, 120), (1080, 300)]):
    m.paste(page(i, (245, 245, 245)), (x, y)); boxes.append([x, y, x + W, y + H])
m.save(out / 'samebg.png'); truth['samebg'] = boxes
# One 16:9 slide with three equal solid cards, and one with a title and body text.
s = Image.new('RGB', (1920, 1080), (42, 47, 54)); d = ImageDraw.Draw(s)
for k in range(3):
    d.rectangle([80 + k * 600, 300, 80 + k * 600 + 560, 720], fill=(60, 66, 76))
s.save(out / 'cards.png')
s = Image.new('RGB', (1920, 1080), (245, 241, 232)); d = ImageDraw.Draw(s)
d.rectangle([120, 160, 1300, 260], fill=(30, 30, 30))
for k in range(5):
    d.rectangle([120, 400 + k * 60, 1600, 430 + k * 60], fill=(90, 90, 90))
s.save(out / 'slide.png')
# A page directory in natural order, mixed formats.
pages = out / 'pages-dir'; pages.mkdir()
page(0, bgs[0]).save(pages / '2.png'); page(1, bgs[1]).save(pages / '10.jpg', quality=95); page(2, bgs[2]).save(pages / '1.png')
(pages / 'notes.json').write_text('{}')
(out / 'truth.json').write_text(json.dumps(truth))`;
    assert.equal(spawnSync(PILLOW, ['-I', '-c', make, work], { encoding: 'utf8' }).status, 0);
    const truth = JSON.parse(readFileSync(join(work, 'truth.json'), 'utf8'));
    const intake = (src, out, ...args) => spawnSync(PILLOW, ['-I', join(TOOLING, 'ref_intake.py'), join(work, src), '--out', join(work, out), ...args], { encoding: 'utf8' });
    const read = out => JSON.parse(readFileSync(join(work, out, 'intake.json'), 'utf8'));
    for (const name of ['light', 'dark', 'widegap', 'irregular']) {
      const run = intake(`${name}.png`, `out-${name}`);
      assert.equal(run.status, 0, `${name}: ${run.stderr}`);
      const result = read(`out-${name}`);
      assert.equal(result.source.kind, 'montage', name);
      assert.equal(result.pages.length, 10, name);
      const boxes = result.pages.map(p => p.from.box);
      for (const expected of truth[name]) assert.ok(Math.max(...boxes.map(b => iou(b, expected))) >= 0.95, `${name}: no page matches ${expected}`);
      assert.equal(readdirSync(join(work, `out-${name}`, 'pages')).length, 10);
      assert.match(result.pages[0].sha256, /^[0-9a-f]{64}$/);
    }
    assert.deepEqual(read('out-light').pages.map(p => p.from.box), truth.light, 'grid pages come row by row, left to right');
    assert.match(read('out-dark').notes.join(), /底色与背景相近，按网格补齐/);
    assert.match(read('out-widegap').notes.join(), /按网格补齐/);
    // Pages on the overview's own background: refuse (exit 3) and say how to go on; --boxes then cuts them.
    const unsure = intake('samebg.png', 'out-samebg');
    assert.equal(unsure.status, 3, unsure.stderr);
    assert.match(unsure.stderr, /--boxes/);
    assert.throws(() => readdirSync(join(work, 'out-samebg')), 'nothing left behind after a refusal');
    writeFileSync(join(work, 'boxes.json'), JSON.stringify(truth.samebg));
    const boxed = intake('samebg.png', 'out-boxed', '--boxes', join(work, 'boxes.json'));
    assert.equal(boxed.status, 0, boxed.stderr);
    assert.deepEqual(read('out-boxed').pages.map(p => [p.width, p.height]), [[480, 270], [480, 270], [480, 270]]);
    assert.equal(read('out-boxed').montage.method, 'boxes');
    writeFileSync(join(work, 'bad-boxes.json'), JSON.stringify([[0, 0, 500, 300], [400, 200, 900, 500]]));
    assert.equal(intake('samebg.png', 'out-bad', '--boxes', join(work, 'bad-boxes.json')).status, 2, 'overlapping boxes are refused');
    // A 16:9 slide stays one page; equal cards on it only raise a warning.
    const cards = intake('cards.png', 'out-cards');
    assert.equal(cards.status, 0, cards.stderr);
    assert.equal(read('out-cards').pages.length, 1);
    assert.match(read('out-cards').notes.join(), /发现 3 个像页面的块；如果这是总览图，加 --montage/);
    assert.doesNotMatch(read('out-cards').notes.join(), /按网格补齐/, 'a page kept whole reports no grid filling');
    assert.equal(intake('cards.png', 'out-cards-m', '--montage').status, 0);
    assert.equal(read('out-cards-m').pages.length, 3);
    const slide = intake('slide.png', 'out-slide');
    assert.equal(slide.status, 0, slide.stderr);
    assert.equal(read('out-slide').pages.length, 1);
    assert.deepEqual(read('out-slide').notes, []);
    // A directory: natural file-name order, other files only listed.
    assert.equal(intake('pages-dir', 'out-dir').status, 0);
    const dir = read('out-dir');
    assert.equal(dir.source.kind, 'images');
    assert.deepEqual(dir.pages.map(p => p.from.file), ['1.png', '2.png', '10.jpg']);
    assert.match(dir.notes.join(), /notes\.json/);
    // An existing non-empty output directory is never overwritten.
    const again = intake('pages-dir', 'out-dir');
    assert.equal(again.status, 2);
    assert.match(again.stderr, /不覆盖/);
    assert.equal(read('out-dir').pages.length, 3, 'the refusal keeps the earlier run intact');
    assert.equal(readdirSync(join(work, 'out-dir', 'pages')).length, 3);
    // Broken input ends in a message and exit 2, never a traceback, and leaves nothing behind (an empty --out the
    // caller made stays).
    writeFileSync(join(work, 'not-a-zip.pptx'), 'hello');
    mkdirSync(join(work, 'mine'));
    const broken = intake('not-a-zip.pptx', 'mine', '--no-render');
    assert.equal(broken.status, 2, broken.stderr);
    assert.doesNotMatch(broken.stderr, /Traceback/);
    assert.deepEqual(readdirSync(join(work, 'mine')), []);
    const badSize = String.raw`
import sys, zipfile
with zipfile.ZipFile(sys.argv[1], 'w') as z:
    z.writestr('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz/></p:presentation>')
with zipfile.ZipFile(sys.argv[2], 'w') as z:
    z.writestr('ppt/presentation.xml', '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>')`;
    assert.equal(spawnSync(PILLOW, ['-I', '-c', badSize, join(work, 'nosize.pptx'), join(work, 'doctype.pptx')]).status, 0);
    for (const [file, message] of [['nosize.pptx', /sldSz 不完整/], ['doctype.pptx', /DOCTYPE/]]) {
      const run = intake(file, `out-${file}`, '--no-render');
      assert.equal(run.status, 2, run.stderr);
      assert.match(run.stderr, message);
      assert.throws(() => readdirSync(join(work, `out-${file}`)));
    }
    for (const boxes of ['[[NaN, 0, 10, 10]]', '[[true, 0, 10, 10]]', '[]']) {
      writeFileSync(join(work, 'odd-boxes.json'), boxes);
      const run = intake('samebg.png', 'out-odd', '--boxes', join(work, 'odd-boxes.json'));
      assert.equal(run.status, 2, `${boxes}: ${run.stderr}`);
      assert.doesNotMatch(run.stderr, /Traceback/);
    }
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('measure_image reads colours, ink boxes, lines and size guesses; PDF text layers come with line boxes', { skip: !PILLOW && 'no python with Pillow' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'measure-image-test-'));
  try {
    const make = String.raw`
import sys
from PIL import Image, ImageDraw
im = Image.new('RGB', (1920, 1080), (42, 47, 54)); d = ImageDraw.Draw(im)
d.rectangle([200, 300, 999, 599], fill=(120, 42, 44))              # a card: canvas 100,150 400×150
d.rectangle([240, 340, 639, 379], fill=(255, 142, 145))           # line 1 inside it: canvas 120,170, 20 px high
d.rectangle([240, 420, 559, 459], fill=(255, 142, 145))           # line 2: canvas top 210 (pitch 40)
d.line([120, 900, 1799, 900], fill=(74, 81, 89), width=2)          # a rule: canvas y 450
im.save(sys.argv[1])`;
    assert.equal(spawnSync(PILLOW, ['-I', '-c', make, join(work, 'p.png')]).status, 0);
    const run = spawnSync(PILLOW, ['-I', join(TOOLING, 'measure_image.py'), join(work, 'p.png'), '--roi', 'card=80,130,440,190', '--roi', 'text=110,160,300,80', '--roi', 'rule=40,440,880,20'], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    const page = JSON.parse(run.stdout).pages[0];
    assert.equal(page.scale_to_canvas, 0.5);
    assert.deepEqual(page.canvas, [960, 540]);
    assert.equal(page.background, '#2A2F36');
    assert.ok(page.palette.some(c => c.hex === '#782A2C'));
    const roi = Object.fromEntries(page.rois.map(r => [r.name, r]));
    assert.deepEqual(roi.card.ink.box_canvas, { x: 100, y: 150, w: 400, h: 150 });
    assert.equal(roi.card.ink.colours[0].hex, '#782A2C');
    assert.equal(roi.text.region_background, '#782A2C', 'the ROI frame sits on the card');
    assert.equal(roi.text.ink.colours[0].hex, '#FF8E91');
    assert.deepEqual(roi.text.ink.lines.map(l => [l.top, l.height]), [[170, 20], [210, 20]]);
    assert.equal(roi.text.ink.pitch_px, 40);
    assert.deepEqual(roi.text.ink.size_estimate.size_px_from_pitch, 33.3);
    assert.equal(roi.rule.ink.colours[0].hex, '#4A5159');
    assert.equal(roi.rule.ink.box_canvas.y, 450);
    assert.equal(spawnSync(PILLOW, ['-I', join(TOOLING, 'measure_image.py'), join(work, 'p.png'), '--roi', 'bad']).status, 2);
    if (POPPLER) {
      writeFileSync(join(work, 'text.pdf'), textPdf());
      const intake = spawnSync(PILLOW, ['-I', join(TOOLING, 'ref_intake.py'), join(work, 'text.pdf'), '--out', join(work, 'in'), '--width', '960'], { encoding: 'utf8' });
      assert.equal(intake.status, 0, intake.stderr);
      const result = JSON.parse(readFileSync(join(work, 'in', 'intake.json'), 'utf8'));
      assert.equal(result.pdf.text_layer, true);
      assert.equal(result.pages[0].width, 960);
      assert.ok(result.pdf.fonts.some(f => /Helvetica/.test(f.name)));
      const measured = spawnSync(PILLOW, ['-I', join(TOOLING, 'measure_image.py'), join(work, 'in', 'intake.json')], { encoding: 'utf8' });
      assert.equal(measured.status, 0, measured.stderr);
      const [line] = JSON.parse(measured.stdout).pages[0].text_lines;
      assert.equal(line.text, 'Hello Distill');
      assert.equal(line.source, 'pdf-text-layer');
      assert.ok(Math.abs(line.box_canvas.x - 100) < 1, JSON.stringify(line));
      assert.ok(line.box_canvas.y > 110 && line.box_canvas.y < 140, JSON.stringify(line));
      assert.ok(line.size_px_est > 18 && line.size_px_est < 30, JSON.stringify(line));
    }
  } finally { rmSync(work, { recursive: true, force: true }); }
});

test('preflight reference mode: pages, sizes, PDF/PPTX font names, fsType, design ownership; packages keep the old path', { skip: !SYSTEM_PYTHON && 'no python3' }, () => {
  const work = mkdtempSync(join(tmpdir(), 'preflight-ref-test-'));
  try {
    const ref = join(work, 'ref');
    mkdirSync(join(ref, 'fonts'), { recursive: true });
    writeFileSync(join(ref, '01.png'), png(32, 18, [42, 47, 54]));
    writeFileSync(join(ref, '02.png'), png(40, 30, [245, 241, 232]));
    writeFileSync(join(ref, 'deck.pdf'), textPdf());
    writeFileSync(join(ref, 'fonts', 'Brand.ttf'), sfnt('Brand Display', 0x0002));
    writeFileSync(join(ref, 'README.md'), '设计稿截图');
    cpSync(join(DEFAULT_SKILL, 'styles', 'default', 'template.pptx'), join(ref, 'template.pptx'));
    const run = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), ref, '--out', join(work, 'r.md'), '--json', join(work, 'r.json')], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    assert.equal(JSON.parse(run.stdout).mode, 'reference');
    const report = JSON.parse(readFileSync(join(work, 'r.json'), 'utf8'));
    const byFile = Object.fromEntries(report.files.map(f => [f.file, f]));
    assert.deepEqual([byFile['01.png'].width, byFile['01.png'].height, byFile['01.png'].aspect], [32, 18, '16:9']);
    assert.equal(byFile['02.png'].aspect, '4:3');
    assert.equal(byFile['deck.pdf'].pages, 1);
    assert.equal(byFile['template.pptx'].pptx.slide_size_px.join('×'), '960×540');
    assert.equal(byFile['template.pptx'].pages, 0);
    assert.equal(report.pages_total, 3);
    const font = report.fonts.files.find(f => f.family === 'Brand Display');
    assert.equal(font.fs_type, 2);
    assert.equal(font.fs_type_label, '限制嵌入');
    assert.ok(report.fonts.named.some(f => f.name === 'Poppins'), 'theme fonts of the PPTX are listed by name');
    if (POPPLER) {
      assert.equal(byFile['deck.pdf'].pdf.text_layer, true);
      assert.ok(report.fonts.named.some(f => /Helvetica/.test(f.name) && f.verdict === '不能分发'));
      // pdffonts pushes columns right for long names: the name must come back whole.
      const long = 'ABCDEF+SourceHanSansSC-Regular-VeryLongFontNameHere-Identity';
      writeFileSync(join(work, 'long.pdf'), textPdf(long));
      const longRun = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), join(work, 'long.pdf'), '--out', join(work, 'l.md'), '--json', join(work, 'l.json')], { encoding: 'utf8' });
      assert.equal(longRun.status, 0, longRun.stderr);
      assert.deepEqual(JSON.parse(readFileSync(join(work, 'l.json'), 'utf8')).files[0].pdf.fonts, [{ name: 'SourceHanSansSC-Regular-VeryLongFontNameHere-Identity', type: 'Type 1', embedded: false }]);
    }
    assert.match(report.verdict.decide[0], /^设计归属待用户确认/);
    assert.ok(report.verdict.blockers.some(b => /template\.pptx 没有页面/.test(b)));
    const md = readFileSync(join(work, 'r.md'), 'utf8');
    assert.match(md, /模式：参考稿模式，目录只含参考文件/);
    assert.match(md, /\| fonts\/Brand\.ttf \| Brand Display \| \*\*待查\*\* \| .* \| 0x0002 限制嵌入 \|/);
    // A single PPTX is a reference now (it used to be unpacked as a source package); macro decks are refused.
    const single = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), join(ref, 'template.pptx'), '--out', join(work, 's.md')], { encoding: 'utf8' });
    assert.equal(JSON.parse(single.stdout).mode_reason, '单个 PPTX');
    cpSync(join(ref, 'template.pptx'), join(work, 'macro.pptm'));
    const macro = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), join(work, 'macro.pptm'), '--out', join(work, 'm.md'), '--json', join(work, 'm.json')], { encoding: 'utf8' });
    assert.match(JSON.parse(readFileSync(join(work, 'm.json'), 'utf8')).files[0].error, /只收 \.pptx/);
    // A source package (scripts, json) stays in package mode; its font table gains fsType.
    const pkg = join(work, 'pkg');
    mkdirSync(join(pkg, 'fonts'), { recursive: true });
    writeFileSync(join(pkg, 'SKILL.md'), '生成 PPTX\n');
    writeFileSync(join(pkg, 'build.py'), 'from pptx import Presentation\n');
    writeFileSync(join(pkg, 'fonts', 'Edit.ttf'), sfnt('Edit Sans', 0x0108));
    writeFileSync(join(pkg, 'shot.png'), png(16, 9, [0, 0, 0]));
    const asPackage = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), pkg, '--out', join(work, 'p.md'), '--json', join(work, 'p.json')], { encoding: 'utf8' });
    assert.equal(asPackage.status, 0, asPackage.stderr);
    assert.equal(JSON.parse(asPackage.stdout).mode, 'package');
    const family = JSON.parse(readFileSync(join(work, 'p.json'), 'utf8')).fonts.families.find(f => f.family === 'Edit Sans');
    assert.deepEqual([family.fs_type, family.fs_type_label], [[0x0108], '可编辑嵌入、不可子集']);
    assert.match(readFileSync(join(work, 'p.md'), 'utf8'), /模式：源包模式，目录含 2 个非参考文件/);
    const forced = spawnSync(SYSTEM_PYTHON, ['-I', join(TOOLING, 'preflight.py'), pkg, '--ref', '--out', join(work, 'f.md')], { encoding: 'utf8' });
    assert.equal(JSON.parse(forced.stdout).mode, 'reference');
  } finally { rmSync(work, { recursive: true, force: true }); }
});
