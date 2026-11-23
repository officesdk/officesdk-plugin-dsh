import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { cpSync, existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listZip, readZipEntry } from '../../core/src/zip.mjs';
import { DEFAULT_STYLE, loadStyle, styleNames } from '../src/host/index.mjs';
import { readLayouts } from '../src/host/style-layouts.mjs';

const SKILL = fileURLToPath(new URL('../skills/shimo-pptx/', import.meta.url));
const STYLES = join(SKILL, 'styles');
const PACKS = readdirSync(STYLES, { withFileTypes: true }).filter(entry => entry.isDirectory()).map(entry => entry.name);

/** A python with python-pptx: SHIMO_PYTHON, else DSH's runtime python (what the model runs the scripts with). */
const PYTHON = [process.env.SHIMO_PYTHON, '/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/python/bin/python3.12', 'python3']
  .filter(Boolean).find(candidate => spawnSync(candidate, ['-c', 'import pptx, lxml'], { stdio: 'ignore' }).status === 0);
const py = (args, options) => spawnSync(PYTHON, args, { encoding: 'utf8', ...options });

function theme(template) {
  const archive = listZip(template);
  const entry = archive.entries.find(item => /^ppt\/theme\/theme\d+\.xml$/.test(item.name));
  return readZipEntry(archive, entry, 4 * 1024 * 1024).toString('utf8');
}
const slot = (xml, name) => new RegExp(`<a:${name}><a:srgbClr val="([0-9A-F]{6})"`).exec(xml)?.[1];
const hex = value => value.replace('#', '').toUpperCase();

test('every style pack has style.json, a live template, a layouts doc and an example', () => {
  assert.ok(PACKS.includes(DEFAULT_STYLE));
  for (const name of PACKS) for (const file of ['style.json', 'layouts.json', 'template.pptx', 'layouts.md', 'example.json', 'thumb.png']) assert.ok(existsSync(join(STYLES, name, file)), `${name}/${file}`);
  assert.deepEqual(styleNames(STYLES), [...PACKS].sort());
});

test('each live template carries its style.json colours and fonts (one source for both modes)', () => {
  for (const name of PACKS) {
    const style = JSON.parse(readFileSync(join(STYLES, name, 'style.json'), 'utf8'));
    const palette = style.palettes[style.default_palette];
    const xml = theme(readFileSync(join(STYLES, name, 'template.pptx')));
    assert.equal(style.name, name);
    for (const [slotName, value] of [['dk1', palette.primary], ['accent1', palette.primary], ['dk2', palette.ink], ['lt1', palette.paper], ['accent2', palette.secondary], ['accent5', palette.lineOnPaper]]) {
      assert.equal(slot(xml, slotName), hex(value), `${name} theme ${slotName}`);
    }
    assert.match(xml, new RegExp(`<a:majorFont><a:latin typeface="${style.fonts.latin}"`));
    assert.match(xml, new RegExp(`<a:ea typeface="${style.fonts.cjk}"`));
  }
});

test('the Host loads a style by name, defaults to default, and names the packs when one is unknown', () => {
  const style = loadStyle(STYLES);
  assert.equal(style.summary.name, 'default');
  assert.equal(style.summary.palette.primary, '#2826C9');
  assert.deepEqual(style.summary.fonts, { cjk: 'Noto Sans SC', latin: 'Poppins' });
  assert.equal(style.summary.sizesPt.page_title, 32);
  assert.deepEqual(style.template, readFileSync(join(STYLES, 'default', 'template.pptx')));
  for (const bad of ['nope', '../default', 'Default', '', 7]) assert.throws(() => loadStyle(STYLES, bad), { code: 'invalid-edit', message: /available: (?:[a-z0-9-]+（[^）]+）, )*default（蓝色品牌）/ });
  // A second pack is picked up by name without any code change.
  const root = mkdtempSync(join(tmpdir(), 'officesdk-styles-'));
  try {
    cpSync(join(STYLES, 'default'), join(root, 'default'), { recursive: true });   // only default: the test counts the packs
    cpSync(join(root, 'default'), join(root, 'second'), { recursive: true });
    const second = JSON.parse(readFileSync(join(root, 'second', 'style.json'), 'utf8'));
    second.name = 'second';
    second.default_palette = 'color-2';
    writeFileSync(join(root, 'second', 'style.json'), JSON.stringify(second));
    assert.deepEqual(styleNames(root), ['default', 'second']);
    assert.equal(loadStyle(root, 'second').summary.palette.primary, second.palettes['color-2'].primary);
    writeFileSync(join(root, 'second', 'style.json'), '{ broken');
    assert.throws(() => loadStyle(root, 'second'), { code: 'invalid-edit', message: /Style second cannot be read/ });
    rmSync(join(root, 'second', 'style.json'));
    assert.deepEqual(styleNames(root), ['default'], 'a pack without its style.json is not offered');
    cpSync(join(root, 'default', 'style.json'), join(root, 'second', 'style.json'));
    rmSync(join(root, 'second', 'template.pptx'));
    assert.deepEqual(styleNames(root), ['default'], 'a pack without its template is not offered');
    cpSync(join(root, 'default', 'template.pptx'), join(root, 'second', 'template.pptx'));
    assert.deepEqual(styleNames(root), ['default', 'second']);
    rmSync(join(root, 'second', 'layouts.json'));
    assert.deepEqual(styleNames(root), ['default'], 'a pack without its layouts.json is not offered (loadStyle needs it)');
  } finally { rmSync(root, { recursive: true, force: true }); }
});

test('offline: brand_deck.py builds the example of every pack by style name and refuses an unknown style', { skip: !PYTHON && 'no python with python-pptx' }, () => {
  const out = mkdtempSync(join(tmpdir(), 'officesdk-offline-'));
  try {
    for (const name of PACKS) {
      const run = py([join(SKILL, 'scripts', 'brand_deck.py'), 'build', '--spec', join(STYLES, name, 'example.json'), '--out', join(out, `${name}.pptx`), '--style', name]);
      assert.equal(run.status, 0, run.stderr);
      const report = JSON.parse(run.stdout);
      assert.equal(report.style, name);
      assert.ok(report.slides > 0);
      // default's example has a native table and chart: the reason to use the offline mode (other packs may have none).
      const parts = listZip(readFileSync(join(out, `${name}.pptx`))).entries.map(entry => entry.name);
      if (name === 'default') assert.ok(parts.some(part => part.startsWith('ppt/charts/chart')), `${name}: chart`);
    }
    const unknown = py([join(SKILL, 'scripts', 'brand_deck.py'), 'build', '--spec', join(STYLES, 'default', 'example.json'), '--out', join(out, 'x.pptx'), '--style', 'nope']);
    assert.equal(unknown.status, 1);
    assert.match(unknown.stderr, /未知样式 nope，可选 (?:[a-z0-9-]+, )*default\b/);
    assert.ok(!existsSync(join(out, 'x.pptx')));
    // A second pack, added as a folder only, builds by its name; a folder without its template is not a pack.
    const skill = mkdtempSync(join(tmpdir(), 'officesdk-offline-skill-'));
    try {
      cpSync(SKILL, skill, { recursive: true });
      cpSync(join(skill, 'styles', 'default'), join(skill, 'styles', 'second'), { recursive: true });
      const second = JSON.parse(readFileSync(join(skill, 'styles', 'second', 'style.json'), 'utf8'));
      second.name = 'second';
      second.default_palette = 'color-2';
      writeFileSync(join(skill, 'styles', 'second', 'style.json'), JSON.stringify(second));
      const built = py([join(skill, 'scripts', 'brand_deck.py'), 'build', '--spec', join(skill, 'styles', 'second', 'example.json'), '--out', join(out, 'second.pptx'), '--style', 'second']);
      assert.equal(built.status, 0, built.stderr);
      assert.equal(JSON.parse(built.stdout).style, 'second');
      rmSync(join(skill, 'styles', 'second', 'template.pptx'));
      const partial = py([join(skill, 'scripts', 'brand_deck.py'), 'build', '--spec', join(skill, 'styles', 'second', 'example.json'), '--out', join(out, 'partial.pptx'), '--style', 'second']);
      assert.equal(partial.status, 1);
      assert.match(partial.stderr, /未知样式 second，可选 (?:[a-z0-9-]+, )*default\b/);
      cpSync(join(skill, 'styles', 'default', 'template.pptx'), join(skill, 'styles', 'second', 'template.pptx'));
      rmSync(join(skill, 'styles', 'second', 'layouts.json'));
      const noLayouts = py([join(skill, 'scripts', 'brand_deck.py'), 'build', '--spec', join(skill, 'styles', 'second', 'example.json'), '--out', join(out, 'nolayouts.pptx'), '--style', 'second']);
      assert.equal(noLayouts.status, 1, 'a folder without layouts.json is not a pack either (the Host offers the same ones)');
      assert.match(noLayouts.stderr, /未知样式 second，可选 (?:[a-z0-9-]+, )*default\b/);
    } finally { rmSync(skill, { recursive: true, force: true }); }
    // Without --style it is default.
    const plain = py([join(SKILL, 'scripts', 'brand_deck.py'), 'build', '--spec', join(STYLES, 'default', 'example.json'), '--out', join(out, 'plain.pptx')]);
    assert.equal(JSON.parse(plain.stdout).style, 'default');
  } finally { rmSync(out, { recursive: true, force: true }); }
});

test('build_template.py rebuilds a pack\'s template from its style.json and refuses an unknown style', { skip: !PYTHON && 'no python with python-pptx' }, () => {
  const root = mkdtempSync(join(tmpdir(), 'officesdk-template-'));
  try {
    cpSync(SKILL, root, { recursive: true });
    const style = JSON.parse(readFileSync(join(root, 'styles', 'default', 'style.json'), 'utf8'));
    style.default_palette = 'color-3';
    writeFileSync(join(root, 'styles', 'default', 'style.json'), JSON.stringify(style));
    const run = py([join(root, 'scripts', 'build_template.py'), '--style', 'default']);
    assert.equal(run.status, 0, run.stderr);
    assert.equal(slot(theme(readFileSync(join(root, 'styles', 'default', 'template.pptx'))), 'accent1'), hex(style.palettes['color-3'].primary));
    const unknown = py([join(root, 'scripts', 'build_template.py'), '--style', 'nope']);
    assert.notEqual(unknown.status, 0);
    assert.match(unknown.stderr, /未知样式 nope，可选 (?:[a-z0-9-]+, )*default\b/);
  } finally { rmSync(root, { recursive: true, force: true }); }
});

test('the default live cover and closing follow the brand cover, and photo-statement is a near full-bleed slot', () => {
  const live = JSON.parse(readFileSync(join(STYLES, 'default', 'layouts.json'), 'utf8')).live;
  const layout = id => live.layouts.find(item => item.id === id);
  // spec §4: pt = inches × 72.
  assert.deepEqual(layout('cover').boxes, { title: [43.2, 187.2, 590.4, 158.4], subtitle: [388.8, 93.6, 309.6, 50.4] });
  assert.deepEqual(layout('closing').boxes, { body: [43.2, 187.2, 590.4, 158.4] }, 'THANKS sits in the cover title box');
  assert.deepEqual(live.slots.find(item => item.id === 'photo-statement'), {
    id: 'photo-statement', layout: 'blank', kind: 'image', box: [0, 50.4, 720, 354.6], art_size: '1920x945', darken: 0.3,
    title_box: [43.2, 72.0, 496.8, 108.0], text_box: [518.4, 201.6, 180.0, 136.8],
    note: '全幅图文：本页 setBackground 主色，顶部主色条，左叠两行标题，右短说明',
  });
  const archive = listZip(readFileSync(join(STYLES, 'default', 'template.pptx')));
  const layouts = archive.entries.filter(item => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(item.name))
    .map(item => readZipEntry(archive, item).toString('utf8'));
  const named = name => layouts.find(xml => xml.includes(`<p:cSld name="${name}"`));
  const style = JSON.parse(readFileSync(join(STYLES, 'default', 'style.json'), 'utf8'));
  const line = hex(style.palettes[style.default_palette].lineOnPrimary);
  const EMU = 12700;
  for (const name of ['封面', '结束页']) {
    const xml = named(name);
    assert.ok(xml, name);
    const rule = /<p:sp>(?:(?!<\/p:sp>).)*name="细横线"(?:(?!<\/p:sp>).)*<\/p:sp>/s.exec(xml)?.[0];
    assert.ok(rule, `${name} has the rule`);
    assert.match(rule, new RegExp(`<a:off x="0" y="${165.6 * EMU}"/><a:ext cx="${720 * EMU}" cy="${EMU}"/>`));
    assert.match(rule, new RegExp(`<a:srgbClr val="${line}"/>`));
    assert.match(xml, /<a:bodyPr[^>]*anchor="b"/, `${name} text sits on its box bottom`);
  }
  assert.doesNotMatch(named('标题和内容'), /细横线/);
});

test('the Host hands the model the live layouts and slots from the pack\'s layouts.json, and refuses a broken one', () => {
  const live = JSON.parse(readFileSync(join(STYLES, 'default', 'layouts.json'), 'utf8')).live;
  const { summary } = loadStyle(STYLES);
  assert.deepEqual(summary.layouts, { canvas: live.canvas, layouts: live.layouts, slots: live.slots });
  assert.deepEqual(summary.layouts.slots.find(slot => slot.id === 'left-image'), { id: 'left-image', layout: 'title-only', kind: 'image', box: [36, 100, 312, 270], art_size: '1248x1080', text_box: [372, 100, 312, 270], note: '左图右文，只给竖构图' });
  // The offline part is read too: its image slots with fixed geometry, conditional ones marked.
  const description = readLayouts(STYLES, 'default');
  assert.deepEqual(description.offline.order[0], 'cover');
  assert.deepEqual(description.offline.layouts['image-statement'].slots, [{ id: 'image', kind: 'image', field: 'image', box: [0, 64, 960, 476], dynamic: false, conditional: false }]);
  assert.equal(description.offline.layouts.statement.slots.find(slot => slot.id === 'image3').conditional, true);
  assert.equal(description.offline.layouts.table.slots[0].dynamic, true, 'the table top depends on the subtitle');
  assert.equal(description.offline.layouts.closing.slots.length, 0);
  const root = mkdtempSync(join(tmpdir(), 'officesdk-layouts-'));
  try {
    cpSync(join(STYLES, 'default'), join(root, 'default'), { recursive: true });   // only default: "available: none" below
    const file = join(root, 'default', 'layouts.json');
    const good = readFileSync(file, 'utf8');
    const broken = [
      [desc => { desc.offline.layouts.cover.objects[1].text = "open('x')"; }, /调用了不允许的函数 open/],
      [desc => { desc.offline.layouts.cover.objects[1].text = 'd.items'; }, /只允许 \.get\(\)/],
      [desc => { desc.offline.layouts.cover.objects.push({ type: 'video' }); }, /未知对象类型 "video"/],
      [desc => { desc.live.slots[0].box = [1, 2, 3]; }, /box \/ text_box/],
      [desc => { desc.live.slots[0].layout = 'nope'; }, /layout "nope" 不在 live\.layouts/],
      [desc => { desc.live.layouts[0].boxes = {}; }, /boxes 要给每个占位符/],
      [desc => { desc.offline.order.pop(); }, /offline\.order/],
      [desc => { desc.version = 2; }, /version 应为 1/],
    ];
    for (const [edit, message] of broken) {
      const desc = JSON.parse(good);
      edit(desc);
      writeFileSync(file, JSON.stringify(desc));
      assert.throws(() => readLayouts(root, 'default'), message);
      assert.throws(() => loadStyle(root, 'default'), { code: 'invalid-edit', message: /Style default has a broken layouts\.json/ });
    }
    rmSync(file);
    assert.throws(() => readLayouts(root, 'default'), /layouts\.json 读不了/);
    // Without its layouts.json the folder is not a pack at all: not offered, and named as unknown.
    assert.throws(() => loadStyle(root, 'default'), { code: 'invalid-edit', message: /Unknown style "default"; available: none/ });
  } finally { rmSync(root, { recursive: true, force: true }); }
});
