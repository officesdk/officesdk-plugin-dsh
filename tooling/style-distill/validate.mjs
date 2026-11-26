#!/usr/bin/env node
/**
 * Check shimo-pptx style packs (style-distill tooling; developer machines and CI).
 *
 *   node tooling/style-distill/validate.mjs [--skill <shimo-pptx dir>] [--styles <packs dir>] [--style <id>] [--python <path>] [--allow-no-python] [--json]
 *
 * Every pack under <skill>/styles (default: packages/dsh-adapter/skills/shimo-pptx) is checked for: files; style.json
 * and its meta; template.pptx against style.json (theme colours and fonts) and against layouts.json (live layouts:
 * placeholders and their boxes); smallest text; text boxes per page; image slots against meta.requires.photos;
 * expressions (whitelist and defined names); the Host reading layouts.json (and layouts.md carrying every live slot);
 * the offline renderer reading it and building example.json (needs a Python with python-pptx); thumb.png.
 * --styles checks packs in another folder (staging, e.g. tooling/style-distill/staging) with the same rules, except that
 * meta.design_source.cleared=false is a note ("not registered yet") there instead of a problem; their ids and aliases
 * may not clash with the registered packs under <skill>/styles either.
 * Exit 1 when any pack fails. Without such a Python the renderer check fails unless --allow-no-python (then skipped).
 * Format: pack-format.md beside this file.
 */
import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir, homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listZip, readZipEntry } from '../../packages/core/src/zip.mjs';
import { exprNames, foldNumber, parseExpr } from '../../packages/dsh-adapter/src/host/style-expr.mjs';
import { checkLayouts, colorExpression, COLOR_KEYS } from '../../packages/dsh-adapter/src/host/style-layouts.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
export const DEFAULT_SKILL = resolve(HERE, '../../packages/dsh-adapter/skills/shimo-pptx');
/** Packs whose design source is not cleared yet wait here (tooling/ is not shipped); same format as styles/<id>/. */
export const STAGING_STYLES = join(HERE, 'staging');
const PACK_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
export const LIMITS = Object.freeze({ minTextPt: 9, minChartPt: 8, minLivePt: 9, maxTextBoxes: 30, thumbMaxWidth: 640, thumbMaxBytes: 150 * 1024, fitMinRatio: 0.62, legendMax: 6 });
const FILES = ['style.json', 'layouts.json', 'template.pptx', 'layouts.md', 'example.json', 'thumb.png'];
const ROLES = ['primary', 'ink', 'onPrimary', 'paper', 'secondary', 'onSecondary', 'inkOnSecondary', 'chartSecondary', 'chartTertiary', 'lineOnPaper', 'lineOnPrimary'];
const FREE_LICENSES = new Set(['OFL-1.1', 'Apache-2.0', 'UFL-1.0', 'MIT']);
const TEXT_OBJECTS = new Set(['text', 'tag', 'highlight', 'arrow']);
const SHAPE_KINDS = ['triangle', 'rtTriangle', 'parallelogram', 'ellipse'];   // brand_deck.py Deck.SHAPES
const PT_PER_PX = 0.75;

const isObject = value => value !== null && typeof value === 'object' && !Array.isArray(value);
const nonEmpty = value => typeof value === 'string' && value.trim().length > 0;
const bilingual = value => isObject(value) && nonEmpty(value.zh) && nonEmpty(value.en);

/** A Python that can run the renderer: explicit, SHIMO_PYTHON, DSH's runtime python, then python3. */
export function findPython(explicit) {
  const works = candidate => spawnSync(candidate, ['-c', 'import pptx, lxml'], { stdio: 'ignore' }).status === 0;
  if (explicit) return works(explicit) ? explicit : null;   // an explicit choice is not silently replaced
  const candidates = [process.env.SHIMO_PYTHON,
    '/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/python/bin/python3.12',
    join(homedir(), '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3'), 'python3'].filter(Boolean);
  return candidates.find(works) ?? null;
}

/** Chart text sizes the renderer uses unless a chart object overrides them (read from brand_deck.py: one source). */
function chartFonts(skill) {
  const source = readFileSync(join(skill, 'scripts', 'brand_deck.py'), 'utf8');
  const m = /^CHART_FONTS = (\{[^}]*\})/m.exec(source);
  if (!m) throw new Error('brand_deck.py 里找不到 CHART_FONTS');
  return JSON.parse(m[1].replace(/'/g, '"'));
}

// ---------- template.pptx ----------
function templateFacts(data) {
  const archive = listZip(data);
  const themeEntry = archive.entries.find(item => /^ppt\/theme\/theme\d+\.xml$/.test(item.name));
  const theme = readZipEntry(archive, themeEntry, 4 * 1024 * 1024).toString('utf8');
  const slot = name => new RegExp(`<a:${name}><a:srgbClr val="([0-9A-Fa-f]{6})"`).exec(theme)?.[1]?.toUpperCase();
  const layouts = archive.entries.filter(item => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(item.name)).map(item => {
    const xml = readZipEntry(archive, item, 4 * 1024 * 1024).toString('utf8');
    const name = /<p:cSld name="([^"]*)"/.exec(xml)?.[1] ?? '';
    const background = /<p:bg>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(xml)?.[1]?.toUpperCase() ?? null;
    const boxes = {}, colors = {};
    for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []) {
      const ph = /<p:ph(?:\s+type="([^"]+)")?[^>]*\/>/.exec(sp);
      if (!ph) continue;
      const kind = { ctrTitle: 'title', title: 'title', subTitle: 'subtitle', body: 'body', obj: 'body' }[ph[1] ?? 'body'];
      if (!kind) continue;
      const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(sp), ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(sp);
      boxes[kind] = off && ext ? [off[1], off[2], ext[1], ext[2]].map(v => Number(v) / 12700) : null;
      colors[kind] = /<a:defRPr[^>]*>\s*<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(sp)?.[1]?.toUpperCase() ?? null;
    }
    return { name, boxes, colors, background };
  });
  return {
    colors: Object.fromEntries(['dk1', 'dk2', 'lt1', 'accent1', 'accent2', 'accent5'].map(name => [name, slot(name)])),
    majorLatin: /<a:majorFont><a:latin typeface="([^"]*)"/.exec(theme)?.[1],
    ea: /<a:ea typeface="([^"]*)"/.exec(theme)?.[1],
    layouts,
  };
}

// ---------- layouts.json walks ----------
/** Lower bound of a font size (px): numbers, fit(...) at its floor, let names resolved, arithmetic folded. */
function minSize(value, lets) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return undefined;
  const walk = node => {
    if (node.type === 'call' && node.fn === 'fit') {
      const size = walk(node.args[1]);
      const ratio = node.args[3] ? walk(node.args[3]) : LIMITS.fitMinRatio;
      return size === undefined || ratio === undefined ? undefined : size * ratio;
    }
    if (node.type === 'name') return lets.has(node.id) ? minSize(lets.get(node.id), lets) : undefined;
    if (node.type === 'const') return typeof node.value === 'number' ? node.value : undefined;
    if (node.type === 'unary' && node.op === '-') return undefined;
    if (node.type === 'bin') {
      // Lower bounds compose only for + and *; for -, /, //, % the right side would need its upper bound, so a
      // non-constant right side there is "cannot tell" (an error), never a guessed floor.
      if (['-', '/', '//', '%'].includes(node.op) && foldNumber(node.right) === undefined) return undefined;
      const a = walk(node.left), b = walk(node.right);
      if (a === undefined || b === undefined) return undefined;
      return foldNumber({ type: 'bin', op: node.op, left: { type: 'const', value: a }, right: { type: 'const', value: b } });
    }
    if (node.type === 'if') {
      const a = walk(node.body), b = walk(node.orelse);
      return a === undefined || b === undefined ? undefined : Math.min(a, b);
    }
    if (node.type === 'call' && (node.fn === 'max' || node.fn === 'min') && node.args.length) {
      const bounds = node.args.map(walk);
      if (bounds.some(v => v === undefined)) return undefined;
      return node.fn === 'max' ? Math.max(...bounds) : Math.min(...bounds);
    }
    return undefined;
  };
  return walk(parseExpr(value));
}

/** Keys every slide may carry besides its layout's fields (read by the renderer itself, not by layouts). */
const SLIDE_KEYS = new Set(['layout', 'notes', 'sources']);
/** Field names an expression reads from the slide: d['x'] and d.get('x'). */
function fieldRefs(tree, out = new Set()) {
  if (!tree || typeof tree !== 'object') return out;
  const isD = node => node?.type === 'name' && node.id === 'd';
  if (tree.type === 'index' && isD(tree.value) && tree.index?.type === 'const' && typeof tree.index.value === 'string') out.add(tree.index.value);
  if (tree.type === 'get' && isD(tree.object) && tree.args[0]?.type === 'const' && typeof tree.args[0].value === 'string') out.add(tree.args[0].value);
  for (const value of Object.values(tree)) {
    if (Array.isArray(value)) value.forEach(item => fieldRefs(item, out));
    else if (value && typeof value === 'object') fieldRefs(value, out);
  }
  return out;
}

function walkLayout(layout, visit) {
  const base = new Map(Object.entries(layout.let ?? {}));
  const names = new Set(['d', 'deck', 'style', 'P', ...Object.keys(layout.vars ?? {}), ...base.keys()]);
  const each = (objects, lets, scope, multiplier, path) => {
    for (const [index, object] of objects.entries()) {
      const at = `${path}[${index}]`;
      let here = new Set(scope), letsHere = new Map(lets);
      if (object.type === 'repeat') for (const name of ['item', 'i', 'n']) here.add(name);
      for (const [name, source] of Object.entries(object.let ?? {})) { here.add(name); letsHere.set(name, source); }
      visit(object, { names: here, lets: letsHere, multiplier, at, own: Object.entries(object.let ?? {}) });
      if (object.objects) each(object.objects, letsHere, here, object.type === 'repeat' ? multiplier * (object.max ?? 1) : multiplier, `${at}.objects`);
    }
  };
  each(layout.objects ?? [], base, names, 1, 'objects');
  return { names, base };
}

function checkPackImages(dir, meta, stats, err) {
  // thumb.png and the optional previews: PNG, 16:9, not too wide or heavy.
  const image = name => {
    const png = readFileSync(join(dir, name));
    const sig = png.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
    let size;
    if (!sig) err(`${name} 不是 PNG`);
    else {
      const width = png.readUInt32BE(16), height = png.readUInt32BE(20);
      if (width > LIMITS.thumbMaxWidth) err(`${name} 宽 ${width}px，超过 ${LIMITS.thumbMaxWidth}`);
      if (Math.abs(width / height - 16 / 9) > (16 / 9) * 0.02) err(`${name} 是 ${width}×${height}，不是 16:9`);
      size = `${width}×${height}, ${png.length} B`;
    }
    if (png.length > LIMITS.thumbMaxBytes) err(`${name} 有 ${png.length} 字节，超过 ${LIMITS.thumbMaxBytes}`);
    return size;
  };
  if (isObject(meta) && typeof meta.thumb === 'string' && existsSync(join(dir, meta.thumb))) stats.thumb = image(meta.thumb);
  else if (isObject(meta)) err(`meta.thumb 指向的文件 ${meta.thumb} 不存在`);
  if (isObject(meta) && Array.isArray(meta.previews)) {
    for (const name of meta.previews) if (typeof name === 'string' && !name.includes('/') && !existsSync(join(dir, name))) err(`meta.previews 里的 ${name} 不存在`);
    stats.previews = meta.previews.filter(name => typeof name === 'string' && !name.includes('/') && existsSync(join(dir, name))).map(image).length;
  }

}

function checkFixedPages(dir, style, example, err) {
  const { slideCount, textSlots } = style.fixed_deck;
  if (textSlots !== 'text-slots.json' || !Number.isInteger(slideCount) || slideCount < 1) {
    err('fixed_deck requires a positive slideCount and text-slots.json');
    return;
  }
  try {
    const slots = JSON.parse(readFileSync(join(dir, textSlots), 'utf8'));
    const archive = listZip(readFileSync(join(dir, 'template.pptx')));
    const slides = archive.entries.filter(item => /^ppt\/slides\/slide\d+\.xml$/.test(item.name));
    if (slideCount !== slides.length || slots.length !== slideCount || example?.slides?.length !== slideCount) {
      err('The original deck, text slot list and example must have the declared number of pages');
      return;
    }
    for (let index = 0; index < slideCount; index++) {
      const part = slides.find(item => item.name === `ppt/slides/slide${index + 1}.xml`);
      const xml = readZipEntry(archive, part, 4 * 1024 * 1024).toString('utf8');
      const shapes = new Set();
      for (const [shape] of xml.matchAll(/<(?:\w+:)?sp(?:\s[^>]*|)>[\s\S]*?<\/(?:\w+:)?sp>/g)) {
        if (/<(?:\w+:)?txBody[\s>]/.test(shape)) shapes.add(/<(?:\w+:)?cNvPr\b[^>]*?\bname="([^"]+)"/.exec(shape)?.[1]);
      }
      const page = slots[index];
      if (!Array.isArray(page) || new Set(page).size !== page.length || page.some(key => !shapes.has(key))) {
        err(`Slide ${index + 1} has a duplicate or missing editable text slot`);
      }
    }
  } catch (error) {
    err(`Cannot validate fixed pages: ${error.message}`);
  }
}

function checkPack(stylesRoot, staging, id, problems, stats, notes, fonts) {
  const dir = join(stylesRoot, id);
  const err = message => problems.push(message);
  for (const file of FILES) if (!existsSync(join(dir, file))) err(`缺少文件 ${file}`);
  if (problems.length) return null;
  let style, desc, example;
  try { style = JSON.parse(readFileSync(join(dir, 'style.json'), 'utf8')); } catch (error) { err(`style.json 读不了：${error.message}`); return null; }
  try { example = JSON.parse(readFileSync(join(dir, 'example.json'), 'utf8')); } catch (error) { err(`example.json 读不了：${error.message}`); }

  // style.json
  if (style.name !== id) err(`style.json name=${JSON.stringify(style.name)}，应等于目录名 ${id}`);
  if (!nonEmpty(style.title)) err('style.json 缺少 title');
  const c = style.canvas ?? {};
  if (c.width_px !== 960 || c.height_px !== 540 || c.emu_per_px !== 9525 || c.pt_per_px !== PT_PER_PX) err('style.json canvas 应为 960×540 px、emu_per_px 9525、pt_per_px 0.75');
  const palette = style.palettes?.[style.default_palette];
  if (!palette) err(`default_palette ${JSON.stringify(style.default_palette)} 不在 palettes 里`);
  for (const [pid, p] of Object.entries(style.palettes ?? {})) {
    const missing = ROLES.filter(role => !/^#[0-9A-Fa-f]{6}$/.test(p?.[role] ?? ''));
    if (missing.length) err(`配色 ${pid} 缺少或写错颜色：${missing.join(', ')}`);
    if (!nonEmpty(p?.name)) err(`配色 ${pid} 缺少 name`);
  }
  for (const key of ['cjk', 'latin', 'latin_light']) if (!nonEmpty(style.fonts?.[key])) err(`fonts.${key} 缺失`);
  if (style.fonts?.roles !== undefined) {
    if (!isObject(style.fonts.roles)) err('fonts.roles 应为 {角色: {latin, cjk, latin_light?}}');
    else for (const [name, role] of Object.entries(style.fonts.roles)) if (!nonEmpty(role?.latin) || !nonEmpty(role?.cjk)) err(`fonts.roles.${name} 要有 latin 和 cjk`);
  }
  const tcolors = style.template?.colors;
  if (tcolors !== undefined) {
    const known = ['cover_bg', 'closing_bg', 'content_bg', 'cover_text', 'title', 'body'];
    if (!isObject(tcolors)) err('template.colors 应为 {用处: 配色角色名}');
    else for (const [key, role] of Object.entries(tcolors)) {
      if (!known.includes(key)) err(`template.colors.${key} 不认识，可选 ${known.join(', ')}`);
      for (const [pid, p] of Object.entries(style.palettes ?? {})) if (!/^#[0-9A-Fa-f]{6}$/.test(p?.[role] ?? '')) err(`template.colors.${key}=${role} 在配色 ${pid} 里没有`);
    }
  }
  if (style.template?.dark !== undefined && typeof style.template.dark !== 'boolean') err('template.dark 应为 true/false');
  if (style.template?.cover !== undefined && style.template.cover !== 'brand') err('template.cover 只能是 "brand"（或不写）');
  const sizes = style.template?.sizes_pt;
  if (style.fixed_deck && sizes === undefined) { /* Native slide XML owns the type sizes. */ }
  else if (!isObject(sizes) || !Object.values(sizes).every(v => typeof v === 'number')) err('template.sizes_pt 应为数字表');
  else for (const [name, pt] of Object.entries(sizes)) if (pt < LIMITS.minLivePt) err(`实时字号 template.sizes_pt.${name}=${pt}pt，小于 ${LIMITS.minLivePt}pt`);

  // meta
  const meta = style.meta;
  if (!isObject(meta)) err('style.json 缺少 meta');
  else {
    if (meta.id !== id) err(`meta.id=${JSON.stringify(meta.id)}，应等于 ${id}`);
    if (!bilingual(meta.name)) err('meta.name 要有 zh 和 en');
    if (!bilingual(meta.summary)) err('meta.summary 要有 zh 和 en');
    if (!Array.isArray(meta.aliases) || !meta.aliases.every(nonEmpty) || new Set(meta.aliases).size !== meta.aliases.length) err('meta.aliases 要是不重复的字符串数组');
    for (const lang of ['zh', 'en']) {
      const list = meta.scenarios?.[lang];
      if (!Array.isArray(list) || list.length < 1 || list.length > 6 || !list.every(nonEmpty)) err(`meta.scenarios.${lang} 要有 1–6 条`);
    }
    const photos = meta.requires?.photos;
    if (!isObject(photos) || !['slots', 'required', 'per_slide_max'].every(key => Number.isInteger(photos[key]) && photos[key] >= 0)) err('meta.requires.photos 要有 slots、required、per_slide_max（非负整数）');
    if (!bilingual(meta.requires?.note)) err('meta.requires.note 要有 zh 和 en');
    if (typeof meta.thumb !== 'string' || !/^[^/\\]+\.png$/.test(meta.thumb) || meta.thumb.includes('..')) err('meta.thumb 应为包内的 .png 文件名');
    // Optional: 2–4 page images the plugin's template gallery shows under the thumbnail.
    if (meta.previews !== undefined && (!Array.isArray(meta.previews) || meta.previews.length < 2 || meta.previews.length > 4 || new Set(meta.previews).size !== meta.previews.length
      || !meta.previews.every(name => typeof name === 'string' && /^[\w-][\w.-]*\.png$/.test(name)))) err('meta.previews 可不写；写了要是 2–4 个不重复的包内 .png 文件名');
    if (!Array.isArray(meta.fonts) || !meta.fonts.length) err('meta.fonts 至少列一个字体');
    else for (const [index, font] of meta.fonts.entries()) {
      if (!['family', 'role', 'license', 'source'].every(key => nonEmpty(font?.[key])) || typeof font?.bundled !== 'boolean') err(`meta.fonts[${index}] 要有 family、role、license、source、bundled`);
      else if (font.bundled && !FREE_LICENSES.has(font.license)) err(`meta.fonts[${index}] ${font.family} 的授权 ${font.license} 不能随包分发（bundled 应为 false）`);
    }
    const roleFamilies = Object.values(isObject(style.fonts?.roles) ? style.fonts.roles : {}).flatMap(role => [role?.latin, role?.cjk, role?.latin_light]);
    for (const family of [style.fonts?.cjk, style.fonts?.latin, style.fonts?.latin_light, ...roleFamilies]) {
      if (family && Array.isArray(meta.fonts) && !meta.fonts.some(font => font?.family === family)) err(`meta.fonts 没有列出 style.json 用到的字体 ${family}`);
    }
    const ds = meta.design_source;
    if (!isObject(ds) || !nonEmpty(ds.from) || !nonEmpty(ds.note) || typeof ds.cleared !== 'boolean') err('meta.design_source 要有 from、note、cleared');
    else if (!ds.cleared && staging) notes.push('待登记：meta.design_source.cleared 为 false，设计来源和授权还没经用户或法务确认（暂存，不能挪进 styles/）');
    else if (!ds.cleared) err('meta.design_source.cleared 为 false：设计来源和授权还没经用户或法务确认，不能进 styles/');
  }

  // layouts.json, read the way the Host reads it
  try { desc = checkLayouts(JSON.parse(readFileSync(join(dir, 'layouts.json'), 'utf8')), `${id}/layouts.json`); }
  catch (error) { err(`Host 读不了 layouts.json：${error.message}`); return { style, example }; }

  checkPackImages(dir, meta, stats, err);
  if (style.fixed_deck) {
    checkFixedPages(dir, style, example, err);
    return { style, example, desc };
  }

  // live part ↔ template.pptx ↔ layouts.md
  let facts;
  try { facts = templateFacts(readFileSync(join(dir, 'template.pptx'))); } catch (error) { err(`template.pptx 读不了：${error.message}`); }
  if (facts && palette) {
    const hex = value => value?.replace('#', '').toUpperCase();
    for (const [slot, value] of [['dk1', palette.primary], ['accent1', palette.primary], ['dk2', palette.ink], ['lt1', palette.paper], ['accent2', palette.secondary], ['accent5', palette.lineOnPaper]]) {
      if (facts.colors[slot] !== hex(value)) err(`模板主题 ${slot}=${facts.colors[slot]}，style.json 是 ${hex(value)}（重跑 build_template.py）`);
    }
    if (facts.majorLatin !== style.fonts?.latin) err(`模板主题西文字体 ${facts.majorLatin}，style.json 是 ${style.fonts?.latin}`);
    if (facts.ea !== style.fonts?.cjk) err(`模板主题中文字体 ${facts.ea}，style.json 是 ${style.fonts?.cjk}`);
    for (const layout of desc.live.layouts) {
      const found = facts.layouts.find(item => item.name === layout.name);
      if (!found) { err(`模板里没有版式「${layout.name}」（live.layouts.${layout.id}）`); continue; }
      const have = Object.keys(found.boxes).sort().join(), want = [...layout.placeholders].sort().join();
      if (have !== want) { err(`模板版式「${layout.name}」的占位符是 [${have}]，layouts.json 写的是 [${want}]`); continue; }
      for (const [ph, box] of Object.entries(layout.boxes)) {
        const actual = found.boxes[ph];
        if (!actual || actual.some((v, k) => Math.abs(v - box[k]) > 0.5)) err(`模板版式「${layout.name}」${ph} 的位置 ${JSON.stringify(actual?.map(v => +v.toFixed(2)))}，layouts.json 写的是 ${JSON.stringify(box)}`);
      }
    }
    // Layout backgrounds as build_template.py sets them: template.colors, else the default rules (cover and closing on
    // primary; content layouts keep the master background unless content_bg is given).
    const tc = style.template?.colors ?? {};
    const want = { 封面: tc.cover_bg ?? 'primary', 结束页: tc.closing_bg ?? tc.cover_bg ?? 'primary' };
    if (tc.content_bg) for (const name of ['标题和内容', '仅标题', '空白']) want[name] = tc.content_bg;
    for (const [name, role] of Object.entries(want)) {
      const found = facts.layouts.find(item => item.name === name);
      if (found && palette[role] && found.background !== hex(palette[role])) err(`模板版式「${name}」背景 ${found.background ?? '无'}，按 template.colors 应为 ${role}（${hex(palette[role])}；重跑 build_template.py）`);
    }
    // Placeholder text colours as build_template.py sets them: cover/closing text cover_text (default onPrimary), page
    // titles title (default primary), body text body when given.
    const text = [['封面', 'title', tc.cover_text ?? 'onPrimary'], ['封面', 'subtitle', tc.cover_text ?? 'onPrimary'], ['结束页', 'body', tc.cover_text ?? 'onPrimary'],
      ['标题和内容', 'title', tc.title ?? 'primary'], ['仅标题', 'title', tc.title ?? 'primary'], ...(tc.body ? [['标题和内容', 'body', tc.body]] : [])];
    for (const [name, ph, role] of text) {
      const found = facts.layouts.find(item => item.name === name);
      if (found && ph in found.boxes && palette[role] && found.colors[ph] !== hex(palette[role])) err(`模板版式「${name}」${ph} 文字颜色 ${found.colors[ph] ?? '无'}，按 template.colors 应为 ${role}（${hex(palette[role])}；重跑 build_template.py）`);
    }
    if (facts.layouts.length !== desc.live.layouts.length) err(`模板有 ${facts.layouts.length} 个版式，layouts.json live.layouts 写了 ${desc.live.layouts.length} 个`);
  }
  const doc = readFileSync(join(dir, 'layouts.md'), 'utf8');
  for (const slot of desc.live.slots) if (!doc.includes(slot.box.join(', '))) err(`layouts.md 没有写实时槽位 ${slot.id} 的位置 ${slot.box.join(', ')}`);

  // offline part: names, fonts, text boxes, image slots
  const perLayout = {};
  let minText = Infinity, minChart = Infinity;
  for (const [lid, entry] of Object.entries(desc.offline.layouts)) {
    const layout = entry.raw;
    let textBoxes = 0, images = 0, needed = 0;
    const { names: top } = walkLayout(layout, () => {});
    const checkColor = (value, names, where) => {
      if (value === null || value === undefined || /^#[0-9A-Fa-f]{6}$/.test(value)) return;
      if (typeof value !== 'string') { err(`${where} 的颜色必须是字符串`); return; }
      const source = colorExpression(value);
      if (source) { checkExpr(source, names, where); return; }
      const roles = value.split('|');
      // Like the renderer: the first role the current palette has; a lone name that is no role is read as an expression.
      if (roles.length === 1 && names.has(value)) return;
      for (const [pid, colors] of Object.entries(style.palettes ?? {})) if (!roles.some(role => role in colors)) err(`${where} 的颜色 ${value} 在配色 ${pid} 里没有`);
    };
    const checkExpr = (source, names, where) => {
      const tree = parseExpr(source);
      for (const used of exprNames(tree)) if (!names.has(used)) err(`${where} 用到未定义的名字 ${used}`);
      for (const field of fieldRefs(tree)) if (!(field in (layout.fields ?? {})) && !SLIDE_KEYS.has(field)) err(`${where} 读了字段 d['${field}']，但 fields 里没有声明它`);
    };
    checkColor(layout.background ?? 'paper', top, `${lid}.background`);
    for (const [name, source] of Object.entries(layout.let ?? {})) checkExpr(source, top, `${lid}.let.${name}`);
    walkLayout(layout, (object, { names, lets, multiplier, at }) => {
      for (const [key, value] of Object.entries(object)) {
        if (COLOR_KEYS.includes(key)) { checkColor(value, names, `${lid}.${at}.${key}`); continue; }
        if (key === 'gradient') {
          if (!isObject(value) || !Array.isArray(value.stops) || value.stops.length < 2 || !value.stops.every(stop => Array.isArray(stop) && stop.length === 3 && typeof stop[0] === 'number' && stop[0] >= 0 && stop[0] <= 1 && (stop[2] === null || (typeof stop[2] === 'number' && stop[2] >= 0 && stop[2] <= 1)))) err(`${lid}.${at}.gradient 要写 {angle, stops: [[位置 0–1, 颜色, 不透明度 0–1 或 null], …]}，至少 2 个`);
          else value.stops.forEach((stop, k) => checkColor(stop[1], names, `${lid}.${at}.gradient.stops[${k}]`));
          if (value?.angle !== undefined && typeof value.angle !== 'number') err(`${lid}.${at}.gradient.angle 要是数字（度）`);
          continue;
        }
        if (key === 'para_colors' && object.type === 'text') {
          if (!Array.isArray(value) || !value.length) err(`${lid}.${at}.para_colors 要是颜色数组（第 N 段用第 N 个，其余用 color）`);
          else value.forEach((c, k) => checkColor(c, names, `${lid}.${at}.para_colors[${k}]`));
          continue;
        }
        if (key === 'font' && object.type === 'text') {
          if (typeof value !== 'string' || !isObject(style.fonts?.roles?.[value])) err(`${lid}.${at}.font=${JSON.stringify(value)} 不在 style.json fonts.roles 里`);
          continue;
        }
        if (object.type === 'shape' && key === 'kind') {
          if (!SHAPE_KINDS.includes(value)) err(`${lid}.${at}.kind=${JSON.stringify(value)}，可选 ${SHAPE_KINDS.join(', ')}`);
          continue;
        }
        if (typeof value !== 'string' || ['type', 'name', 'slot', 'align', 'anchor', 'label', 'extends', 'kind', 'style_id'].includes(key)) continue;
        checkExpr(value, names, `${lid}.${at}.${key}`);
      }
      for (const [name, source] of Object.entries(object.let ?? {})) checkExpr(source, names, `${lid}.${at}.let.${name}`);
      if (TEXT_OBJECTS.has(object.type)) {
        textBoxes += multiplier;
        const base = minSize(object.size, lets);
        const px = base !== undefined && object.type === 'highlight' ? base * LIMITS.fitMinRatio : base;
        if (px === undefined) err(`${lid}.${at} 的字号 ${JSON.stringify(object.size)} 推算不出下限：写数字、fit(...)、min/max 或它们的算式`);
        else {
          minText = Math.min(minText, px * PT_PER_PX);
          if (px * PT_PER_PX < LIMITS.minTextPt) err(`${lid}.${at} 字号最小 ${(px * PT_PER_PX).toFixed(2)}pt，小于 ${LIMITS.minTextPt}pt`);
        }
      }
      if (object.type === 'header') { textBoxes += multiplier; minText = Math.min(minText, 12 * PT_PER_PX); }
      if (object.type === 'legend') {
        textBoxes += multiplier * LIMITS.legendMax;
        const pt = (object.size ?? 13) * PT_PER_PX;
        minText = Math.min(minText, pt);
        if (pt < LIMITS.minTextPt) err(`${lid}.${at} 图例字号 ${pt}pt，小于 ${LIMITS.minTextPt}pt`);
      }
      if (object.type === 'table') {
        const pt = Math.min(object.font ?? 15, object.font_small ?? 13) * PT_PER_PX;
        minText = Math.min(minText, pt);
        if (pt < LIMITS.minTextPt) err(`${lid}.${at} 表格字号 ${pt}pt，小于 ${LIMITS.minTextPt}pt`);
      }
      if (object.type === 'chart') {
        const merged = { ...fonts, ...(object.fonts ?? {}) };
        const sizesPt = { ...Object.fromEntries(Object.entries(merged).map(([k, v]) => [k, v * PT_PER_PX])), doughnut_labels: merged.doughnut_labels * PT_PER_PX * 1.3 };
        for (const [k, pt] of Object.entries(sizesPt)) {
          minChart = Math.min(minChart, pt);
          if (pt < LIMITS.minChartPt) err(`${lid}.${at} 图表 ${k} 字号 ${pt.toFixed(2)}pt，小于 ${LIMITS.minChartPt}pt`);
        }
      }
      if (object.type === 'image') {
        images += 1;
        if (layout.fields?.[object.slot]?.needed) needed += 1;
        if (!layout.fields?.[object.slot] || layout.fields[object.slot].type !== 'image') err(`${lid}.${at} 的图片槽 ${object.slot} 没有在 fields 里声明为 image`);
      }
    });
    if (textBoxes > LIMITS.maxTextBoxes) err(`${lid} 最多 ${textBoxes} 个文本框，超过 ${LIMITS.maxTextBoxes}`);
    perLayout[lid] = { textBoxes, images, needed };
  }
  const photos = {
    slots: Object.values(perLayout).reduce((sum, item) => sum + item.images, 0),
    required: Object.values(perLayout).reduce((sum, item) => sum + item.needed, 0),
    per_slide_max: Math.max(0, ...Object.values(perLayout).map(item => item.images)),
  };
  const declared = meta?.requires?.photos;
  if (isObject(declared)) for (const key of Object.keys(photos)) if (declared[key] !== photos[key]) err(`meta.requires.photos.${key}=${declared[key]}，按 layouts.json 算是 ${photos[key]}`);
  if (example && Array.isArray(example.slides)) for (const [index, slide] of example.slides.entries()) if (!desc.offline.layouts[slide?.layout]) err(`example.json 第 ${index + 1} 页的 layout ${JSON.stringify(slide?.layout)} 不在 layouts.json 里`);

  Object.assign(stats, { layouts: Object.keys(desc.offline.layouts).length, textBoxes: Object.fromEntries(Object.entries(perLayout).map(([k, v]) => [k, v.textBoxes])), photos,
    minTextPt: Number.isFinite(minText) ? +minText.toFixed(2) : null, minChartPt: Number.isFinite(minChart) ? +minChart.toFixed(2) : null });
  return { style, example, desc };
}

function checkRenderer(skill, stylesRoot, id, python, example, desc, problems) {
  const script = join(skill, 'scripts', 'brand_deck.py');
  const root = ['--styles-dir', stylesRoot];
  const listed = spawnSync(python, ['-I', script, 'layouts', '--style', id, ...root], { encoding: 'utf8' });
  if (listed.status !== 0) { problems.push(`渲染器读不了 ${id}：${listed.stderr.trim()}`); return; }
  const names = Object.keys(JSON.parse(listed.stdout));
  if (desc && names.join() !== desc.offline.order.join()) problems.push(`渲染器列出的版式 [${names}] 与 layouts.json order [${desc.offline.order}] 不一致`);
  const out = mkdtempSync(join(tmpdir(), 'style-validate-'));
  try {
    const built = spawnSync(python, ['-I', script, 'build', '--spec', join(stylesRoot, id, 'example.json'), '--out', join(out, 'example.pptx'), '--style', id, ...root], { encoding: 'utf8' });
    if (built.status !== 0) { problems.push(`渲染器出不了 example.json：${built.stderr.trim()}`); return; }
    const report = JSON.parse(built.stdout);
    if (example && report.slides !== example.slides.length) problems.push(`example.json 有 ${example.slides.length} 页，渲染出 ${report.slides} 页`);
  } finally { rmSync(out, { recursive: true, force: true }); }
}

/**
 * Validate packs. Returns { ok, python, packs: [{ id, ok, problems, notes, stats, renderer }] }.
 * options: skill (shimo-pptx dir), styles (packs dir; default <skill>/styles, anything else is staging), style (one id),
 * python (path, or false to skip), allowNoPython. Returns also staging: true|false.
 */
export function validatePacks({ skill = DEFAULT_SKILL, styles, style, python, allowNoPython = false } = {}) {
  const registered = join(skill, 'styles');
  const stylesRoot = styles ? resolve(styles) : registered;
  const staging = stylesRoot !== resolve(registered);
  const packIds = root => existsSync(root) ? readdirSync(root, { withFileTypes: true }).filter(entry => entry.isDirectory() && PACK_NAME.test(entry.name)).map(entry => entry.name).sort() : [];
  const ids = packIds(stylesRoot);
  if (style && !ids.includes(style)) throw new Error(`没有样式包 ${style}，可选 ${ids.join(', ') || '无'}`);
  const py = python === false ? null : findPython(python);
  const fonts = chartFonts(skill);
  const packs = [];
  const aliasOwner = new Map();
  for (const id of style ? [style] : ids) {
    const problems = [], notes = [], stats = {};
    const read = checkPack(stylesRoot, staging, id, problems, stats, notes, fonts);
    let renderer = 'skipped';
    if (py && read) { checkRenderer(skill, stylesRoot, id, py, read.example, read.desc, problems); renderer = 'checked'; }
    else if (!py && !allowNoPython) { problems.push('找不到带 python-pptx 的 Python，渲染器检查没跑（用 --python 指定，或 --allow-no-python 跳过）'); renderer = 'missing'; }
    packs.push({ id, problems, notes, stats, renderer, meta: read?.style?.meta });
  }
  // Aliases and ids across every pack (also those not selected), so a new pack cannot take another's name; a staging
  // run also counts the registered packs (checked first, so a clash is reported on the staging pack).
  const everyPack = [...(staging ? packIds(registered).map(id => [registered, id]) : []), ...ids.map(id => [stylesRoot, id])];
  const report = (owner, id, text) => (packs.find(pack => pack.id === id) ?? packs.find(pack => pack.id === owner))?.problems.push(text);
  for (const [root, id] of everyPack) {
    let meta;
    try { meta = JSON.parse(readFileSync(join(root, id, 'style.json'), 'utf8')).meta; } catch { continue; }
    for (const alias of [id, ...(Array.isArray(meta?.aliases) ? meta.aliases : [])]) {
      const key = String(alias).toLowerCase();
      const owner = aliasOwner.get(key);
      const where = root === stylesRoot ? id : `${id}（已登记）`;
      if (owner && owner !== where) report(owner, id, `别名或 id「${alias}」同时属于 ${owner} 和 ${where}`);
      else aliasOwner.set(key, where);
    }
  }
  for (const pack of packs) { pack.ok = pack.problems.length === 0; delete pack.meta; }
  return { ok: packs.every(pack => pack.ok), python: py, staging, packs };
}

function main(argv) {
  const args = { json: false, allowNoPython: false };
  for (let i = 0; i < argv.length; i += 1) {
    const flag = argv[i];
    if (flag === '--json') args.json = true;
    else if (flag === '--allow-no-python') args.allowNoPython = true;
    else if (['--skill', '--styles', '--style', '--python'].includes(flag)) args[flag.slice(2)] = argv[++i];
    else { console.error(`未知参数 ${flag}\n用法：node validate.mjs [--skill <dir>] [--styles <packs dir>] [--style <id>] [--python <path>] [--allow-no-python] [--json]`); return 2; }
  }
  let result;
  try { result = validatePacks({ skill: args.skill ? resolve(args.skill) : undefined, styles: args.styles, style: args.style, python: args.python, allowNoPython: args.allowNoPython }); }
  catch (error) { console.error(error.message); return 2; }
  if (args.json) console.log(JSON.stringify(result, null, 2));
  else {
    console.log(`Python：${result.python ?? '未找到'}${result.staging ? '（暂存校验：cleared=false 只记待登记）' : ''}`);
    for (const pack of result.packs) {
      console.log(`${pack.ok ? '✓' : '✗'} ${pack.id}  渲染器：${pack.renderer}`);
      const s = pack.stats;
      if (s.layouts) console.log(`  版式 ${s.layouts} 个；照片槽 ${s.photos.slots}（应有图 ${s.photos.required}，单页最多 ${s.photos.per_slide_max}）；最小字号 文字 ${s.minTextPt ?? '-'}pt / 图表 ${s.minChartPt === null ? '-（无图表）' : `${s.minChartPt}pt`}；缩略图 ${s.thumb ?? '-'}${s.previews ? `；样张 ${s.previews} 张` : ''}`);
      if (s.textBoxes) console.log(`  每页最多文本框：${Object.entries(s.textBoxes).map(([k, v]) => `${k} ${v}`).join('，')}`);
      for (const note of pack.notes) console.log(`  注：${note}`);
      for (const problem of pack.problems) console.log(`  ✗ ${problem}`);
    }
  }
  return result.ok ? 0 : 1;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = main(process.argv.slice(2));
