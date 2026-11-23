import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseExpr } from './style-expr.mjs';

/**
 * A style pack's layout description, styles/<name>/layouts.json: the offline page types brand_deck.py draws and the
 * live layouts and slots presentation_create hands the model. One file, read by both modes; the format is
 * tooling/style-distill/pack-format.md in the source repository. readLayouts checks its shape and every expression
 * and throws an Error naming the first problem.
 */
export const OBJECT_TYPES = Object.freeze(['text', 'rect', 'shape', 'hline', 'vline', 'arrow', 'tag', 'highlight', 'image', 'chart', 'legend', 'table', 'header', 'logo', 'group', 'repeat']);
const FIELD_TYPES = new Set(['text', 'list', 'object', 'image', 'chart', 'table']);
const PLACEHOLDERS = new Set(['title', 'subtitle', 'body']);
/** Object keys holding expressions (or numbers); `text`, `items`, `chart`, `table` and `when` must be expressions. */
const EXPR_KEYS = new Set(['text', 'items', 'chart', 'table', 'when', 'x', 'y', 'w', 'h', 'size', 'weight', 'lh', 'max_h', 'alpha', 'mask', 'pad', 'left', 'right', 'max_w', 'bottom', 'label_x', 'label_w', 'spacing']);
/** Colour keys: '#RRGGBB', palette roles ('primary', 'dividerOnPrimary|lineOnPrimary') or an expression. */
export const COLOR_KEYS = Object.freeze(['color', 'fill', 'outline', 'panel', 'bg', 'rule', 'head_fill', 'head_color', 'body_fill', 'body_color', 'border']);
const isHex = value => /^#[0-9A-Fa-f]{6}$/.test(value);
const isRoles = value => /^[A-Za-z_][A-Za-z0-9_]*(?:\|[A-Za-z_][A-Za-z0-9_]*)*$/.test(value);
/** A colour that needs parsing as an expression (not hex, not role names), else null. */
export const colorExpression = value => (typeof value === 'string' && !isHex(value) && !isRoles(value) ? value : null);
const MUST_BE_EXPR = { text: ['text'], tag: ['text'], highlight: ['text'], repeat: ['items'], chart: ['chart'], legend: ['chart'], table: ['table'] };

const isBox = box => Array.isArray(box) && box.length === 4 && box.every(v => typeof v === 'number' && Number.isFinite(v));
const isObject = value => value !== null && typeof value === 'object' && !Array.isArray(value);

function fail(where, message) { throw new Error(`${where}：${message}`); }

/** Every expression string in a layout's lets, vars-free objects and nested objects, with where it sits. */
function* expressions(layout, id) {
  for (const [name, source] of Object.entries(layout.let ?? {})) yield { where: `${id}.let.${name}`, source };
  if (colorExpression(layout.background)) yield { where: `${id}.background`, source: layout.background };
  function* walk(objects, path) {
    for (const [index, object] of objects.entries()) {
      const at = `${path}[${index}]`;
      for (const [name, source] of Object.entries(object.let ?? {})) yield { where: `${at}.let.${name}`, source };
      for (const [key, value] of Object.entries(object)) {
        if (EXPR_KEYS.has(key) && typeof value === 'string') yield { where: `${at}.${key}`, source: value };
        if (COLOR_KEYS.includes(key) && colorExpression(value)) yield { where: `${at}.${key}`, source: value };
      }
      if (object.objects) yield* walk(object.objects, `${at}.objects`);
    }
  }
  yield* walk(layout.objects ?? [], `${id}.objects`);
}

function checkObjects(objects, where) {
  if (!Array.isArray(objects)) fail(where, 'objects 必须是数组');
  objects.forEach((object, index) => {
    const at = `${where}[${index}]`;
    if (!isObject(object)) fail(at, '对象必须是 JSON 对象');
    if (!OBJECT_TYPES.includes(object.type)) fail(at, `未知对象类型 ${JSON.stringify(object.type)}，可选 ${OBJECT_TYPES.join(', ')}`);
    for (const key of MUST_BE_EXPR[object.type] ?? []) if (typeof object[key] !== 'string') fail(at, `${object.type} 的 ${key} 必须是表达式字符串`);
    if (object.when !== undefined && typeof object.when !== 'string') fail(at, 'when 必须是表达式字符串');
    if (object.type === 'image' && typeof object.slot !== 'string') fail(at, 'image 要写 slot（字段名）');
    if (object.type === 'repeat' && !(Number.isInteger(object.max) && object.max > 0)) fail(at, 'repeat 要写 max（正整数）：渲染器按它截断，校验按它算文本框上限');
    for (const key of COLOR_KEYS) if (object[key] !== undefined && object[key] !== null && typeof object[key] !== 'string') fail(at, `${key} 必须是颜色字符串或 null`);
    if ((object.type === 'group' || object.type === 'repeat') || object.objects !== undefined) checkObjects(object.objects ?? [], `${at}.objects`);
  });
}

function resolve(layouts) {
  const out = {};
  for (const [id, layout] of Object.entries(layouts)) {
    if (!isObject(layout)) fail(`offline.layouts.${id}`, '版式必须是对象');
    let merged = layout;
    if (layout.extends !== undefined) {
      const base = layouts[layout.extends];
      if (!isObject(base) || base.extends !== undefined) fail(`offline.layouts.${id}`, `extends=${JSON.stringify(layout.extends)} 不存在或也是变体`);
      merged = { ...base, ...layout, vars: { ...base.vars, ...layout.vars }, let: { ...base.let, ...layout.let } };
    }
    out[id] = { fields: {}, ...merged };
  }
  return out;
}

/** The slots (image, chart, table) of a page type: field, box in px when its geometry is fixed, and whether it is conditional. */
function slotsOf(layout) {
  const slots = [];
  const walk = (objects, conditional) => {
    for (const object of objects) {
      const cond = conditional || object.when !== undefined || object.type === 'repeat';
      if (['image', 'chart', 'table'].includes(object.type)) {
        const values = ['x', 'y', 'w', 'h'].map(key => object[key]);
        const fixed = values.every(value => typeof value === 'number');
        const field = object.type === 'image' ? object.slot : /d\['([^']+)'\]/.exec(object[object.type])?.[1] ?? null;
        slots.push({ id: object.type === 'image' ? object.slot : `${object.type}${slots.filter(s => s.kind === object.type).length + 1}`, kind: object.type, field, box: fixed ? values : null, dynamic: !fixed, conditional: cond });
      }
      if (object.objects) walk(object.objects, cond);
    }
  };
  walk(layout.objects ?? [], false);
  return slots;
}

/** Check a parsed layouts.json; returns the normalised description. */
export function checkLayouts(desc, where = 'layouts.json') {
  if (!isObject(desc)) fail(where, '顶层必须是对象');
  if (desc.version !== 1) fail(where, 'version 应为 1');
  const { offline, live } = desc;
  if (!isObject(offline) || !isObject(offline.layouts) || !Object.keys(offline.layouts).length) fail(where, '缺少 offline.layouts');
  if (!isObject(offline.canvas) || offline.canvas.width !== 960 || offline.canvas.height !== 540) fail(where, 'offline.canvas 应为 960×540 px');
  const layouts = resolve(offline.layouts);
  const order = offline.order ?? Object.keys(layouts);
  if (!Array.isArray(order) || order.some(id => !layouts[id]) || new Set(order).size !== order.length || order.length !== Object.keys(layouts).length) fail(where, 'offline.order 要正好列出每个版式一次');
  for (const [id, layout] of Object.entries(layouts)) {
    const at = `${where} offline.layouts.${id}`;
    if (typeof layout.description !== 'string' || !layout.description) fail(at, '缺少 description');
    if (!isObject(layout.fields)) fail(at, 'fields 必须是对象');
    for (const [name, field] of Object.entries(layout.fields)) if (!isObject(field) || !FIELD_TYPES.has(field.type)) fail(`${at}.fields.${name}`, `type 应为 ${[...FIELD_TYPES].join(' / ')}`);
    if (layout.let !== undefined && !isObject(layout.let)) fail(at, 'let 必须是对象');
    if (layout.background !== undefined && typeof layout.background !== 'string') fail(at, 'background 必须是颜色字符串');
    checkObjects(layout.objects, `${at}.objects`);
    for (const { where: place, source } of expressions(layout, id)) {
      try { parseExpr(source); } catch (error) { fail(`${where} ${place}`, error.message); }
    }
  }
  if (desc.deck !== undefined) {
    if (!isObject(desc.deck)) fail(where, 'deck 必须是对象');
    for (const end of ['first', 'last']) if (desc.deck[end] !== undefined && !layouts[desc.deck[end]]) fail(where, `deck.${end}=${JSON.stringify(desc.deck[end])} 不是版式`);
  }
  if (!isObject(live) || !isObject(live.canvas) || live.canvas.width !== 720 || live.canvas.height !== 405 || live.canvas.unit !== 'pt') fail(where, 'live.canvas 应为 720×405 pt');
  if (!Array.isArray(live.layouts) || !live.layouts.length) fail(where, '缺少 live.layouts');
  const liveIds = new Set();
  for (const [index, layout] of live.layouts.entries()) {
    const at = `${where} live.layouts[${index}]`;
    if (!isObject(layout) || typeof layout.id !== 'string' || typeof layout.name !== 'string') fail(at, '要有 id 和 name');
    if (liveIds.has(layout.id)) fail(at, `id ${layout.id} 重复`);
    liveIds.add(layout.id);
    if (!Array.isArray(layout.placeholders) || layout.placeholders.some(p => !PLACEHOLDERS.has(p))) fail(at, 'placeholders 只能是 title / subtitle / body');
    if (!isObject(layout.boxes) || Object.keys(layout.boxes).sort().join() !== [...layout.placeholders].sort().join() || !Object.values(layout.boxes).every(isBox)) fail(at, 'boxes 要给每个占位符一个 [l, t, w, h]');
  }
  if (!Array.isArray(live.slots)) fail(where, 'live.slots 必须是数组');
  const slotIds = new Set();
  for (const [index, slot] of live.slots.entries()) {
    const at = `${where} live.slots[${index}]`;
    if (!isObject(slot) || typeof slot.id !== 'string' || slotIds.has(slot.id)) fail(at, 'id 缺失或重复');
    slotIds.add(slot.id);
    if (!liveIds.has(slot.layout)) fail(at, `layout ${JSON.stringify(slot.layout)} 不在 live.layouts 里`);
    if (!['image', 'text'].includes(slot.kind)) fail(at, 'kind 应为 image 或 text');
    if (!isBox(slot.box) || (slot.text_box !== undefined && !isBox(slot.text_box))) fail(at, 'box / text_box 应为 [l, t, w, h]');
    if (slot.kind === 'image' && !/^\d+x\d+$/.test(slot.art_size ?? '')) fail(at, '图片槽要写 art_size（如 1920x1080）');
  }
  return {
    version: 1,
    deck: desc.deck ?? {},
    offline: {
      canvas: offline.canvas,
      order,
      layouts: Object.fromEntries(order.map(id => [id, { title: layouts[id].title ?? id, description: layouts[id].description, fields: layouts[id].fields, slots: slotsOf(layouts[id]), raw: layouts[id] }])),
    },
    live: { canvas: live.canvas, layouts: live.layouts, slots: live.slots },
  };
}

/** Read and check styles/<name>/layouts.json under a styles root. */
export function readLayouts(root, name) {
  const path = join(root, name, 'layouts.json');
  let desc;
  try { desc = JSON.parse(readFileSync(path, 'utf8')); }
  catch (error) { throw new Error(`${name}/layouts.json 读不了（${error?.code ?? error?.name ?? 'error'}）`); }
  return checkLayouts(desc, `${name}/layouts.json`);
}

/** What presentation_create hands the model: the live part as written (pt). */
export function liveLayouts(description) {
  const { canvas, layouts, slots } = description.live;
  return { canvas, layouts, slots };
}
