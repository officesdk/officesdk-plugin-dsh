import { SuiteError } from '../errors.mjs';

/**
 * Structured presentation edits. The model never supplies code: the editor frame shares the
 * application's origin, so every operation here maps to a fixed Office.js call in the frame.
 */
export const PPTX_ALL_OPS = Object.freeze(['setText', 'replaceText', 'setFont', 'setAlignment', 'setFill', 'setGeometry', 'addTextBox', 'addShape', 'addImage', 'addTable', 'addChart', 'setBackground', 'deleteShape', 'addSlide', 'deleteSlide', 'moveSlide']);
/**
 * Implemented but not offered: a table or chart added through the JS API shows in the editor, but office-kit's
 * Kit then fails to save the whole file (`convert-failed`, office-kit request K6). Offer them once Kit saves them.
 */
export const PPTX_HELD_OPS = Object.freeze(['addTable', 'addChart']);
/** The operations the model may send. */
export const PPTX_OPS = Object.freeze(PPTX_ALL_OPS.filter(op => !PPTX_HELD_OPS.includes(op)));
export const PPTX_LIMITS = Object.freeze({
  maxOps: 200, maxText: 20_000, maxSummaryChars: 60_000, maxShapeText: 2_000,
  maxImageBytes: 10 * 1024 * 1024, maxBatchImageBytes: 20 * 1024 * 1024,
  maxTableRows: 30, maxTableColumns: 12, maxCellText: 500,
  maxChartSeries: 6, maxChartCategories: 24,
  // New slides per edit: a batch plays together and is undone as a whole when it fails, so it stays small (spec §11.2).
  maxNewSlides: 4,
});
/** The `invalid-edit` message when a batch adds more slides than one edit may. */
export const NEW_SLIDES_MESSAGE = `At most ${PPTX_LIMITS.maxNewSlides} new slides per edit; build the deck in several presentation_edit calls.`;
/** Counts addSlide operations, before anything is checked (the Host turns a batch with too many away early). */
export const countNewSlides = ops => Array.isArray(ops) ? ops.filter(op => op?.op === 'addSlide').length : 0;
/** Preset shapes the model may draw (PowerPoint.GeometricShapeType values). */
export const PPTX_SHAPES = Object.freeze(['Rectangle', 'RoundRectangle', 'Ellipse', 'Triangle', 'RightTriangle', 'Diamond', 'Pentagon', 'Hexagon', 'Octagon', 'Parallelogram', 'Trapezoid',
  'RightArrow', 'LeftArrow', 'UpArrow', 'DownArrow', 'LeftRightArrow', 'Chevron', 'HomePlate', 'Plus', 'Star5', 'Donut', 'Can', 'Cube', 'FlowChartProcess', 'FlowChartDecision', 'WedgeRectCallout']);
export const PPTX_CHARTS = Object.freeze(['column', 'bar', 'line', 'pie', 'donut', 'area']);
/** Placeholders addSlide can fill on a layout, by the placeholder type each name matches. */
const PLACEHOLDER_KEYS = Object.freeze({ title: ['Title', 'CenterTitle'], subtitle: ['Subtitle'], body: ['Body', 'Content', 'Object'] });

const COLOR = /^#[0-9A-Fa-f]{6}$/;
const REF = /^[A-Za-z][A-Za-z0-9_-]{0,31}$/;
const ALIGN = ['Left', 'Center', 'Right', 'Justify'];
const GEOMETRY = ['left', 'top', 'width', 'height'];
const FONT = ['bold', 'italic', 'underline', 'size', 'color', 'name'];
const IMAGE_TYPES = ['image/png', 'image/jpeg'];
const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/;

const invalid = (index, message) => { throw new SuiteError('invalid-edit', `Operation ${index + 1}: ${message}`, { index }); };
const text = (value, index, field, max = PPTX_LIMITS.maxText) => {
  if (typeof value !== 'string') invalid(index, `${field} must be a string.`);
  if (value.length > max) invalid(index, `${field} is longer than ${max} characters.`);
  return value;
};
const number = (value, index, field, min, max) => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) invalid(index, `${field} must be a number between ${min} and ${max}.`);
  return value;
};
const color = (value, index, field) => { if (!COLOR.test(value ?? '')) invalid(index, `${field} must look like #RRGGBB.`); return value; };
/** A new object's box: all four values, in points. */
const box = (op, index) => Object.fromEntries(GEOMETRY.map(key => [key, number(op[key], index, key, key === 'width' || key === 'height' ? 1 : -5000, 5000)]));
const fontOf = (op, index, { required = false } = {}) => {
  const font = {};
  for (const key of FONT) if (op[key] !== undefined) font[key] = op[key];
  if (required && !Object.keys(font).length) invalid(index, 'set at least one of bold, italic, underline, size, color, name.');
  for (const key of ['bold', 'italic', 'underline']) if (key in font && typeof font[key] !== 'boolean') invalid(index, `${key} must be true or false.`);
  if ('size' in font) number(font.size, index, 'size', 1, 400);
  if ('color' in font) color(font.color, index, 'color');
  if ('name' in font && (typeof font.name !== 'string' || !font.name || font.name.length > 64)) invalid(index, 'name must be a font name.');
  return font;
};
/** Bytes of a base64 payload, without decoding it. */
const base64Bytes = value => Math.floor(value.length * 3 / 4) - (value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0);

/**
 * Checks a batch against the deck the user is looking at and resolves what each step changes.
 * `deck` is the frame's `inspect` result; every target must exist and stay alive within the batch.
 * An object a step creates may carry a `ref`; later steps in the same batch target it with `slideRef`
 * or `shapeRef`, and the frame swaps each ref for the id the editor gave the object.
 */
export function validatePptxOps(ops, deck, { enabled = PPTX_OPS } = {}) {
  if (!Array.isArray(ops) || ops.length === 0) throw new SuiteError('invalid-edit', 'Provide at least one operation.');
  if (ops.length > PPTX_LIMITS.maxOps) throw new SuiteError('invalid-edit', `At most ${PPTX_LIMITS.maxOps} operations per edit.`);
  // Keys: a live id as it is, a ref of this batch as `@ref`; ids never start with `@` in a key built here.
  const slides = new Map(deck.slides.map(slide => [slide.id, { ...slide, shapes: new Map(slide.shapes.map(shape => [shape.id, { ...shape }])) }]));
  const layouts = new Map((deck.layouts ?? []).flatMap(layout => [[layout.id, layout], [`name:${layout.name}`, layout]]));
  const refs = new Set();
  let count = deck.slides.length, imageBytes = 0;
  if (countNewSlides(ops) > PPTX_LIMITS.maxNewSlides) throw new SuiteError('invalid-edit', NEW_SLIDES_MESSAGE);
  const claim = (op, index) => {
    if (op.ref === undefined) return undefined;
    if (typeof op.ref !== 'string' || !REF.test(op.ref)) invalid(index, 'ref must start with a letter and use letters, digits, - or _ (at most 32).');
    if (refs.has(op.ref)) invalid(index, `ref ${JSON.stringify(op.ref)} is already used in this edit.`);
    refs.add(op.ref);
    return op.ref;
  };
  const slideTarget = (op, index) => {
    if (op.slideRef !== undefined && op.slideId !== undefined) invalid(index, 'give slideId or slideRef, not both.');
    if (op.slideRef !== undefined) {
      const slide = slides.get(`@${op.slideRef}`);
      if (!slide) invalid(index, `slideRef ${JSON.stringify(op.slideRef)} is not a slide added earlier in this edit.`);
      return { slide, target: { slideRef: op.slideRef } };
    }
    const slide = slides.get(op.slideId);
    if (!slide || typeof op.slideId !== 'string') invalid(index, `slide ${JSON.stringify(op.slideId)} is not in the presentation.`);
    return { slide, target: { slideId: op.slideId } };
  };
  const where = (op) => op.slideRef !== undefined ? `@${op.slideRef}` : op.slideId;
  const shapeTarget = (op, index) => {
    const { slide, target } = slideTarget(op, index);
    if (op.shapeRef !== undefined && op.shapeId !== undefined) invalid(index, 'give shapeId or shapeRef, not both.');
    if (op.shapeRef !== undefined) {
      const shape = slide.shapes.get(`@${op.shapeRef}`);
      if (!shape) invalid(index, `shapeRef ${JSON.stringify(op.shapeRef)} is not a shape added earlier in this edit on slide ${where(op)}.`);
      return { slide, shape, target: { ...target, shapeRef: op.shapeRef } };
    }
    const shape = slide.shapes.get(op.shapeId);
    if (!shape || typeof op.shapeId !== 'string') invalid(index, `shape ${JSON.stringify(op.shapeId)} is not on slide ${where(op)}.`);
    return { slide, shape, target: { ...target, shapeId: op.shapeId } };
  };
  const textShape = (op, index) => {
    const found = shapeTarget(op, index);
    if (typeof found.shape.text !== 'string') invalid(index, `shape ${op.shapeId ?? `@${op.shapeRef}`} has no text.`);
    return found;
  };
  /** Registers a shape the step creates so later steps can target it. */
  const created = (slide, ref, shape) => { if (ref) slide.shapes.set(`@${ref}`, { id: `@${ref}`, ...shape }); };
  const steps = ops.map((op, index) => {
    if (!op || typeof op !== 'object' || !PPTX_ALL_OPS.includes(op.op)) invalid(index, `unknown operation ${JSON.stringify(op?.op)}.`);
    if (!enabled.includes(op.op)) invalid(index, `${op.op} is not available yet: the document could not be saved with it. Tell the user plainly instead of approximating it.`);
    switch (op.op) {
      case 'setText': {
        const { shape, target } = textShape(op, index);
        const step = { op: 'setText', ...target, before: shape.text, text: text(op.text, index, 'text') };
        shape.text = step.text;
        return step;
      }
      case 'replaceText': {
        const { shape, target } = textShape(op, index);
        const find = text(op.find, index, 'find'), replace = text(op.replace, index, 'replace');
        if (!find) invalid(index, 'find must not be empty.');
        const at = shape.text.indexOf(find);
        const name = op.shapeId ?? `@${op.shapeRef}`;
        if (at < 0) invalid(index, `${JSON.stringify(find)} does not occur in shape ${name}.`);
        if (shape.text.includes(find, at + find.length)) invalid(index, `${JSON.stringify(find)} occurs more than once in shape ${name}; quote more context.`);
        const step = { op: 'setText', ...target, before: shape.text, text: shape.text.slice(0, at) + replace + shape.text.slice(at + find.length), find, replace };
        if (step.text.length > PPTX_LIMITS.maxText) invalid(index, 'the resulting text is too long.');
        shape.text = step.text;
        return step;
      }
      case 'setFont': {
        const { target } = textShape(op, index);
        return { op: 'setFont', ...target, font: fontOf(op, index, { required: true }) };
      }
      case 'setAlignment': {
        const { target } = textShape(op, index);
        if (!ALIGN.includes(op.align)) invalid(index, `align must be one of ${ALIGN.join(', ')}.`);
        return { op: 'setAlignment', ...target, align: op.align };
      }
      case 'setFill': {
        const { target } = shapeTarget(op, index);
        return { op: 'setFill', ...target, color: color(op.color, index, 'color') };
      }
      case 'setGeometry': {
        const { shape, target } = shapeTarget(op, index);
        const geometry = {};
        for (const key of GEOMETRY) if (op[key] !== undefined) geometry[key] = number(op[key], index, key, key === 'width' || key === 'height' ? 1 : -5000, 5000);
        if (!Object.keys(geometry).length) invalid(index, 'set at least one of left, top, width, height.');
        const step = { op: 'setGeometry', ...target, geometry, before: Object.fromEntries(GEOMETRY.map(key => [key, shape[key]])) };
        Object.assign(shape, geometry);
        return step;
      }
      case 'addTextBox': {
        const { slide, target } = slideTarget(op, index);
        const ref = claim(op, index);
        const geometry = box(op, index);
        const font = fontOf(op, index);
        if (op.align !== undefined && !ALIGN.includes(op.align)) invalid(index, `align must be one of ${ALIGN.join(', ')}.`);
        const step = { op: 'addTextBox', ...target, text: text(op.text, index, 'text'), geometry, ...(ref && { ref }), ...(Object.keys(font).length && { font }), ...(op.align && { align: op.align }) };
        created(slide, ref, { type: 'TextBox', ...geometry, text: step.text });
        return step;
      }
      case 'addShape': {
        const { slide, target } = slideTarget(op, index);
        const ref = claim(op, index);
        if (!PPTX_SHAPES.includes(op.shape)) invalid(index, `shape must be one of ${PPTX_SHAPES.join(', ')}.`);
        const geometry = box(op, index);
        const step = { op: 'addShape', ...target, shape: op.shape, geometry, ...(ref && { ref }) };
        if (op.fill !== undefined) step.fill = color(op.fill, index, 'fill');
        if (op.line !== undefined) step.line = op.line === 'none' ? 'none' : color(op.line, index, 'line');
        if (op.text !== undefined) step.text = text(op.text, index, 'text');
        const font = fontOf(op, index);
        if (Object.keys(font).length) { if (step.text === undefined) invalid(index, 'font settings need text.'); step.font = font; }
        if (op.align !== undefined) { if (!ALIGN.includes(op.align)) invalid(index, `align must be one of ${ALIGN.join(', ')}.`); step.align = op.align; }
        created(slide, ref, { type: 'GeometricShape', ...geometry, text: step.text ?? '' });
        return step;
      }
      case 'addImage': {
        const { slide, target } = slideTarget(op, index);
        const ref = claim(op, index);
        // The Host reads `path` and fills these in; the model never sends image bytes itself.
        if (!IMAGE_TYPES.includes(op.imageType) || typeof op.imageBase64 !== 'string' || !op.imageBase64 || !BASE64.test(op.imageBase64)) invalid(index, 'the image was not loaded; give the path of a PNG or JPEG file.');
        const bytes = base64Bytes(op.imageBase64);
        if (bytes > PPTX_LIMITS.maxImageBytes) invalid(index, `the image is larger than ${PPTX_LIMITS.maxImageBytes / 1024 / 1024} MB.`);
        imageBytes += bytes;
        if (imageBytes > PPTX_LIMITS.maxBatchImageBytes) invalid(index, `images in one edit add up to more than ${PPTX_LIMITS.maxBatchImageBytes / 1024 / 1024} MB; split the edit.`);
        const geometry = box(op, index);
        created(slide, ref, { type: 'Image', ...geometry });
        return { op: 'addImage', ...target, geometry, imageBase64: op.imageBase64, imageType: op.imageType, ...(typeof op.path === 'string' && { path: op.path }), ...(ref && { ref }) };
      }
      case 'addTable': {
        const { slide, target } = slideTarget(op, index);
        const ref = claim(op, index);
        const values = op.values;
        if (!Array.isArray(values) || !values.length || values.length > PPTX_LIMITS.maxTableRows) invalid(index, `values must be 1–${PPTX_LIMITS.maxTableRows} rows.`);
        const columns = Array.isArray(values[0]) ? values[0].length : 0;
        if (!columns || columns > PPTX_LIMITS.maxTableColumns) invalid(index, `each row needs 1–${PPTX_LIMITS.maxTableColumns} cells.`);
        const rows = values.map((row, r) => {
          if (!Array.isArray(row) || row.length !== columns) invalid(index, `row ${r + 1} must have ${columns} cells, like the first row.`);
          return row.map(cell => text(typeof cell === 'number' ? String(cell) : cell, index, 'cell', PPTX_LIMITS.maxCellText));
        });
        const geometry = box(op, index);
        created(slide, ref, { type: 'Table', ...geometry });
        return { op: 'addTable', ...target, values: rows, geometry, ...(ref && { ref }) };
      }
      case 'addChart': {
        const { slide, target } = slideTarget(op, index);
        const ref = claim(op, index);
        const chartType = op.chartType ?? 'column';
        if (!PPTX_CHARTS.includes(chartType)) invalid(index, `chartType must be one of ${PPTX_CHARTS.join(', ')}.`);
        const categories = op.categories, series = op.series;
        if (!Array.isArray(categories) || !categories.length || categories.length > PPTX_LIMITS.maxChartCategories) invalid(index, `categories must list 1–${PPTX_LIMITS.maxChartCategories} names.`);
        categories.forEach(name => text(typeof name === 'number' ? String(name) : name, index, 'category', 100));
        if (!Array.isArray(series) || !series.length || series.length > PPTX_LIMITS.maxChartSeries) invalid(index, `series must list 1–${PPTX_LIMITS.maxChartSeries} series; merge or split the chart.`);
        if ((chartType === 'pie' || chartType === 'donut') && series.length !== 1) invalid(index, `a ${chartType} chart takes exactly one series.`);
        for (const entry of series) {
          if (!entry || typeof entry !== 'object') invalid(index, 'each series needs a name and values.');
          text(entry.name, index, 'series name', 100);
          if (!Array.isArray(entry.values) || entry.values.length !== categories.length || entry.values.some(value => typeof value !== 'number' || !Number.isFinite(value))) invalid(index, `series ${JSON.stringify(entry.name)} needs ${categories.length} numbers, one per category.`);
        }
        const geometry = box(op, index);
        const step = { op: 'addChart', ...target, chartType, geometry, data: [['', ...categories.map(String)], ...series.map(entry => [entry.name, ...entry.values])], ...(ref && { ref }) };
        if (op.title !== undefined) step.title = text(op.title, index, 'title', 200);
        if (op.seriesColors !== undefined) {
          if (!Array.isArray(op.seriesColors) || op.seriesColors.length > series.length) invalid(index, 'seriesColors lists at most one color per series.');
          step.seriesColors = op.seriesColors.map(value => color(value, index, 'seriesColors'));
        }
        if (op.legend !== undefined) { if (!['none', 'top', 'bottom', 'left', 'right'].includes(op.legend)) invalid(index, 'legend must be none, top, bottom, left or right.'); step.legend = op.legend; }
        created(slide, ref, { type: 'Chart', ...geometry });
        return step;
      }
      case 'setBackground': {
        const { target } = slideTarget(op, index);
        return { op: 'setBackground', ...target, color: color(op.color, index, 'color') };
      }
      case 'deleteShape': {
        const { slide, target } = shapeTarget(op, index);
        slide.shapes.delete(op.shapeRef !== undefined ? `@${op.shapeRef}` : op.shapeId);
        return { op: 'deleteShape', ...target };
      }
      case 'addSlide': {
        const ref = claim(op, index);
        const at = op.index === undefined ? count : number(op.index, index, 'index', 0, count);
        if (!Number.isInteger(at)) invalid(index, 'index must be a whole number.');
        const step = { op: 'addSlide', index: at, ...(ref && { ref }) };
        let layout;
        if (op.layout !== undefined) {
          layout = typeof op.layout === 'string' ? layouts.get(op.layout) ?? layouts.get(`name:${op.layout}`) : undefined;
          if (!layout) invalid(index, `layout ${JSON.stringify(op.layout)} is not in this presentation; use an id or name from presentation_read's layouts.`);
          step.layoutId = layout.id;
          step.layoutName = layout.name;
        }
        if (op.placeholders !== undefined) {
          if (!step.layoutId) invalid(index, 'placeholders need a layout.');
          if (!op.placeholders || typeof op.placeholders !== 'object' || Array.isArray(op.placeholders)) invalid(index, 'placeholders maps title, subtitle or body to text.');
          const fill = {};
          for (const [key, value] of Object.entries(op.placeholders)) {
            if (!(key in PLACEHOLDER_KEYS)) invalid(index, `placeholders takes ${Object.keys(PLACEHOLDER_KEYS).join(', ')}.`);
            fill[key] = text(value, index, `placeholders.${key}`);
            // The read lists each layout's placeholders: a missing one fails here, before anything is written.
            if (fill[key] && layout?.placeholders && !layout.placeholders.includes(key)) invalid(index, `layout ${JSON.stringify(step.layoutName ?? step.layoutId)} has no ${key} placeholder (it has: ${layout.placeholders.join(', ') || 'none'}).`);
          }
          if (Object.keys(fill).length) step.placeholders = fill;
        }
        count++;
        if (ref) slides.set(`@${ref}`, { id: `@${ref}`, shapes: new Map() });
        return step;
      }
      case 'deleteSlide': {
        const { target } = slideTarget(op, index);
        if (count <= 1) invalid(index, 'a presentation keeps at least one slide.');
        slides.delete(op.slideRef !== undefined ? `@${op.slideRef}` : op.slideId); count--;
        return { op: 'deleteSlide', ...target };
      }
      case 'moveSlide': {
        const { target } = slideTarget(op, index);
        const to = number(op.index, index, 'index', 0, count - 1);
        if (!Number.isInteger(to)) invalid(index, 'index must be a whole number.');
        return { op: 'moveSlide', ...target, index: to };
      }
    }
    return undefined;
  });
  const touched = new Set(steps.map(step => step.slideId).filter(Boolean));
  const reasons = [];
  if (steps.some(step => step.op === 'deleteSlide' || step.op === 'deleteShape')) reasons.push('deletes');
  if (deck.slides.length > 3 && touched.size > deck.slides.length / 2) reasons.push('wide');
  return { steps, touchedSlideIds: [...touched], danger: reasons.length > 0, reasons };
}

/** The placeholder names addSlide fills, by the placeholder type the editor reports. */
export function placeholderKey(type) {
  for (const [key, types] of Object.entries(PLACEHOLDER_KEYS)) if (types.includes(type)) return key;
  return undefined;
}

const clip = (value, size = 120) => value.length > size ? value.slice(0, size - 1) + '…' : value;

/** Rows for the confirmation card: where, what was there, what will be there. */
export function describePptxSteps(steps, deck) {
  const position = new Map(deck.slides.map((slide, index) => [slide.id, index + 1]));
  const shapeName = new Map(deck.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape.name || shape.id])));
  const slideLabel = step => step.slideRef !== undefined ? `@${step.slideRef}` : position.has(step.slideId) ? `#${position.get(step.slideId)}` : step.slideId;
  const shapeLabel = step => step.shapeRef !== undefined ? `@${step.shapeRef}` : shapeName.get(`${step.slideId}/${step.shapeId}`) ?? step.shapeId;
  return steps.map(step => {
    const where = step.slideId !== undefined || step.slideRef !== undefined ? `${slideLabel(step)}${step.shapeId !== undefined || step.shapeRef !== undefined ? ` · ${shapeLabel(step)}` : ''}` : '';
    switch (step.op) {
      case 'setText': return { op: step.op, where, before: clip(step.find ?? step.before), after: clip(step.replace ?? step.text) };
      case 'setFont': return { op: step.op, where, after: Object.entries(step.font).map(([key, value]) => `${key}=${value}`).join(', ') };
      case 'setAlignment': return { op: step.op, where, after: step.align };
      case 'setFill': return { op: step.op, where, after: step.color };
      case 'setGeometry': return { op: step.op, where, before: GEOMETRY.filter(key => key in step.geometry).map(key => `${key}=${Math.round(step.before[key] ?? 0)}`).join(', '), after: Object.entries(step.geometry).map(([key, value]) => `${key}=${Math.round(value)}`).join(', ') };
      case 'addTextBox': return { op: step.op, where, after: clip(step.text) };
      case 'addShape': return { op: step.op, where, after: clip(step.text ? `${step.shape}: ${step.text}` : step.shape) };
      case 'addImage': return { op: step.op, where, after: clip(step.path ?? step.imageType) };
      case 'addTable': return { op: step.op, where, after: clip(`${step.values.length}×${step.values[0].length}: ${step.values[0].join(' | ')}`) };
      case 'addChart': return { op: step.op, where, after: clip(`${step.chartType}${step.title ? `: ${step.title}` : ''}`) };
      case 'setBackground': return { op: step.op, where, after: step.color };
      case 'deleteShape': return { op: step.op, where };
      case 'addSlide': return { op: step.op, where: `#${step.index + 1}${step.ref ? ` @${step.ref}` : ''}`, ...((step.layoutName || step.placeholders) && { after: clip([step.layoutName, step.placeholders?.title].filter(Boolean).join(': ')) }) };
      case 'deleteSlide': return { op: step.op, where };
      case 'moveSlide': return { op: step.op, where, after: `#${step.index + 1}` };
    }
    return { op: step.op, where };
  });
}

/** What the model reads: ids it can target, the text in each shape and where it sits. */
export function summarizeDeck(deck, { maxChars = PPTX_LIMITS.maxSummaryChars, maxShapeText = PPTX_LIMITS.maxShapeText } = {}) {
  let budget = maxChars, truncated = false;
  const slides = deck.slides.map((slide, index) => ({
    id: slide.id, index,
    shapes: slide.shapes.map(shape => {
      const entry = { id: shape.id, name: shape.name || undefined, type: shape.type, box: GEOMETRY.map(key => Math.round(shape[key] ?? 0)) };
      if (typeof shape.text === 'string') {
        let value = shape.text.length > maxShapeText ? (truncated = true, shape.text.slice(0, maxShapeText)) : shape.text;
        if (value.length > budget) { value = value.slice(0, Math.max(0, budget)); truncated = true; }
        budget -= value.length;
        entry.text = value;
      }
      return entry;
    }),
  }));
  return { slideSize: deck.slideSize, ...(deck.layouts?.length && { layouts: deck.layouts.map(({ id, name, placeholders }) => ({ id, name, ...(placeholders && { placeholders }) })) }),
    selectedSlideIds: deck.selectedSlideIds ?? [], selectedShapeIds: deck.selectedShapeIds ?? [], slides, ...(truncated && { truncated: true }) };
}
