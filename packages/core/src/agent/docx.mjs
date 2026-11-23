import { SuiteError } from '../errors.mjs';

/**
 * Structured Word edits (spec 261105 §3.2). The model never supplies code: each operation maps to a fixed Word JS API
 * call on the editor page. There is no rollback (office-kit K3): a batch that fails or is stopped part way keeps the
 * steps it finished, and says how far it got.
 *
 * Paragraph and table numbers are the ones document_read gave before this call; the page maps them to where those
 * paragraphs are now as the batch inserts and deletes. Objects a step creates carry a `ref` later steps can name.
 */
export const DOCX_ALL_OPS = Object.freeze(['insertParagraph', 'setText', 'replaceText', 'formatParagraph', 'deleteParagraph', 'insertTable', 'setCell', 'insertPageBreak']);
/**
 * Implemented but not offered: a table the Writer API inserts shows in the editor, but office-kit's Kit then cannot save
 * the document (it answers the save with a failure; office-kit request K14, like PPT's K6). Offer them once Kit saves them.
 */
export const DOCX_HELD_OPS = Object.freeze(['insertTable', 'setCell']);
/** The operations the model may send. */
export const DOCX_OPS = Object.freeze(DOCX_ALL_OPS.filter(op => !DOCX_HELD_OPS.includes(op)));
export const DOCX_LIMITS = Object.freeze({
  maxOps: 100, maxText: 20_000, maxParagraphText: 5_000,
  maxTableRows: 30, maxTableColumns: 12, maxCellText: 500,
  maxSummaryChars: 60_000, maxParagraphSummary: 2_000, maxTableSummaryCells: 400,
});
/** `style` values and the built-in paragraph style each applies. */
export const DOCX_STYLES = Object.freeze({ title: 'Title', subtitle: 'Subtitle', heading1: 'Heading1', heading2: 'Heading2', heading3: 'Heading3', body: 'Normal', quote: 'Quote' });
/**
 * Styles that usually carry an outline level, which Kit cannot save once the Writer API applied it (office-kit request
 * K15). They are applied only in documents marked with DOCX_LIVE_MARK (shimo-docx's templates, built without outline
 * levels); anywhere else they get DOCX_STYLE_FALLBACK's direct formatting.
 */
export const DOCX_OUTLINE_STYLES = Object.freeze(['title', 'heading1', 'heading2', 'heading3']);
export const DOCX_LIVE_MARK = 'Shimo Live';
/**
 * Direct formatting used when the document has no style with that id (Chinese Word files name theirs "1", "a3"…;
 * the Writer API applies styles by id). Sizes in points. No outline level: Kit cannot save a document after the Writer
 * API set one (office-kit request K15), so headings get their look only.
 */
export const DOCX_STYLE_FALLBACK = Object.freeze({
  title: { size: 22, bold: true, align: 'center' },
  subtitle: { size: 16, align: 'center' },
  heading1: { size: 16, bold: true },
  heading2: { size: 15, bold: true },
  heading3: { size: 14, bold: true },
  body: { bold: false },
  quote: { italic: true },
});
export const DOCX_ALIGN = Object.freeze({ left: 'Left', center: 'Centered', right: 'Right', justify: 'Justified' });
export const DOCX_LISTS = Object.freeze(['bullet', 'number']);
const FORMAT_KEYS = ['align', 'bold', 'italic', 'underline', 'size', 'color', 'font', 'firstLineIndent', 'leftIndent', 'lineSpacing', 'spaceBefore', 'spaceAfter'];
const POINTS = { firstLineIndent: [-1000, 1000], leftIndent: [0, 1000], lineSpacing: [1, 400], spaceBefore: [0, 1000], spaceAfter: [0, 1000] };
const COLOR = /^#[0-9A-Fa-f]{6}$/;
const REF = /^[A-Za-z][A-Za-z0-9_-]{0,31}$/;
/** Operations that name what document_read returned, so they need its readToken. */
const POSITIONED = ['after', 'before'];

const invalid = (index, message) => { throw new SuiteError('invalid-edit', `Operation ${index + 1}: ${message}`, { index }); };
const integer = (value, index, field, min, max) => {
  if (!Number.isInteger(value) || value < min || value > max) invalid(index, `${field} must be a whole number between ${min} and ${max}.`);
  return value;
};

/**
 * A short digest of what the model read: every paragraph's text and the table count. It changes when anyone edits
 * the text, so a batch built on an older read is turned away instead of landing on the wrong paragraphs.
 */
export function docxReadToken(doc) {
  const paragraphs = Array.isArray(doc?.paragraphs) ? doc.paragraphs : [];
  const source = `${paragraphs.length}\u0001${(doc?.tables ?? []).length}\u0001${paragraphs.map(paragraph => paragraph?.text ?? '').join('\u0002')}`;
  // FNV-1a, 2 × 32 bits with different seeds: deterministic in the page and in Node, no crypto needed.
  let a = 0x811c9dc5, b = 0x01000193 ^ 0x5bd1e995;
  for (let i = 0; i < source.length; i++) {
    const code = source.charCodeAt(i);
    a = Math.imul(a ^ code, 0x01000193) >>> 0;
    b = Math.imul(b ^ code, 0x01000197) >>> 0;
  }
  return `d${a.toString(16).padStart(8, '0')}${b.toString(16).padStart(8, '0')}`;
}

/** Whether a batch names paragraphs or tables of the read (and so needs its readToken). */
export function docxNeedsToken(ops) {
  return (Array.isArray(ops) ? ops : []).some(op => op && typeof op === 'object' && (op.paragraph !== undefined || op.table !== undefined || POSITIONED.some(key => op[key] !== undefined)));
}

/** Throws `source-changed` when the batch names the read's paragraphs and the document is no longer that read. */
export function checkDocxToken(ops, doc, readToken) {
  if (!docxNeedsToken(ops)) return;
  if (typeof readToken !== 'string' || !readToken) throw new SuiteError('invalid-edit', 'This edit names paragraphs or tables by number: pass the readToken from document_read (or from the last document_edit).');
  if (readToken !== docxReadToken(doc)) throw new SuiteError('source-changed', 'The document changed since it was read (the user may have edited it). Read it again with document_read and build the edit on that.');
}

function formatOf(op, index) {
  const format = {};
  for (const key of FORMAT_KEYS) if (op[key] !== undefined) format[key] = op[key];
  if ('align' in format && !Object.hasOwn(DOCX_ALIGN, format.align)) invalid(index, `align must be one of ${Object.keys(DOCX_ALIGN).join(', ')}.`);
  for (const key of ['bold', 'italic', 'underline']) if (key in format && typeof format[key] !== 'boolean') invalid(index, `${key} must be true or false.`);
  if ('size' in format && (typeof format.size !== 'number' || !(format.size >= 1 && format.size <= 400))) invalid(index, 'size must be a number of points between 1 and 400.');
  if ('color' in format && !COLOR.test(format.color ?? '')) invalid(index, 'color must look like #RRGGBB.');
  if ('font' in format && (typeof format.font !== 'string' || !format.font || format.font.length > 64)) invalid(index, 'font must be a font name.');
  for (const [key, [min, max]] of Object.entries(POINTS)) {
    if (key in format && (typeof format[key] !== 'number' || !Number.isFinite(format[key]) || format[key] < min || format[key] > max)) invalid(index, `${key} must be a number of points between ${min} and ${max}.`);
  }
  return format;
}

/**
 * Checks a batch against the document the user is looking at and resolves each step. `doc` is the page's `inspect`
 * result: `{ paragraphs: [{ text, style?, list?, table? }], tables: [{ rows, columns, values }] }`. Targets come out as
 * `{ paragraph: <number in doc> }` / `{ ref }` (tables: `{ table }` / `{ ref }`); positions as `{ at: 'start'|'end' }`
 * or `{ after|before: target }`.
 */
export function validateDocxOps(ops, doc, { enabled = DOCX_OPS } = {}) {
  if (!Array.isArray(ops) || ops.length === 0) throw new SuiteError('invalid-edit', 'Provide at least one operation.');
  if (ops.length > DOCX_LIMITS.maxOps) throw new SuiteError('invalid-edit', `At most ${DOCX_LIMITS.maxOps} operations per edit; split the edit.`);
  const paragraphs = Array.isArray(doc?.paragraphs) ? doc.paragraphs : [];
  const tables = Array.isArray(doc?.tables) ? doc.tables : [];
  // What each addressable paragraph holds as the batch goes; deleted ones are gone from the map.
  const live = new Map(paragraphs.map((paragraph, number) => [`#${number}`, { text: String(paragraph?.text ?? ''), table: paragraph?.table === true }]));
  const liveTables = new Map(tables.map((table, number) => [`#${number}`, { rows: table.rows, columns: table.columns }]));
  const refs = new Set();
  let total = 0;
  const count = (value, index) => {
    total += value.length;
    if (total > DOCX_LIMITS.maxText) invalid(index, `the edit writes more than ${DOCX_LIMITS.maxText} characters; split it.`);
  };
  const text = (value, index, field, { max = DOCX_LIMITS.maxParagraphText, empty = true } = {}) => {
    if (typeof value !== 'string') invalid(index, `${field} must be a string.`);
    if (/[\r\n\u000b\u2028\u2029]/.test(value)) invalid(index, `${field} must be one paragraph: no line breaks (use one operation per paragraph).`);
    if (!empty && !value) invalid(index, `${field} must not be empty.`);
    if (value.length > max) invalid(index, `${field} is longer than ${max} characters.`);
    count(value, index);
    return value;
  };
  const claim = (op, index) => {
    if (op.ref === undefined) return undefined;
    if (typeof op.ref !== 'string' || !REF.test(op.ref)) invalid(index, 'ref must start with a letter and use letters, digits, - or _ (at most 32).');
    if (refs.has(op.ref)) invalid(index, `ref ${JSON.stringify(op.ref)} is already used in this edit.`);
    refs.add(op.ref);
    return op.ref;
  };
  const paragraphTarget = (op, index, number = 'paragraph', ref = 'paragraphRef') => {
    if (op[number] !== undefined && op[ref] !== undefined) invalid(index, `give ${number} or ${ref}, not both.`);
    if (op[ref] !== undefined) {
      if (typeof op[ref] !== 'string' || !live.has(`@${op[ref]}`)) invalid(index, `${ref} ${JSON.stringify(op[ref])} is not a paragraph added earlier in this edit (or it was deleted).`);
      return { key: `@${op[ref]}`, target: { ref: op[ref] } };
    }
    if (op[number] === undefined) invalid(index, `give ${number} (a number from document_read) or ${ref}.`);
    integer(op[number], index, number, 0, Number.MAX_SAFE_INTEGER);
    if (op[number] >= paragraphs.length) invalid(index, `paragraph ${op[number]} is not in the document.`);
    if (!live.has(`#${op[number]}`)) invalid(index, `paragraph ${op[number]} was deleted earlier in this edit.`);
    return { key: `#${op[number]}`, target: { paragraph: op[number] } };
  };
  const position = (op, index, { table = false } = {}) => {
    const given = ['at', 'after', 'before', 'afterRef', 'beforeRef'].filter(key => op[key] !== undefined);
    if (given.length > 1) invalid(index, `give only one of at, after, before, afterRef, beforeRef (got ${given.join(', ')}).`);
    if (!given.length) return { at: 'end' };
    const [key] = given;
    if (key === 'at') {
      if (op.at !== 'start' && op.at !== 'end') invalid(index, 'at must be "start" or "end".');
      // The document starts with a table: "start" would land inside its first cell.
      if (op.at === 'start' && paragraphs[0]?.table) invalid(index, 'the document starts with a table; insert after its first paragraph outside the table instead.');
      return { at: op.at };
    }
    const side = key.startsWith('after') ? 'after' : 'before';
    const { key: anchor, target } = paragraphTarget(op, index, side, `${side}Ref`);
    if (live.get(anchor).table) invalid(index, `paragraph ${op[side] ?? `@${op[`${side}Ref`]}`} is inside a table; insert next to a paragraph outside tables.`);
    // Two tables must not touch: the editor would number them as one.
    if (table && target.paragraph !== undefined && paragraphs[target.paragraph + (side === 'after' ? 1 : -1)]?.table) invalid(index, `a table there would touch the table next to paragraph ${target.paragraph}; put a paragraph between them.`);
    return { [side]: target };
  };
  const tableTarget = (op, index) => {
    if (op.table !== undefined && op.tableRef !== undefined) invalid(index, 'give table or tableRef, not both.');
    if (op.tableRef !== undefined) {
      if (typeof op.tableRef !== 'string' || !liveTables.has(`@${op.tableRef}`)) invalid(index, `tableRef ${JSON.stringify(op.tableRef)} is not a table added earlier in this edit.`);
      return { table: liveTables.get(`@${op.tableRef}`), target: { ref: op.tableRef } };
    }
    if (op.table === undefined) invalid(index, 'give table (a number from document_read) or tableRef.');
    integer(op.table, index, 'table', 0, Number.MAX_SAFE_INTEGER);
    if (!liveTables.has(`#${op.table}`)) invalid(index, `table ${op.table} is not in the document.`);
    return { table: liveTables.get(`#${op.table}`), target: { table: op.table } };
  };
  const styleOf = (op, index) => {
    if (op.style === undefined) return undefined;
    if (!Object.hasOwn(DOCX_STYLES, op.style)) invalid(index, `style must be one of ${Object.keys(DOCX_STYLES).join(', ')}.`);
    return op.style;
  };
  let danger = false;
  const steps = ops.map((op, index) => {
    if (!op || typeof op !== 'object' || !DOCX_ALL_OPS.includes(op.op)) invalid(index, `unknown operation ${JSON.stringify(op?.op)}; the operations are ${enabled.join(', ')}.`);
    if (!enabled.includes(op.op)) invalid(index, `${op.op} is not available yet: the document could not be saved with it. Tell the user plainly instead of approximating it (a table can be made with the offline script; then say in one sentence why the document was not written in the editor).`);
    switch (op.op) {
      case 'insertParagraph': {
        const where = position(op, index);
        const step = { op: 'insertParagraph', where, text: text(op.text, index, 'text') };
        const ref = claim(op, index);
        if (ref) step.ref = ref;
        const style = styleOf(op, index);
        if (style) step.style = style;
        if (op.list !== undefined) { if (!DOCX_LISTS.includes(op.list)) invalid(index, `list must be one of ${DOCX_LISTS.join(', ')}.`); step.list = op.list; }
        const format = formatOf(op, index);
        if (Object.keys(format).length) step.format = format;
        if (ref) live.set(`@${ref}`, { text: step.text, table: false });
        return step;
      }
      case 'setText': {
        const { key, target } = paragraphTarget(op, index);
        const step = { op: 'setText', target, before: live.get(key).text, text: text(op.text, index, 'text') };
        live.get(key).text = step.text;
        return step;
      }
      case 'replaceText': {
        const { key, target } = paragraphTarget(op, index);
        const find = text(op.find, index, 'find', { empty: false }), replace = text(op.replace, index, 'replace');
        const current = live.get(key).text;
        const at = current.indexOf(find);
        const name = target.ref !== undefined ? `@${target.ref}` : target.paragraph;
        if (at < 0) invalid(index, `${JSON.stringify(find)} does not occur in paragraph ${name}.`);
        if (current.includes(find, at + find.length)) invalid(index, `${JSON.stringify(find)} occurs more than once in paragraph ${name}; quote more context.`);
        const next = current.slice(0, at) + replace + current.slice(at + find.length);
        if (next.length > DOCX_LIMITS.maxParagraphText) invalid(index, 'the resulting paragraph is too long.');
        live.get(key).text = next;
        return { op: 'setText', target, before: current, text: next, find, replace };
      }
      case 'formatParagraph': {
        const { target } = paragraphTarget(op, index);
        const step = { op: 'formatParagraph', target };
        const style = styleOf(op, index);
        if (style) step.style = style;
        if (op.list !== undefined) { if (!DOCX_LISTS.includes(op.list)) invalid(index, `list must be one of ${DOCX_LISTS.join(', ')}.`); step.list = op.list; }
        const format = formatOf(op, index);
        if (Object.keys(format).length) step.format = format;
        if (!step.style && !step.format && !step.list) invalid(index, 'set at least one of style, list or a format field.');
        return step;
      }
      case 'deleteParagraph': {
        const { key, target } = paragraphTarget(op, index);
        if (live.get(key).table) invalid(index, 'that paragraph is inside a table; change the cell with setCell instead.');
        if ([...live.values()].length <= 1) invalid(index, 'a document keeps at least one paragraph; change its text with setText instead.');
        const step = { op: 'deleteParagraph', target, before: live.get(key).text };
        live.delete(key);
        danger = true;
        return step;
      }
      case 'insertTable': {
        const where = position(op, index, { table: true });
        const values = op.values;
        if (!Array.isArray(values) || !values.length || values.length > DOCX_LIMITS.maxTableRows) invalid(index, `values must be 1–${DOCX_LIMITS.maxTableRows} rows.`);
        const columns = Array.isArray(values[0]) ? values[0].length : 0;
        if (columns < 1 || columns > DOCX_LIMITS.maxTableColumns) invalid(index, `each row must have 1–${DOCX_LIMITS.maxTableColumns} cells.`);
        const rows = values.map((row, number) => {
          if (!Array.isArray(row) || row.length !== columns) invalid(index, `row ${number + 1} must have ${columns} cells like the first row.`);
          return row.map(cell => {
            const value = typeof cell === 'number' && Number.isFinite(cell) ? String(cell) : cell;
            return text(value ?? '', index, 'cell', { max: DOCX_LIMITS.maxCellText });
          });
        });
        if (op.header !== undefined && typeof op.header !== 'boolean') invalid(index, 'header must be true or false.');
        const step = { op: 'insertTable', where, values: rows, ...(op.header === true && { header: true }) };
        const ref = claim(op, index);
        if (ref) { step.ref = ref; liveTables.set(`@${ref}`, { rows: rows.length, columns }); }
        return step;
      }
      case 'setCell': {
        const { table, target } = tableTarget(op, index);
        integer(op.row, index, 'row', 0, Math.max(0, table.rows - 1));
        integer(op.column, index, 'column', 0, Math.max(0, table.columns - 1));
        return { op: 'setCell', table: target, row: op.row, column: op.column, text: text(op.text, index, 'text', { max: DOCX_LIMITS.maxCellText }) };
      }
      case 'insertPageBreak':
        return { op: 'insertPageBreak', where: position(op, index) };
    }
    return undefined;
  });
  const deletions = steps.filter(step => step.op === 'deleteParagraph').length;
  return { steps, danger, reasons: danger ? [`deletes ${deletions} paragraph${deletions === 1 ? '' : 's'}`] : [] };
}

const clip = (value, size = 80) => typeof value === 'string' && value.length > size ? `${value.slice(0, size - 1)}…` : value;
const targetLabel = target => target?.ref !== undefined ? `@${target.ref}` : Number.isInteger(target?.paragraph) ? `¶${target.paragraph + 1}` : Number.isInteger(target?.table) ? `T${target.table + 1}` : '';
const whereLabel = where => where?.at ?? (where?.after ? `after ${targetLabel(where.after)}` : where?.before ? `before ${targetLabel(where.before)}` : '');
const formatText = format => Object.entries(format ?? {}).map(([key, value]) => `${key}=${value}`).join(', ');

/** Rows for the approval card: what, where (¶n is the paragraph's 1-based place in the read), and its content. */
export function describeDocxSteps(steps) {
  return steps.map(step => {
    switch (step.op) {
      case 'insertParagraph': return { op: step.op, where: [whereLabel(step.where), step.ref ? `@${step.ref}` : ''].filter(Boolean).join(' '), after: clip([step.style, step.list, step.text].filter(Boolean).join(' · ')) };
      case 'setText': return { op: step.find !== undefined ? 'replaceText' : 'setText', where: targetLabel(step.target), before: clip(step.find ?? step.before), after: clip(step.replace ?? step.text) };
      case 'formatParagraph': return { op: step.op, where: targetLabel(step.target), after: clip([step.style, step.list, formatText(step.format)].filter(Boolean).join(' · ')) };
      case 'deleteParagraph': return { op: step.op, where: targetLabel(step.target), before: clip(step.before) };
      case 'insertTable': return { op: step.op, where: whereLabel(step.where), after: clip(`${step.values.length}×${step.values[0].length}: ${step.values[0].join(' | ')}`) };
      case 'setCell': return { op: step.op, where: `${targetLabel(step.table)} R${step.row + 1}C${step.column + 1}`, after: clip(step.text) };
      case 'insertPageBreak': return { op: step.op, where: whereLabel(step.where) };
    }
    return { op: step.op, where: '' };
  });
}

/** What the model reads: numbered paragraphs (text, style, list, inside a table), tables with their cells, readToken. */
export function summarizeDocument(doc, { maxChars = DOCX_LIMITS.maxSummaryChars, maxParagraph = DOCX_LIMITS.maxParagraphSummary } = {}) {
  let budget = maxChars, truncated = false;
  const take = value => {
    let text = String(value ?? '');
    if (text.length > maxParagraph) { text = text.slice(0, maxParagraph); truncated = true; }
    if (text.length > budget) { text = text.slice(0, Math.max(0, budget)); truncated = true; }
    budget -= text.length;
    return text;
  };
  const paragraphs = (doc?.paragraphs ?? []).map((paragraph, index) => ({
    index, text: take(paragraph?.text),
    ...(paragraph?.style && paragraph.style !== 'Normal' && { style: paragraph.style }),
    ...(paragraph?.list && { list: true }),
    ...(paragraph?.table && { table: true }),
  }));
  const tables = (doc?.tables ?? []).map((table, index) => {
    let cells = 0;
    const values = (table?.values ?? []).map(row => row.map(cell => {
      if (++cells > DOCX_LIMITS.maxTableSummaryCells) { truncated = true; return ''; }
      return take(cell);
    }));
    return { index, rows: table?.rows ?? values.length, columns: table?.columns ?? values[0]?.length ?? 0, values };
  });
  return { paragraphCount: paragraphs.length, paragraphs, tables, readToken: docxReadToken(doc), ...(truncated && { truncated: true }) };
}

/** The progress sentence every failed or stopped edit ends with: there is no rollback. */
export function docxProgress(done, total, refs = {}, partial, extra = []) {
  const finished = done > 0 ? `Steps 1–${done} of ${total} were applied` : `None of the ${total} steps was applied`;
  const kept = done > 0 ? ' and stay in the editor: nothing was undone and nothing was saved.' : '.';
  // A step of several editor calls (or a paragraph written in parts) can stop part way: what it wrote is there.
  const halves = (Array.isArray(partial) ? partial : Number.isInteger(partial) ? [partial] : []).map(number => number + 1);
  const half = halves.length ? ` ${halves.length === 1 ? `Step ${halves[0]} was` : `Steps ${halves.join(', ')} were`} partly applied: what was written is in the document; check it with document_read.` : '';
  const also = extra.length ? ` Step${extra.length === 1 ? '' : 's'} ${extra.map(number => number + 1).join(', ')} also finished.` : '';
  const named = Object.keys(refs).length ? ` refs: ${JSON.stringify(refs)}` : '';
  return `${finished}${kept}${half}${also}${named}`;
}
