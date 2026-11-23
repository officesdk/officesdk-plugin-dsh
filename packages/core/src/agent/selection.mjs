import { SuiteError } from '../errors.mjs';
import { DOCX_LIMITS } from './docx.mjs';
import { XLSX_LIMITS, columnLetters } from './xlsx.mjs';

/**
 * Selection rewrite (spec 261109-selection-ai-edit §3–§5): the user selects text in a Word document or a block of cells
 * in a spreadsheet, writes what to change, and the model answers with the replacement only (selection_rewrite). Nothing
 * here touches an editor: these are the checks and plans the Host and the editor page share.
 *
 * Word anchors use the paragraph numbers of document_read (`body.paragraphs`, table paragraphs included) and offsets
 * in Unicode code points inside the first and last paragraph — the unit of the Writer API's `getSubrange` (measured on
 * office-kit 0.1.2: an emoji counts as one).
 */
export const SELECTION_LIMITS = Object.freeze({
  maxText: 2_000, maxParagraphs: 20, maxCells: 200, maxCellText: 2_000,
  maxInstruction: 500, maxRewrite: 4_000, maxRewriteParagraphs: 50,
});
export const SELECTION_KINDS = Object.freeze(['docx', 'xlsx']);
export const SELECTION_ID = /^sel-[0-9a-f]{8}$/;

const invalid = message => { throw new SuiteError('invalid-edit', message); };
const stale = message => { throw new SuiteError('source-changed', message); };
const whole = (value, min = 0, max = Number.MAX_SAFE_INTEGER) => Number.isInteger(value) && value >= min && value <= max;

/** Length in code points. */
export const cpLength = text => Array.from(String(text ?? '')).length;
/** `text` cut at code points [from, to). */
export const cpSlice = (text, from, to) => Array.from(String(text ?? '')).slice(from, to).join('');
/** One paragraph per line; Word reports paragraph breaks as `\r`. */
export const lines = text => String(text ?? '').replace(/\r\n?/g, '\n').split('\n');

/** The selected text of `paragraphs` (strings) under `anchor`, paragraphs joined with `\n`. */
export function selectedText(paragraphs, { paragraphs: [first, last], start, end }) {
  if (first === last) return cpSlice(paragraphs[first], start, end);
  return [cpSlice(paragraphs[first], start), ...paragraphs.slice(first + 1, last), cpSlice(paragraphs[last], 0, end)].join('\n');
}

function checkDocAnchor(anchor, { maxParagraphs = SELECTION_LIMITS.maxParagraphs } = {}) {
  if (!anchor || typeof anchor !== 'object' || !Array.isArray(anchor.paragraphs) || anchor.paragraphs.length !== 2) invalid('anchor.paragraphs must be [first, last].');
  const [first, last] = anchor.paragraphs;
  if (!whole(first) || !whole(last) || last < first) invalid('anchor.paragraphs must be whole numbers with first ≤ last.');
  if (last - first + 1 > maxParagraphs) invalid(`A selection spans at most ${maxParagraphs} paragraphs.`);
  if (!whole(anchor.start) || !whole(anchor.end)) invalid('anchor.start and anchor.end must be whole numbers.');
  if (first === last && anchor.end < anchor.start) invalid('anchor.end must not be before anchor.start.');
  return { paragraphs: [first, last], start: anchor.start, end: anchor.end };
}

function checkSheetAnchor(anchor) {
  if (!anchor || typeof anchor !== 'object') invalid('anchor is required.');
  const { sheet, row, column, rowCount, columnCount } = anchor;
  if (typeof sheet !== 'string' || !sheet || sheet.length > 31) invalid('anchor.sheet must be the worksheet name.');
  if (!whole(row) || !whole(column) || !whole(rowCount, 1) || !whole(columnCount, 1)) invalid('anchor must name a block of cells.');
  if (rowCount * columnCount > SELECTION_LIMITS.maxCells) invalid(`A selection holds at most ${SELECTION_LIMITS.maxCells} cells.`);
  return { sheet, row, column, rowCount, columnCount };
}

/** One cell as the page reads it: shown text, formula (without "=") and the raw value (string, number, boolean or null). */
function checkCell(cell) {
  if (!cell || typeof cell !== 'object') invalid('Each cell must be an object.');
  const text = typeof cell.text === 'string' ? cell.text : '';
  if (text.length > SELECTION_LIMITS.maxCellText) invalid('A cell holds too much text for a selection rewrite.');
  const raw = cell.raw === undefined ? null : cell.raw;
  if (raw !== null && !['string', 'number', 'boolean'].includes(typeof raw)) invalid('A cell value must be text, a number, true/false or empty.');
  if (typeof raw === 'number' && !Number.isFinite(raw)) invalid('A cell value must be a finite number.');
  return { text, raw, ...(typeof cell.formula === 'string' && cell.formula && { formula: cell.formula.replace(/^=/, '') }), ...(cell.date === true && typeof raw === 'number' && { date: true }) };
}

/**
 * Checks what the page sends when the user asks for a rewrite and returns it normalized:
 * `{ kind, anchor, original, cells?, instruction }`.
 */
export function validateSelectionRequest(request) {
  if (!request || typeof request !== 'object') invalid('A selection request is required.');
  const { kind } = request;
  if (!SELECTION_KINDS.includes(kind)) invalid(`kind must be one of ${SELECTION_KINDS.join(', ')}.`);
  const instruction = typeof request.instruction === 'string' ? request.instruction.trim() : '';
  if (!instruction) invalid('Write what to change.');
  if (instruction.length > SELECTION_LIMITS.maxInstruction) invalid(`The request is longer than ${SELECTION_LIMITS.maxInstruction} characters.`);
  if (kind === 'docx') {
    const anchor = checkDocAnchor(request.anchor);
    if (typeof request.original !== 'string' || !request.original.trim()) invalid('The selection is empty.');
    if (/\r/.test(request.original)) invalid('original separates paragraphs with \\n.');
    if (cpLength(request.original) > SELECTION_LIMITS.maxText) invalid(`A selection holds at most ${SELECTION_LIMITS.maxText} characters.`);
    if (lines(request.original).length !== anchor.paragraphs[1] - anchor.paragraphs[0] + 1) invalid('original does not match the paragraphs of the anchor.');
    return { kind, anchor, original: request.original, instruction };
  }
  const anchor = checkSheetAnchor(request.anchor);
  const cells = request.cells;
  if (!Array.isArray(cells) || cells.length !== anchor.rowCount || cells.some(row => !Array.isArray(row) || row.length !== anchor.columnCount)) invalid('cells must match the size of the selection.');
  // Empty cells are fine too ("fill in a header row").
  const normalized = cells.map(row => row.map(checkCell));
  return { kind, anchor, cells: normalized, original: sheetOriginalText(normalized), instruction };
}

/** "B2:D5" for an anchor. */
export function sheetAddress({ row, column, rowCount, columnCount }) {
  const from = `${columnLetters(column)}${row + 1}`;
  return rowCount === 1 && columnCount === 1 ? from : `${from}:${columnLetters(column + columnCount - 1)}${row + rowCount}`;
}

const cellShown = cell => cell.formula ? `=${cell.formula}` : cell.text || (cell.raw === null ? '' : String(cell.raw));
function sheetOriginalText(cells) { return cells.map(row => row.map(cellShown).join('\t')).join('\n'); }

const quote = text => lines(text).map(line => `> ${line}`).join('\n');
const tableCell = text => String(text).replace(/\|/g, '\\|').replace(/\n/g, ' ');

/**
 * The user's message that carries the request into the conversation (spec §3.2): visible, complete (the model writes
 * the replacement from it) and naming the selection id selection_rewrite takes.
 */
export function selectionMessage({ id, name, request, lang = 'zh' }) {
  const en = lang === 'en';
  if (request.kind === 'docx') {
    const [first, last] = request.anchor.paragraphs;
    const where = first === last ? (en ? `paragraph ${first + 1}` : `第 ${first + 1} 段`) : (en ? `paragraphs ${first + 1}–${last + 1}` : `第 ${first + 1}–${last + 1} 段`);
    return en
      ? `Rewrite the text I selected in "${name}" (selection ${id}, ${where}):\n${quote(request.original)}\nRequest: ${request.instruction}\n(Answer with selection_rewrite.)`
      : `修改「${name}」里我选中的文字（选区 ${id}，${where}）：\n${quote(request.original)}\n要求：${request.instruction}\n（用 selection_rewrite 修改）`;
  }
  const { anchor, cells } = request;
  const address = `${anchor.sheet}!${sheetAddress(anchor)}`;
  const header = `| ${Array.from({ length: anchor.columnCount }, (_, c) => columnLetters(anchor.column + c)).join(' | ')} |`;
  const rule = `|${' --- |'.repeat(anchor.columnCount)}`;
  const rows = cells.map(row => `| ${row.map(cell => tableCell(cellShown(cell))).join(' | ')} |`);
  const table = [header, rule, ...rows].join('\n');
  return en
    ? `Change the cells I selected in "${name}" ${address} (selection ${id}, ${anchor.rowCount}×${anchor.columnCount}):\n${table}\nRequest: ${request.instruction}\n(Answer with selection_rewrite.)`
    : `修改「${name}」${address} 我选中的单元格（选区 ${id}，${anchor.rowCount} 行 × ${anchor.columnCount} 列）：\n${table}\n要求：${request.instruction}\n（用 selection_rewrite 修改）`;
}

/** Checks the model's replacement text for a Word selection. */
export function checkRewriteText(text) {
  if (typeof text !== 'string') invalid('text is required: the new text for the selection (paragraphs separated by \\n).');
  if (cpLength(text) > SELECTION_LIMITS.maxRewrite) invalid(`text is longer than ${SELECTION_LIMITS.maxRewrite} characters.`);
  const rows = lines(text);
  if (rows.length > SELECTION_LIMITS.maxRewriteParagraphs) invalid(`text has more than ${SELECTION_LIMITS.maxRewriteParagraphs} paragraphs.`);
  if (rows.some(row => row.length > DOCX_LIMITS.maxParagraphText)) invalid(`A paragraph of text is longer than ${DOCX_LIMITS.maxParagraphText} characters.`);
  return rows.join('\n');
}

/** Checks the model's replacement values for a spreadsheet selection of `anchor`'s size. */
export function checkRewriteValues(values, anchor) {
  if (!Array.isArray(values) || values.length !== anchor.rowCount || values.some(row => !Array.isArray(row) || row.length !== anchor.columnCount)) {
    invalid(`values must be ${anchor.rowCount} row${anchor.rowCount === 1 ? '' : 's'} of ${anchor.columnCount} value${anchor.columnCount === 1 ? '' : 's'}, the size of the selection.`);
  }
  let total = 0;
  return values.map(row => row.map(value => {
    if (value === null || value === undefined) return null;
    if (typeof value === 'number') { if (!Number.isFinite(value)) invalid('Numbers must be finite.'); return value; }
    if (typeof value === 'boolean') return value;
    if (typeof value !== 'string') invalid('Each value must be text, a number, true/false or null.');
    if (value.length > SELECTION_LIMITS.maxCellText) invalid(`A value is longer than ${SELECTION_LIMITS.maxCellText} characters.`);
    if (value.startsWith('=') && (value.length < 2 || value.startsWith('=='))) invalid(`${JSON.stringify(value)} is not a formula; text starting with "=" is written as a formula (for example "=SUM(B2:B5)").`);
    total += value.length;
    if (total > XLSX_LIMITS.maxText) invalid(`values hold more than ${XLSX_LIMITS.maxText} characters.`);
    return value;
  }));
}

/**
 * How to replace the selection under `anchor` in the document now (`paragraphs`: `{ text, table? }` as inspectDocument
 * reads them) with `text`. The selected text must still be `original` (else `source-changed`). Returns
 * - `{ mode: 'inline', paragraph, start, length, text }` — one paragraph, no new paragraphs: the Writer subrange is
 *   replaced and the rest of the paragraph keeps its formatting;
 * - `{ mode: 'paragraphs', ops }` — document_edit operations (setText / insertParagraph / deleteParagraph on the
 *   selected paragraphs only) for runDocxSteps;
 * - `{ mode: 'unchanged' }` when the replacement is what is there already;
 * and either way `after` / `afterText`: where the new text is once written, for revert.
 */
export function planDocRewrite(paragraphs, anchor, original, text, { revert = false } = {}) {
  const { paragraphs: [first, last], start, end } = checkDocAnchor(anchor, revert ? { maxParagraphs: SELECTION_LIMITS.maxRewriteParagraphs } : undefined);
  const texts = paragraphs.map(paragraph => String(paragraph?.text ?? ''));
  if (last >= texts.length) stale('The selected paragraphs are no longer in the document.');
  if (paragraphs.slice(first, last + 1).some(paragraph => paragraph?.table)) invalid('The selection is inside a table; selection rewrite works on body text only.');
  if (start > cpLength(texts[first]) || end > cpLength(texts[last])) stale('The selected text is no longer where it was.');
  if (selectedText(texts, { paragraphs: [first, last], start, end }) !== original) stale('The selected text changed since it was selected (the user may have edited it).');
  const replacement = checkRewriteText(text);
  const rows = lines(replacement);
  if (first === last && rows.length === 1) {
    return { mode: 'inline', paragraph: first, start, length: end - start, text: replacement,
      after: { paragraphs: [first, first], start, end: start + cpLength(replacement) }, afterText: replacement };
  }
  const prefix = cpSlice(texts[first], 0, start), suffix = cpSlice(texts[last], end);
  const full = rows.map((row, index) => `${index === 0 ? prefix : ''}${row}${index === rows.length - 1 ? suffix : ''}`);
  const old = last - first + 1;
  const ops = [];
  for (let index = 0; index < Math.min(old, full.length); index++) {
    if (full[index] !== texts[first + index]) ops.push({ op: 'setText', paragraph: first + index, text: full[index] });
  }
  for (let index = old; index < full.length; index++) {
    ops.push({ op: 'insertParagraph', ...(index === old ? { after: last } : { afterRef: `sel${index - 1}` }), ref: `sel${index}`, text: full[index] });
  }
  for (let index = full.length; index < old; index++) ops.push({ op: 'deleteParagraph', paragraph: first + index });
  return { mode: ops.length ? 'paragraphs' : 'unchanged', ops, after: { paragraphs: [first, first + full.length - 1], start, end: cpLength(rows.at(-1)) + (rows.length === 1 ? start : 0) }, afterText: replacement };
}

/** Taking a rewrite back is the same plan the other way: `after` must still hold `afterText`. */
export const planDocRevert = (paragraphs, after, afterText, original) => planDocRewrite(paragraphs, after, afterText, original, { revert: true });

const formulaKey = formula => String(formula).replace(/^=/, '').replace(/\s+/g, '').toUpperCase();
/** A cell compared loosely: its formula, else its value as text (Sheet may report "1" for 1). */
export function cellKey(cell) {
  if (cell?.formula) return `f:${formulaKey(cell.formula)}`;
  const raw = cell?.raw;
  if (raw === null || raw === undefined || raw === '') return `v:${cell?.text ?? ''}`;
  return `v:${String(raw)}`;
}
/** What a written value reads back as. */
export const valueKey = value => typeof value === 'string' && value.startsWith('=') ? `f:${formulaKey(value)}` : `v:${value === null || value === undefined ? '' : String(value)}`;

/** The value that writes `cell` back (formula with "=", else its raw value). */
export const cellValue = cell => cell.formula ? `=${cell.formula}` : cell.raw === undefined ? null : cell.raw === '' ? null : cell.raw;

/**
 * Replacing a spreadsheet selection: `current` are the cells now (as the page reads them), `cells` the ones recorded
 * when the user asked. Returns the setValues step for runXlsxSteps.
 */
export function planSheetRewrite(current, anchor, cells, values) {
  const range = checkSheetAnchor(anchor);
  const written = checkRewriteValues(values, range);
  if (!Array.isArray(current) || current.length !== range.rowCount) stale('The selected cells are no longer there.');
  for (const [r, row] of cells.entries()) for (const [c, cell] of row.entries()) {
    if (cellKey(current[r]?.[c]) !== cellKey(cell)) stale(`Cell ${columnLetters(range.column + c)}${range.row + r + 1} changed since it was selected (the user may have edited it).`);
  }
  return { step: { op: 'setValues', sheet: range.sheet, range: { row: range.row, column: range.column, rowCount: range.rowCount, columnCount: range.columnCount }, values: written } };
}

/** A recorded cell setValues cannot write back as it was: a date, or plain text that starts with "=" (it would become a formula). */
export const specialCell = cell => !cell.formula && (cell.date === true || (typeof cell.raw === 'string' && cell.raw.startsWith('=')));

/**
 * Taking a spreadsheet rewrite back: the cells must still hold what was written; the recorded cells go back. `specials`
 * are the cells the step leaves empty and the caller writes with their own value type (`{ type, value }`).
 */
export function planSheetRevert(current, anchor, written, cells) {
  const range = checkSheetAnchor(anchor);
  for (const [r, row] of written.entries()) for (const [c, value] of row.entries()) {
    if (cellKey(current[r]?.[c]) !== valueKey(value)) stale(`Cell ${columnLetters(range.column + c)}${range.row + r + 1} changed after the AI wrote it.`);
  }
  const specials = cells.flatMap((row, r) => row.flatMap((cell, c) => specialCell(cell) ? [{ r, c, value: { type: cell.date ? 'date' : 'primitive', value: cell.raw } }] : []));
  return {
    step: { op: 'setValues', sheet: range.sheet, range: { row: range.row, column: range.column, rowCount: range.rowCount, columnCount: range.columnCount }, values: cells.map(row => row.map(cell => specialCell(cell) ? null : cellValue(cell))) },
    specials,
  };
}
