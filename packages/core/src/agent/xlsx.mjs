import { SuiteError } from '../errors.mjs';

/**
 * Structured spreadsheet edits (spec 261105 §6). Each operation maps to a fixed call of the Sheet SDK's members that
 * office-kit's editor page makes public in edit mode (`api.sheet`, office-kit request K16, public since 0.3.0-rc.1). There is no rollback: a batch that fails or is stopped keeps the steps it finished.
 * Cells are named in A1 notation on a worksheet named by `sheet` (the active one without it).
 */
export const XLSX_OPS = Object.freeze(['setValues', 'setFormula', 'formatRange', 'clearRange', 'addSheet', 'renameSheet', 'activateSheet', 'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns', 'setColumnWidth', 'freezeRows', 'addChart']);
export const XLSX_LIMITS = Object.freeze({ maxOps: 100, maxCells: 5_000, maxCellText: 2_000, maxText: 50_000, maxSheetName: 31, maxReadCells: 2_000, maxReadText: 60_000 });
export const XLSX_ALIGN = Object.freeze(['left', 'center', 'right']);
const COLOR = /^#[0-9A-Fa-f]{6}$/;
const SHEET_NAME_BAD = /[\\/:?*[\]]/;
const CELL = /^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/;
const COLUMNS = /^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?$/;
const FORMAT_KEYS = ['bold', 'italic', 'underline', 'color', 'background', 'fontSize', 'align', 'wrap', 'numberFormat'];

const invalid = (index, message) => { throw new SuiteError('invalid-edit', `Operation ${index + 1}: ${message}`, { index }); };

/** Column letters → 0-based index (A = 0). */
export function columnIndex(letters) {
  let value = 0;
  for (const char of letters.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64);
  return value - 1;
}
/** 0-based column index → letters. */
export function columnLetters(index) {
  let letters = '';
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + ((value - 1) % 26)) + letters;
  return letters;
}
/** "B2" or "B2:D5" → { row, column, rowCount, columnCount } (0-based), or undefined. */
export function parseRange(text) {
  if (typeof text !== 'string') return undefined;
  const [first, second, extra] = text.trim().split(':');
  if (extra !== undefined) return undefined;
  const a = CELL.exec(first ?? ''), b = second === undefined ? a : CELL.exec(second);
  if (!a || !b) return undefined;
  const rows = [Number(a[2]) - 1, Number(b[2]) - 1], columns = [columnIndex(a[1]), columnIndex(b[1])];
  const row = Math.min(...rows), column = Math.min(...columns);
  return { row, column, rowCount: Math.max(...rows) - row + 1, columnCount: Math.max(...columns) - column + 1 };
}
export const rangeText = ({ row, column, rowCount = 1, columnCount = 1 }) => {
  const start = `${columnLetters(column)}${row + 1}`;
  return rowCount === 1 && columnCount === 1 ? start : `${start}:${columnLetters(column + columnCount - 1)}${row + rowCount}`;
};

/**
 * Checks a batch against the workbook the user is looking at. `book` is the page's `inspect` result:
 * `{ sheets: [{ name, active, rowCount, columnCount }] }`. Steps come out with sheet names resolved and ranges parsed.
 */
export function validateXlsxOps(ops, book) {
  if (!Array.isArray(ops) || ops.length === 0) throw new SuiteError('invalid-edit', 'Provide at least one operation.');
  if (ops.length > XLSX_LIMITS.maxOps) throw new SuiteError('invalid-edit', `At most ${XLSX_LIMITS.maxOps} operations per edit; split the edit.`);
  const sheets = new Map((book?.sheets ?? []).map(sheet => [sheet.name, { rowCount: sheet.rowCount, columnCount: sheet.columnCount }]));
  let active = (book?.sheets ?? []).find(sheet => sheet.active)?.name ?? book?.sheets?.[0]?.name;
  let cells = 0, text = 0, danger = false;
  const reasons = [];
  const sheetOf = (op, index, field = 'sheet') => {
    const name = op[field] ?? active;
    if (typeof name !== 'string' || !sheets.has(name)) invalid(index, `there is no worksheet ${JSON.stringify(op[field])}; the worksheets are ${[...sheets.keys()].join(', ')}.`);
    return name;
  };
  const rangeOf = (op, index, sheet, field = 'range') => {
    const range = parseRange(op[field]);
    if (!range) invalid(index, `${field} must be a cell like "B2" or a range like "B2:D5".`);
    const size = sheets.get(sheet);
    if (size?.rowCount && range.row + range.rowCount > size.rowCount) invalid(index, `${op[field]} goes past row ${size.rowCount}, the last row of ${sheet}; add rows with insertRows first.`);
    if (size?.columnCount && range.column + range.columnCount > size.columnCount) invalid(index, `${op[field]} goes past column ${columnLetters(size.columnCount - 1)}, the last column of ${sheet}; add columns with insertColumns first.`);
    return range;
  };
  const count = (amount, index) => { cells += amount; if (cells > XLSX_LIMITS.maxCells) invalid(index, `the edit touches more than ${XLSX_LIMITS.maxCells} cells; split it.`); };
  const cellValue = (value, index) => {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') { if (!Number.isFinite(value)) invalid(index, 'numbers must be finite.'); return value; }
    if (typeof value === 'boolean') return value;
    if (typeof value !== 'string') invalid(index, 'cell values must be text, numbers, true/false or null.');
    if (value.startsWith('=') && (value.length < 2 || value.startsWith('=='))) invalid(index, `${JSON.stringify(value)} is not a formula; text starting with "=" is written as a formula (for example "=SUM(B2:B5)").`);
    if (value.length > XLSX_LIMITS.maxCellText) invalid(index, `a cell is longer than ${XLSX_LIMITS.maxCellText} characters.`);
    text += value.length;
    if (text > XLSX_LIMITS.maxText) invalid(index, `the edit writes more than ${XLSX_LIMITS.maxText} characters; split it.`);
    return value;
  };
  const sheetName = (value, index, field = 'name', renaming) => {
    if (typeof value !== 'string' || !value.trim() || value.length > XLSX_LIMITS.maxSheetName || SHEET_NAME_BAD.test(value)) invalid(index, `${field} must be 1–${XLSX_LIMITS.maxSheetName} characters without \\ / : ? * [ ].`);
    // Excel compares sheet names without case.
    if ([...sheets.keys()].some(name => name !== renaming && name.toLowerCase() === value.toLowerCase())) invalid(index, `there is already a worksheet named ${JSON.stringify(value)}.`);
    return value;
  };
  const amount = (value, index, field) => { if (!Number.isInteger(value) || value < 1 || value > 1000) invalid(index, `${field} must be a whole number from 1 to 1000.`); return value; };
  const steps = ops.map((op, index) => {
    if (!op || typeof op !== 'object' || !XLSX_OPS.includes(op.op)) invalid(index, `unknown operation ${JSON.stringify(op?.op)}; the operations are ${XLSX_OPS.join(', ')}.`);
    switch (op.op) {
      case 'setValues': {
        const sheet = sheetOf(op, index);
        if (!Array.isArray(op.values) || !op.values.length || !op.values.every(Array.isArray)) invalid(index, 'values must be rows (arrays) of cell values.');
        const width = op.values[0].length;
        if (!width || op.values.some(row => row.length !== width)) invalid(index, 'every row of values must have the same number of cells.');
        const given = parseRange(op.range);
        if (!given) invalid(index, 'range must be a cell like "B2" (the top-left of the values) or a range like "B2:D5".');
        if ((given.rowCount > 1 || given.columnCount > 1) && (given.rowCount !== op.values.length || given.columnCount !== width)) invalid(index, `range ${op.range} is ${given.rowCount}×${given.columnCount} but values are ${op.values.length}×${width}.`);
        const range = rangeOf({ range: rangeText({ ...given, rowCount: op.values.length, columnCount: width }) }, index, sheet);
        count(op.values.length * width, index);
        const values = op.values.map(row => row.map(value => cellValue(value, index)));
        return { op: 'setValues', sheet, range, values };
      }
      case 'setFormula': {
        const sheet = sheetOf(op, index);
        const range = rangeOf({ range: op.cell }, index, sheet);
        if (range.rowCount !== 1 || range.columnCount !== 1) invalid(index, 'cell must be one cell like "E5".');
        if (typeof op.formula !== 'string' || !op.formula.startsWith('=') || op.formula.length < 2 || op.formula.length > XLSX_LIMITS.maxCellText) invalid(index, 'formula must start with = (for example "=SUM(B2:B5)").');
        count(1, index);
        return { op: 'setFormula', sheet, range, formula: op.formula };
      }
      case 'formatRange': {
        const sheet = sheetOf(op, index);
        const range = rangeOf(op, index, sheet);
        const format = {};
        for (const key of FORMAT_KEYS) if (op[key] !== undefined) format[key] = op[key];
        if (!Object.keys(format).length) invalid(index, `set at least one of ${FORMAT_KEYS.join(', ')}.`);
        for (const key of ['bold', 'italic', 'underline', 'wrap']) if (key in format && typeof format[key] !== 'boolean') invalid(index, `${key} must be true or false.`);
        for (const key of ['color', 'background']) if (key in format && !COLOR.test(format[key] ?? '')) invalid(index, `${key} must look like #RRGGBB.`);
        if ('fontSize' in format && (typeof format.fontSize !== 'number' || !(format.fontSize >= 6 && format.fontSize <= 72))) invalid(index, 'fontSize must be a number of points from 6 to 72.');
        if ('align' in format && !XLSX_ALIGN.includes(format.align)) invalid(index, `align must be one of ${XLSX_ALIGN.join(', ')}.`);
        if ('numberFormat' in format && (typeof format.numberFormat !== 'string' || !format.numberFormat || format.numberFormat.length > 64)) invalid(index, 'numberFormat must be a format code such as "0.00", "#,##0" or "0%".');
        count(range.rowCount * range.columnCount, index);
        return { op: 'formatRange', sheet, range, format };
      }
      case 'clearRange': {
        const sheet = sheetOf(op, index);
        const range = rangeOf(op, index, sheet);
        count(range.rowCount * range.columnCount, index);
        danger = true; reasons.push(`clears ${rangeText(range)}`);
        return { op: 'clearRange', sheet, range };
      }
      case 'addSheet': {
        const name = sheetName(op.name, index);
        // A new sheet's size is the SDK's own default, unknown here: its writes are not size-checked. It becomes the
        // active sheet, as the editor shows it.
        sheets.set(name, {});
        active = name;
        return { op: 'addSheet', name };
      }
      case 'renameSheet': {
        const sheet = sheetOf(op, index);
        const name = sheetName(op.name, index, 'name', sheet);
        const size = sheets.get(sheet);
        sheets.delete(sheet); sheets.set(name, size);
        if (active === sheet) active = name;
        return { op: 'renameSheet', sheet, name };
      }
      case 'activateSheet': {
        const sheet = sheetOf(op, index);
        active = sheet;
        return { op: 'activateSheet', sheet };
      }
      case 'insertRows': case 'deleteRows': {
        const sheet = sheetOf(op, index);
        const size = sheets.get(sheet);
        if (!Number.isInteger(op.row) || op.row < 1 || (size?.rowCount && op.row > size.rowCount + (op.op === 'insertRows' ? 1 : 0))) invalid(index, `row must be a row number (1-based) within ${sheet}.`);
        const rows = amount(op.count ?? 1, index, 'count');
        if (op.op === 'deleteRows' && size?.rowCount && (op.row + rows - 1 > size.rowCount || rows >= size.rowCount)) invalid(index, `rows ${op.row}–${op.row + rows - 1} are not all in ${sheet} (it has ${size.rowCount}), or would delete every row.`);
        if (size?.rowCount) size.rowCount += op.op === 'insertRows' ? rows : -rows;
        if (op.op === 'deleteRows') { danger = true; reasons.push(`deletes ${rows} row${rows === 1 ? '' : 's'}`); }
        return { op: op.op, sheet, row: op.row - 1, count: rows };
      }
      case 'insertColumns': case 'deleteColumns': {
        const sheet = sheetOf(op, index);
        const size = sheets.get(sheet);
        const match = typeof op.column === 'string' ? COLUMNS.exec(op.column) : undefined;
        if (!match || match[2]) invalid(index, 'column must be a column letter like "C".');
        const column = columnIndex(match[1]);
        if (size?.columnCount && column > size.columnCount - (op.op === 'insertColumns' ? 0 : 1)) invalid(index, `column ${op.column} is past the last column of ${sheet}.`);
        const columns = amount(op.count ?? 1, index, 'count');
        if (op.op === 'deleteColumns' && size?.columnCount && (column + columns > size.columnCount || columns >= size.columnCount)) invalid(index, `those ${columns} columns are not all in ${sheet}, or would delete every column.`);
        if (size?.columnCount) size.columnCount += op.op === 'insertColumns' ? columns : -columns;
        if (op.op === 'deleteColumns') { danger = true; reasons.push(`deletes ${columns} column${columns === 1 ? '' : 's'}`); }
        return { op: op.op, sheet, column, count: columns };
      }
      case 'setColumnWidth': {
        const sheet = sheetOf(op, index);
        const match = typeof op.columns === 'string' ? COLUMNS.exec(op.columns) : undefined;
        if (!match) invalid(index, 'columns must be a column like "B" or columns like "B:D".');
        const from = columnIndex(match[1]), to = columnIndex(match[2] ?? match[1]);
        if (typeof op.width !== 'number' || !(op.width >= 10 && op.width <= 1000)) invalid(index, 'width must be a number of pixels from 10 to 1000.');
        const size = sheets.get(sheet);
        if (size?.columnCount && Math.max(from, to) >= size.columnCount) invalid(index, `columns ${op.columns} go past the last column of ${sheet}.`);
        return { op: 'setColumnWidth', sheet, column: Math.min(from, to), count: Math.abs(to - from) + 1, width: op.width };
      }
      case 'addChart': {
        // A clustered column chart of the range (headers in its first row and column), placed one blank row below it.
        const sheet = sheetOf(op, index);
        const range = rangeOf(op, index, sheet);
        if (range.rowCount < 2 || range.columnCount < 2) invalid(index, 'range must include a header row and a label column, for example "A1:D9".');
        if (op.title !== undefined && (typeof op.title !== 'string' || op.title.length > 100)) invalid(index, 'title must be text of at most 100 characters.');
        return { op: 'addChart', sheet, range, ...(op.title && { title: op.title }) };
      }
      case 'freezeRows': {
        const sheet = sheetOf(op, index);
        if (!Number.isInteger(op.count) || op.count < 0 || op.count > 100) invalid(index, 'count must be a whole number of rows from 0 to 100.');
        if (sheets.get(sheet)?.rowCount && op.count >= sheets.get(sheet).rowCount) invalid(index, `${sheet} has only ${sheets.get(sheet).rowCount} rows.`);
        return { op: 'freezeRows', sheet, count: op.count };
      }
    }
    return undefined;
  });
  return { steps, danger, reasons };
}

const clip = (value, size = 80) => typeof value === 'string' && value.length > size ? `${value.slice(0, size - 1)}…` : value;

/** Rows for the approval card: what, where (sheet!range), and its content. */
export function describeXlsxSteps(steps) {
  return steps.map(step => {
    const where = step.range ? `${step.sheet}!${rangeText(step.range)}` : step.sheet ?? '';
    switch (step.op) {
      case 'setValues': return { op: step.op, where, after: clip(step.values.slice(0, 3).map(row => row.map(value => value ?? '').join(' | ')).join(' / ')) };
      case 'setFormula': return { op: step.op, where, after: step.formula };
      case 'formatRange': return { op: step.op, where, after: Object.entries(step.format).map(([key, value]) => `${key}=${value}`).join(', ') };
      case 'clearRange': return { op: step.op, where };
      case 'addSheet': return { op: step.op, where: step.name };
      case 'renameSheet': return { op: step.op, where: step.sheet, after: step.name };
      case 'activateSheet': return { op: step.op, where: step.sheet };
      case 'insertRows': case 'deleteRows': return { op: step.op, where: `${step.sheet}!${step.row + 1}`, after: `×${step.count}` };
      case 'insertColumns': case 'deleteColumns': return { op: step.op, where: `${step.sheet}!${columnLetters(step.column)}`, after: `×${step.count}` };
      case 'setColumnWidth': return { op: step.op, where: `${step.sheet}!${columnLetters(step.column)}${step.count > 1 ? `:${columnLetters(step.column + step.count - 1)}` : ''}`, after: `${step.width}px` };
      case 'freezeRows': return { op: step.op, where: step.sheet, after: String(step.count) };
      case 'addChart': return { op: step.op, where, after: step.title ?? '' };
    }
    return { op: step.op, where };
  });
}

/**
 * What the model reads: each worksheet (name, whether active, size, used range) and the used range's cells as rows of
 * shown text, with formulas listed by address. `book.sheets[].cells` is the page's `{ range, rows: [[{ text, formula }]] }`.
 */
export function summarizeWorkbook(book, { maxCells = XLSX_LIMITS.maxReadCells, maxText = XLSX_LIMITS.maxReadText } = {}) {
  let cellBudget = maxCells, textBudget = maxText, truncated = false;
  const sheets = (book?.sheets ?? []).map(sheet => {
    const used = sheet.cells?.used ?? sheet.cells?.range;
    const entry = { name: sheet.name, ...(sheet.active && { active: true }), rows: sheet.rowCount, columns: sheet.columnCount, used: used ? rangeText(used) : null };
    if (!sheet.cells?.rows?.length) return entry;
    // The page reads at most a capped part of a big sheet.
    if (sheet.cells.used && sheet.cells.used.rowCount > sheet.cells.range.rowCount) truncated = true;
    const formulas = {};
    const values = [];
    for (const [r, row] of sheet.cells.rows.entries()) {
      if (cellBudget <= 0 || textBudget <= 0) { truncated = true; break; }
      const line = [];
      for (const [c, cell] of row.entries()) {
        if (cellBudget-- <= 0) { truncated = true; break; }
        let text = String(cell?.text ?? '');
        if (text.length > textBudget) { text = text.slice(0, Math.max(0, textBudget)); truncated = true; }
        textBudget -= text.length;
        line.push(text);
        if (cell?.formula) formulas[`${columnLetters(sheet.cells.range.column + c)}${sheet.cells.range.row + r + 1}`] = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`;
      }
      values.push(line);
    }
    return { ...entry, values, ...(Object.keys(formulas).length && { formulas }) };
  });
  return { sheets, ...(truncated && { truncated: true }) };
}

/** The progress sentence every failed or stopped spreadsheet edit ends with: there is no rollback. */
export function xlsxProgress(done, total) {
  return done > 0 ? `Steps 1–${done} of ${total} were applied and stay in the editor: nothing was undone and nothing was saved.` : `None of the ${total} steps was applied.`;
}
