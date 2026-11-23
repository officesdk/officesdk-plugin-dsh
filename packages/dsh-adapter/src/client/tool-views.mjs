/**
 * Chat rows for the plugin's own presentation tools, registered in DSH's public keyed slot `tool.call.toolview`
 * under the tools' own names (no built-in tool is shadowed). A row says what the call does — action, file, slides,
 * the edit's summary — expands to the operation list and the result, names the reason inline when the call failed,
 * and can open the file in the right panel. Everything is derived from the call's own arguments and result; pages
 * of existing slides come from what this Client recorded while it ran the edit (lost on reload, then left out).
 * Styling: only the plugin's own `sbt-` classes (styles.mjs), coloured with DSH's theme variables.
 */

export const TOOL_VIEW_NS = 'officesdkToolView';
export const TOOL_VIEW_TOOLS = Object.freeze(['presentation_create', 'presentation_read', 'presentation_edit', 'presentation_undo', 'presentation_save', 'document_create', 'document_read', 'document_edit', 'document_save', 'spreadsheet_create', 'spreadsheet_read', 'spreadsheet_edit', 'spreadsheet_save']);

const ZH = {
  'create.running': '正在新建演示文稿', 'create.ok': '新建了演示文稿', 'create.error': '新建演示文稿失败', 'create.declined': '没有新建演示文稿',
  style: '风格 {name}', generatingNote: '边生成边在右侧显示，可在标签页里停止',
  'read.running': '正在读取演示文稿', 'read.ok': '读取了演示文稿', 'read.error': '读取演示文稿失败',
  'edit.running': '正在修改演示文稿', 'edit.ok': '修改了演示文稿', 'edit.error': '修改演示文稿失败', 'edit.declined': '没有修改演示文稿',
  'undo.running': '正在撤销修改', 'undo.ok': '撤销了修改', 'undo.error': '撤销修改失败',
  'save.running': '正在保存演示文稿', 'save.ok': '保存了演示文稿', 'save.error': '保存演示文稿失败', 'save.unchanged': '没有要保存的修改',
  'edit.awaiting': '等你确认修改', 'save.declined': '没有保存演示文稿', 'read.declined': '没有读取演示文稿', 'undo.declined': '没有撤销修改',
  stopped: '已停止', notApproved: '你没有同意',
  slideCount: '共 {n} 页', unsaved: '有未保存的修改', declined: '你没有同意这次修改', notApprovedSave: '你没有同意写入工作区外的文件，修改还在编辑器里', savedTo: '已写回 {path}', unchangedNote: '文件没有变化',
  page: '第 {pages} 页', pages: '第 {pages} 页', pageJoin: '、', newPage: '新增第 {n} 页', newPages: '新增 {n} 页', toPage: '移到第 {n} 页', reasonJoin: '：',
  applied: '已应用 {n} 步', notSaved: '还没保存到文件', undoable: '可以撤销', undone: '已在编辑器里撤销，还没保存到文件',
  shapes: '{n} 个元素', operations: '操作', result: '结果', error: '错误',
  open: '在右侧打开', openTitle: '在右侧打开 {path}', inspect: '查看原始调用', more: '还有 {n} 页没列出',
  'reason.invalid-edit': '操作无效', 'reason.read-only': '会话只读', 'reason.sandbox-denied': '没有写入权限', 'reason.busy': '文档正忙', 'reason.no-editor': '文档没能在右侧打开',
  'reason.not-found': '文档没有打开', 'reason.source-changed': '文件已被别处修改', 'reason.edit-conflict': '修改冲突', 'reason.unsupported-operation': '不支持的操作',
  'reason.forbidden': '会话已不可用', 'reason.unavailable': '文档暂时不可用', 'reason.lease-unavailable': '文档暂时不可用', 'reason.timeout': '超时', 'reason.editor-failed': '编辑器出错', 'reason.invalid-path': '路径不对', 'reason.exists': '同名文件已存在',
  'op.setText': '改文字', 'op.replaceText': '替换文字', 'op.setFont': '改字体', 'op.setAlignment': '改对齐', 'op.setFill': '改填充色', 'op.setGeometry': '改位置大小',
  'op.addTextBox': '加文本框', 'op.addShape': '加形状', 'op.addImage': '加图片', 'op.addTable': '加表格', 'op.addChart': '加图表', 'op.setBackground': '改背景',
  'op.deleteShape': '删除元素', 'op.addSlide': '新增页', 'op.deleteSlide': '删除页', 'op.moveSlide': '移动页',
  'doc.create.running': '正在新建文档', 'doc.create.ok': '新建了文档', 'doc.create.error': '新建文档失败', 'doc.create.declined': '没有新建文档',
  'doc.read.running': '正在读取文档', 'doc.read.ok': '读取了文档', 'doc.read.error': '读取文档失败', 'doc.read.declined': '没有读取文档',
  'doc.edit.running': '正在修改文档', 'doc.edit.ok': '修改了文档', 'doc.edit.error': '修改文档失败', 'doc.edit.declined': '没有修改文档',
  'doc.save.running': '正在保存文档', 'doc.save.ok': '保存了文档', 'doc.save.error': '保存文档失败', 'doc.save.unchanged': '没有要保存的修改', 'doc.save.declined': '没有保存文档',
  template: '模板 {name}', docCount: '共 {n} 段', tableCount: '{n} 个表格', paragraph: '第 {n} 段', table: '表 {n}', cell: '表 {table} 第 {row} 行第 {column} 列', start: '开头', end: '末尾', afterParagraph: '第 {n} 段之后', beforeParagraph: '第 {n} 段之前',
  partial: '已完成第 1–{done} 步（共 {total} 步），这些改动留在编辑器里，没有撤销，也没有保存', partialNone: '一步也没有执行',
  'op.insertParagraph': '新增段落', 'op.formatParagraph': '改格式', 'op.deleteParagraph': '删除段落', 'op.insertTable': '新增表格', 'op.setCell': '改单元格', 'op.insertPageBreak': '分页',
  'sheet.create.running': '正在新建表格', 'sheet.create.ok': '新建了表格', 'sheet.create.error': '新建表格失败', 'sheet.create.declined': '没有新建表格',
  'sheet.read.running': '正在读取表格', 'sheet.read.ok': '读取了表格', 'sheet.read.error': '读取表格失败', 'sheet.read.declined': '没有读取表格',
  'sheet.edit.running': '正在修改表格', 'sheet.edit.ok': '修改了表格', 'sheet.edit.error': '修改表格失败', 'sheet.edit.declined': '没有修改表格',
  'sheet.save.running': '正在保存表格', 'sheet.save.ok': '保存了表格', 'sheet.save.error': '保存表格失败', 'sheet.save.unchanged': '没有要保存的修改', 'sheet.save.declined': '没有保存表格',
  sheetCount: '{n} 个工作表', activeSheet: '当前 {name}', usedRange: '已用 {range}',
  'op.setValues': '写入数值', 'op.setFormula': '写公式', 'op.formatRange': '改格式', 'op.clearRange': '清空', 'op.addSheet': '新建工作表', 'op.renameSheet': '重命名工作表', 'op.activateSheet': '切换工作表',
  'op.insertRows': '插入行', 'op.deleteRows': '删除行', 'op.insertColumns': '插入列', 'op.deleteColumns': '删除列', 'op.setColumnWidth': '改列宽', 'op.freezeRows': '冻结行',
};
const EN = {
  'create.running': 'Creating presentation', 'create.ok': 'Created presentation', 'create.error': 'Could not create presentation', 'create.declined': 'Presentation not created',
  style: 'style {name}', generatingNote: 'Shown beside the chat as it is built; it can be stopped in the tab',
  'read.running': 'Reading presentation', 'read.ok': 'Read presentation', 'read.error': 'Could not read presentation',
  'edit.running': 'Editing presentation', 'edit.ok': 'Edited presentation', 'edit.error': 'Could not edit presentation', 'edit.declined': 'Presentation not changed',
  'undo.running': 'Undoing edit', 'undo.ok': 'Undid edit', 'undo.error': 'Could not undo edit',
  'save.running': 'Saving presentation', 'save.ok': 'Saved presentation', 'save.error': 'Could not save presentation', 'save.unchanged': 'Nothing to save',
  'edit.awaiting': 'Waiting for your approval', 'save.declined': 'Presentation not saved', 'read.declined': 'Presentation not read', 'undo.declined': 'Edit not undone',
  stopped: 'Stopped', notApproved: 'You did not approve',
  slideCount: '{n} slides', unsaved: 'has unsaved changes', declined: 'You did not approve this change', notApprovedSave: 'You did not allow writing outside the workspace; the changes are still in the editor', savedTo: 'Written to {path}', unchangedNote: 'The file did not change',
  page: 'slide {pages}', pages: 'slides {pages}', pageJoin: ', ', newPage: 'new slide {n}', newPages: '{n} new slides', toPage: 'to slide {n}', reasonJoin: ': ',
  applied: '{n} steps applied', notSaved: 'not saved to the file yet', undoable: 'can be undone', undone: 'Undone in the editor, not saved to the file yet',
  shapes: '{n} shapes', operations: 'Operations', result: 'Result', error: 'Error',
  open: 'Open beside', openTitle: 'Open {path} in the right panel', inspect: 'Inspect call', more: '{n} more slides not listed',
  'reason.invalid-edit': 'Invalid operation', 'reason.read-only': 'Read-only session', 'reason.sandbox-denied': 'No write access', 'reason.busy': 'The document is busy', 'reason.no-editor': 'The document did not open beside the chat',
  'reason.not-found': 'The document is not open', 'reason.source-changed': 'The file changed elsewhere', 'reason.edit-conflict': 'Conflicting edits', 'reason.unsupported-operation': 'Unsupported operation',
  'reason.forbidden': 'The session is no longer available', 'reason.unavailable': 'The document is not available right now', 'reason.lease-unavailable': 'The document is not available right now', 'reason.timeout': 'Timed out', 'reason.editor-failed': 'Editor error', 'reason.invalid-path': 'Wrong path', 'reason.exists': 'A file with that name exists',
  'op.setText': 'Set text', 'op.replaceText': 'Replace text', 'op.setFont': 'Font', 'op.setAlignment': 'Alignment', 'op.setFill': 'Fill', 'op.setGeometry': 'Position and size',
  'op.addTextBox': 'Add text box', 'op.addShape': 'Add shape', 'op.addImage': 'Add image', 'op.addTable': 'Add table', 'op.addChart': 'Add chart', 'op.setBackground': 'Background',
  'op.deleteShape': 'Delete shape', 'op.addSlide': 'Add slide', 'op.deleteSlide': 'Delete slide', 'op.moveSlide': 'Move slide',
  'doc.create.running': 'Creating document', 'doc.create.ok': 'Created document', 'doc.create.error': 'Could not create document', 'doc.create.declined': 'Document not created',
  'doc.read.running': 'Reading document', 'doc.read.ok': 'Read document', 'doc.read.error': 'Could not read document', 'doc.read.declined': 'Document not read',
  'doc.edit.running': 'Editing document', 'doc.edit.ok': 'Edited document', 'doc.edit.error': 'Could not edit document', 'doc.edit.declined': 'Document not changed',
  'doc.save.running': 'Saving document', 'doc.save.ok': 'Saved document', 'doc.save.error': 'Could not save document', 'doc.save.unchanged': 'Nothing to save', 'doc.save.declined': 'Document not saved',
  template: 'template {name}', docCount: '{n} paragraphs', tableCount: '{n} tables', paragraph: 'paragraph {n}', table: 'table {n}', cell: 'table {table} row {row} column {column}', start: 'start', end: 'end', afterParagraph: 'after paragraph {n}', beforeParagraph: 'before paragraph {n}',
  partial: 'Steps 1–{done} of {total} were applied and stay in the editor; nothing was undone or saved', partialNone: 'No step was applied',
  'op.insertParagraph': 'Add paragraph', 'op.formatParagraph': 'Format', 'op.deleteParagraph': 'Delete paragraph', 'op.insertTable': 'Add table', 'op.setCell': 'Set cell', 'op.insertPageBreak': 'Page break',
  'sheet.create.running': 'Creating spreadsheet', 'sheet.create.ok': 'Created spreadsheet', 'sheet.create.error': 'Could not create spreadsheet', 'sheet.create.declined': 'Spreadsheet not created',
  'sheet.read.running': 'Reading spreadsheet', 'sheet.read.ok': 'Read spreadsheet', 'sheet.read.error': 'Could not read spreadsheet', 'sheet.read.declined': 'Spreadsheet not read',
  'sheet.edit.running': 'Editing spreadsheet', 'sheet.edit.ok': 'Edited spreadsheet', 'sheet.edit.error': 'Could not edit spreadsheet', 'sheet.edit.declined': 'Spreadsheet not changed',
  'sheet.save.running': 'Saving spreadsheet', 'sheet.save.ok': 'Saved spreadsheet', 'sheet.save.error': 'Could not save spreadsheet', 'sheet.save.unchanged': 'Nothing to save', 'sheet.save.declined': 'Spreadsheet not saved',
  sheetCount: '{n} sheets', activeSheet: 'active {name}', usedRange: 'used {range}',
  'op.setValues': 'Set values', 'op.setFormula': 'Formula', 'op.formatRange': 'Format', 'op.clearRange': 'Clear', 'op.addSheet': 'Add sheet', 'op.renameSheet': 'Rename sheet', 'op.activateSheet': 'Switch sheet',
  'op.insertRows': 'Insert rows', 'op.deleteRows': 'Delete rows', 'op.insertColumns': 'Insert columns', 'op.deleteColumns': 'Delete columns', 'op.setColumnWidth': 'Column width', 'op.freezeRows': 'Freeze rows',
};
export const TOOL_VIEW_MESSAGES = Object.freeze({ zh: ZH, en: EN });

const KINDS = { presentation_create: 'create', presentation_read: 'read', presentation_edit: 'edit', presentation_undo: 'undo', presentation_save: 'save',
  document_create: 'create', document_read: 'read', document_edit: 'edit', document_save: 'save',
  spreadsheet_create: 'create', spreadsheet_read: 'read', spreadsheet_edit: 'edit', spreadsheet_save: 'save' };
/** Word tools title their rows with the `doc.` keys, spreadsheet tools with `sheet.`; presentation keys are unprefixed. */
const isDocument = toolName => typeof toolName === 'string' && toolName.startsWith('document_');
const isSpreadsheet = toolName => typeof toolName === 'string' && toolName.startsWith('spreadsheet_');
const parse = text => { try { return JSON.parse(text); } catch { return undefined; } };
const isObject = value => typeof value === 'object' && value !== null && !Array.isArray(value);
const str = value => typeof value === 'string' ? value : undefined;
const firstLine = text => { const index = text.indexOf('\n'); return index === -1 ? text : text.slice(0, index); };
const clip = (text, size = 80) => text.length > size ? `${text.slice(0, size - 1)}…` : text;
const page = index => Number.isInteger(index) && index >= 0 ? index + 1 : undefined;

/** A dictionary entry with `{name}` placeholders filled; without a dictionary the key itself. */
export const format = (t, key, params = {}) => {
  const template = (typeof t === 'function' ? t(key) : undefined) ?? key;
  return Object.entries(params).reduce((text, [name, value]) => text.split(`{${name}}`).join(String(value)), String(template));
};
/** Last path segment, either separator. */
export const fileName = path => path.split(/[\\/]/).filter(Boolean).at(-1) ?? path;

/** Text blocks of a settled result, joined the way DSH's own rows flatten them. */
function resultText(block) {
  return (Array.isArray(block?.content) ? block.content : []).filter(item => item?.type === 'text' && typeof item.text === 'string').map(item => item.text).join('\n');
}

/**
 * The call's arguments as far as they are known: the dispatched JSON once there is one, else the complete string
 * fields the preparing reader holds (operations only arrive with the dispatched call).
 */
export function callArgs(phase, block) {
  const raw = phase === 'result' ? block?.call?.argsRaw : phase === 'start' ? block?.argsRaw : undefined;
  const parsed = typeof raw === 'string' ? parse(raw) : undefined;
  if (isObject(parsed)) return parsed;
  const reader = block?.args, out = {};
  for (const key of ['path', 'summary', 'editId', 'style', 'template']) {
    try { if (reader?.complete?.(key)) { const value = reader.text?.(key); if (typeof value === 'string') out[key] = value; } } catch {}
  }
  return out;
}

/** The plugin's own cancellation wordings (Host hub, editor agents, core aborts, Client waits). */
const CANCELLED = /^(?:The (?:document command|document operation|edit) was cancelled|Cancelled\.)/;

/** Codes the Host reports (its CLIENT_CODES plus its own); anything else keeps only its first line. */
const KNOWN_CODES = new Set(['invalid-edit', 'read-only', 'sandbox-denied', 'busy', 'no-editor', 'not-found', 'source-changed', 'edit-conflict', 'unsupported-operation',
  'forbidden', 'unavailable', 'lease-unavailable', 'timeout', 'editor-failed', 'invalid-path', 'exists', 'cancelled', 'selection-unknown', 'selection-stale', 'selection-used']);

/**
 * How a failed result reads. A tool's throw arrives as text `Error: <message>` with no `error` object (dsh-tools
 * `toolErrorResult`; the agent loop only forwards `error.info`); the plugin's messages are `<code>: <reason>`. DSH's
 * own outcomes: `error.code` interrupted or ABORTED_BEFORE_DISPATCH, and its approval rejected or cancelled (text
 * only). `outcome` is stopped | declined | failed.
 */
export function failureOf(block, recordedCode) {
  const text = resultText(block) || str(block?.error?.message) || [block?.error?.name, block?.error?.code].filter(Boolean).join(': ');
  const line = firstLine(text).replace(/^(?:Error:\s*)+/, '');
  const dshCode = block?.error?.code ?? block?.error?.info?.code;
  if (dshCode === 'interrupted' || dshCode === 'ABORTED_BEFORE_DISPATCH' || /^tool call aborted before dispatch/.test(line)
    || /^approval for tool "[^"]*" was cancelled/.test(line)) return { outcome: 'stopped', full: text };
  if (/^the user rejected tool "/.test(line)) return { outcome: 'declined', full: text };
  const match = /^(?:officesdk\/)?([a-z][a-z-]*):\s*(.*)$/.exec(line);
  // The packaged tools and Host bundle the core separately, so the tools' `instanceof SuiteError` misses and the text
  // often carries no code: fall back to what this Client recorded for the call, then to the Host's fixed wording.
  const code = match?.[1] ?? (typeof recordedCode === 'string' ? recordedCode : undefined) ?? (CANCELLED.test(line) ? 'cancelled' : /^The user did not allow /.test(line) ? 'sandbox-denied' : undefined);
  const reason = match ? match[2] : line;
  if (code === 'cancelled') return { outcome: 'stopped', reason, full: text };
  // Creating or saving outside the workspace asks the user (Host: `did not allow … (<outcome>)`): their no is not a failure, a
  // cancelled ask is a stop, no way to ask stays a failure.
  const asked = code === 'sandbox-denied' && /did not allow/.test(reason) ? /outside the workspace \((rejected|cancelled)\)/.exec(reason)?.[1] : undefined;
  if (asked === 'rejected') return { outcome: 'declined', code, reason, full: text };
  if (asked === 'cancelled') return { outcome: 'stopped', code, reason, full: text };
  return { outcome: 'failed', code: KNOWN_CODES.has(code) ? code : undefined, reason, full: text };
}

/**
 * What the bridge records for an edit, from a deck it inspected: each named slide's 1-based page and each named
 * shape's name, and — after the edit — each ref's new page. Before the edit `refs` is empty.
 */
export function editContext(ops, deck, { slideIds = [], refs = {} } = {}) {
  const list = (Array.isArray(ops) ? ops : []).filter(isObject);
  const refIds = isObject(refs) ? refs : {};
  const wanted = new Set([...list.map(op => op.slideId), ...slideIds, ...Object.values(refIds)].filter(id => typeof id === 'string'));
  const slides = {}, shapes = {}, pages = {};
  for (const [index, slide] of (Array.isArray(deck?.slides) ? deck.slides : []).entries()) {
    if (!wanted.has(slide?.id)) continue;
    slides[slide.id] = index + 1;
    for (const shape of Array.isArray(slide.shapes) ? slide.shapes : []) {
      if (list.some(op => op.slideId === slide.id && op.shapeId === shape?.id)) shapes[`${slide.id}/${shape.id}`] = shape.name || shape.id;
    }
  }
  for (const [ref, id] of Object.entries(refIds)) if (slides[id]) pages[ref] = slides[id];
  return { slides, shapes, refs: pages };
}

/**
 * After the edit: the pages as they are now win, slides the edit deleted keep the page they had, `from` keeps the
 * pages before the edit (a moved slide's row says where it came from), and `touched` lists the pages of the result's
 * slide ids (new slides included, even without a ref).
 */
export function settleContext(before, after, slideIds = []) {
  const slides = { ...before?.slides, ...after.slides };
  const touched = [...new Set(slideIds.map(id => slides[id]).filter(Number.isInteger))].sort((a, b) => a - b);
  return { ...before, slides, from: { ...before?.slides }, shapes: { ...before?.shapes, ...after.shapes }, refs: after.refs, touched, awaiting: false };
}

/** The page one operation targets, or undefined: recorded ref or slide, else what the arguments say. */
function pageOf(op, argRefs, context) {
  if (op.op === 'addSlide') return (typeof op.ref === 'string' ? context?.refs?.[op.ref] : undefined) ?? page(op.index);
  if (typeof op.slideRef === 'string') return context?.refs?.[op.slideRef] ?? argRefs.get(op.slideRef);
  // A move names its target in the row's text; where it says the page it left.
  if (op.op === 'moveSlide' && typeof op.slideId === 'string') return context?.from?.[op.slideId] ?? context?.slides?.[op.slideId];
  if (typeof op.slideId === 'string') return context?.slides?.[op.slideId];
  return undefined;
}
const argPages = ops => new Map(ops.filter(op => op.op === 'addSlide' && typeof op.ref === 'string' && page(op.index)).map(op => [op.ref, page(op.index)]));

/** Sorted distinct known pages, and how many new slides have none. */
export function editPages(ops, context) {
  if (Array.isArray(context?.touched) && context.touched.length) return { pages: context.touched, unplaced: 0 };
  const list = (Array.isArray(ops) ? ops : []).filter(isObject);
  const refs = argPages(list);
  const known = new Set();
  let unplaced = 0;
  for (const op of list) {
    const at = pageOf(op, refs, context);
    if (at) known.add(at); else if (op.op === 'addSlide') unplaced += 1;
  }
  return { pages: [...known].sort((a, b) => a - b), unplaced };
}

/** Consecutive pages as one range: `2–5、7` / `2–5, 7`; two in a row stay a pair. */
export function pageRanges(pages, joiner) {
  const runs = [];
  for (const at of pages) {
    const last = runs.at(-1);
    if (last && at === last[1] + 1) last[1] = at; else runs.push([at, at]);
  }
  return runs.map(([from, to]) => from === to ? String(from) : to === from + 1 ? `${from}${joiner}${to}` : `${from}–${to}`).join(joiner);
}

export function pagesLabel(t, { pages, unplaced = 0 }) {
  return [pages.length ? format(t, pages.length === 1 ? 'page' : 'pages', { pages: pageRanges(pages, format(t, 'pageJoin')) }) : '', unplaced ? format(t, 'newPages', { n: unplaced }) : ''].filter(Boolean).join(' · ');
}

/** One line per operation: what, where (page · shape), and its content in short. */
export function editRows(t, ops, context) {
  const list = (Array.isArray(ops) ? ops : []).filter(isObject);
  const refs = argPages(list);
  return list.map(op => {
    const at = pageOf(op, refs, context);
    const shape = typeof op.slideId === 'string' && typeof op.shapeId === 'string' ? context?.shapes?.[`${op.slideId}/${op.shapeId}`] : undefined;
    const where = op.op === 'addSlide' ? (at ? format(t, 'newPage', { n: at }) : '') : [at ? format(t, 'page', { pages: at }) : '', shape ?? ''].filter(Boolean).join(' · ');
    return { op: str(op.op) ?? '?', label: format(t, `op.${op.op}`), where, text: clip(operationText(t, op)) };
  });
}

function operationText(t, op) {
  switch (op.op) {
    case 'setText': case 'addTextBox': return firstLine(str(op.text) ?? '');
    case 'replaceText': return typeof op.find === 'string' && typeof op.replace === 'string' ? `${firstLine(op.find)} → ${firstLine(op.replace)}` : '';
    case 'setFont': return ['name', 'size', 'color', 'bold', 'italic', 'underline'].filter(key => ['string', 'number', 'boolean'].includes(typeof op[key])).map(key => `${key}=${op[key]}`).join(', ');
    case 'setAlignment': return str(op.align) ?? '';
    case 'setFill': case 'setBackground': return str(op.color) ?? '';
    case 'setGeometry': return ['left', 'top', 'width', 'height'].filter(key => typeof op[key] === 'number').map(key => `${key}=${Math.round(op[key])}`).join(', ');
    case 'addShape': return [str(op.shape), firstLine(str(op.text) ?? '')].filter(Boolean).join(': ');
    case 'addImage': return typeof op.path === 'string' ? fileName(op.path) : '';
    case 'addSlide': return firstLine(str(op.placeholders?.title) ?? '');
    case 'moveSlide': return page(op.index) ? format(t, 'toPage', { n: page(op.index) }) : '';
    default: return '';
  }
}

/** A read result's outline (the bridge records the same from the deck it read). */
export function readOutline(deck) {
  const slides = Array.isArray(deck?.slides) ? deck.slides : [];
  return slides.map(slide => {
    const shapes = Array.isArray(slide?.shapes) ? slide.shapes : [];
    const text = shapes.map(shape => str(shape?.text)?.trim()).find(Boolean);
    return { shapes: shapes.length, text: text ? clip(firstLine(text)) : '' };
  });
}

/** One line per Word operation: what, where (paragraph / table numbers are 1-based for people), and its text. */
export function documentRows(t, ops) {
  const n = value => Number.isInteger(value) && value >= 0 ? value + 1 : undefined;
  const target = op => typeof op.paragraphRef === 'string' ? `@${op.paragraphRef}` : n(op.paragraph) ? format(t, 'paragraph', { n: n(op.paragraph) }) : '';
  const where = op => {
    if (op.op === 'setCell') return n(op.table) || typeof op.tableRef === 'string' ? format(t, 'cell', { table: n(op.table) ?? `@${op.tableRef}`, row: n(op.row) ?? '?', column: n(op.column) ?? '?' }) : '';
    if (['setText', 'replaceText', 'formatParagraph', 'deleteParagraph'].includes(op.op)) return target(op);
    if (n(op.after)) return format(t, 'afterParagraph', { n: n(op.after) });
    if (n(op.before)) return format(t, 'beforeParagraph', { n: n(op.before) });
    if (typeof op.afterRef === 'string') return format(t, 'afterParagraph', { n: `@${op.afterRef}` });
    if (typeof op.beforeRef === 'string') return format(t, 'beforeParagraph', { n: `@${op.beforeRef}` });
    return format(t, op.at === 'start' ? 'start' : 'end');
  };
  const content = op => {
    switch (op.op) {
      case 'insertParagraph': case 'setText': case 'setCell': return [str(op.style), str(op.list), firstLine(str(op.text) ?? '')].filter(Boolean).join(' · ');
      case 'replaceText': return typeof op.find === 'string' && typeof op.replace === 'string' ? `${firstLine(op.find)} → ${firstLine(op.replace)}` : '';
      case 'formatParagraph': return [str(op.style), str(op.list), ...['align', 'bold', 'italic', 'underline', 'size', 'color', 'font'].filter(key => ['string', 'number', 'boolean'].includes(typeof op[key])).map(key => `${key}=${op[key]}`)].filter(Boolean).join(', ');
      case 'insertTable': return Array.isArray(op.values) && Array.isArray(op.values[0]) ? `${op.values.length}×${op.values[0].length}: ${op.values[0].join(' | ')}` : '';
      default: return '';
    }
  };
  return (Array.isArray(ops) ? ops : []).filter(isObject).map(op => ({ op: str(op.op) ?? '?', label: format(t, `op.${op.op}`), where: where(op), text: clip(content(op)) }));
}

/** One line per spreadsheet operation: what, where (sheet!range), and its content in short. */
export function spreadsheetRows(t, ops) {
  const where = op => {
    const sheet = str(op.sheet);
    const place = str(op.range) ?? str(op.cell) ?? str(op.columns) ?? str(op.column) ?? (Number.isInteger(op.row) ? String(op.row) : undefined);
    if (op.op === 'addSheet' || op.op === 'renameSheet') return [sheet, str(op.name)].filter(Boolean).join(' → ');
    return [sheet, place].filter(Boolean).join('!');
  };
  const content = op => {
    switch (op.op) {
      case 'setValues': return Array.isArray(op.values) ? op.values.slice(0, 2).map(row => (Array.isArray(row) ? row : []).map(value => value ?? '').join(' | ')).join(' / ') : '';
      case 'setFormula': return str(op.formula) ?? '';
      case 'formatRange': return ['bold', 'italic', 'underline', 'color', 'background', 'fontSize', 'align', 'wrap', 'numberFormat'].filter(key => ['string', 'number', 'boolean'].includes(typeof op[key])).map(key => `${key}=${op[key]}`).join(', ');
      case 'setColumnWidth': return typeof op.width === 'number' ? `${op.width}px` : '';
      case 'addChart': return str(op.title) ?? '';
      case 'insertRows': case 'deleteRows': case 'insertColumns': case 'deleteColumns': case 'freezeRows': return Number.isInteger(op.count) ? `×${op.count}` : '';
      default: return '';
    }
  };
  return (Array.isArray(ops) ? ops : []).filter(isObject).map(op => ({ op: str(op.op) ?? '?', label: format(t, `op.${op.op}`), where: where(op), text: clip(content(op)) }));
}

/** How far a failed or stopped Word edit got, from the page's progress sentence; undefined when it does not say. */
export function documentProgress(t, text) {
  const done = /Steps 1–(\d+) of (\d+) were applied/.exec(text ?? '');
  if (done) return format(t, 'partial', { done: done[1], total: done[2] });
  if (/None of the \d+ steps? was applied/.test(text ?? '')) return format(t, 'partialNone');
  return undefined;
}

/** Keys under which the bridge records a call, and an edit for the undo that names it. */
export const callKey = callId => `call:${callId}`;
export const editKey = editId => `edit:${editId}`;
const MAX_LISTED = 50;

/**
 * Everything one row shows, from the owner props DSH passes (`toolName`, `callId`, `phase`, `block`) and what this
 * Client recorded; `recall(key)` returns a record or undefined.
 */
export function toolViewModel(t, { toolName, callId, phase, block }, recall = () => undefined) {
  const kind = KINDS[toolName] ?? 'edit';
  const sheet = isSpreadsheet(toolName);
  // Word and spreadsheets share the no-rollback rows; only titles and operation rows differ.
  const doc = isDocument(toolName) || sheet;
  // Titles: `doc.<kind>.<state>` for Word, `sheet.<kind>.<state>` for spreadsheets, `<kind>.<state>` for presentations.
  const key = name => doc && /^(create|read|edit|save)\./.test(name) ? `${sheet ? 'sheet' : 'doc'}.${name}` : name;
  const args = callArgs(phase, block);
  const path = str(args.path) ?? '';
  const recorded = kind === 'undo' ? (typeof args.editId === 'string' ? recall(editKey(args.editId)) : undefined) : callId !== undefined ? recall(callKey(callId)) : undefined;
  const model = { kind, path, file: path ? fileName(path) : '', state: 'running', tone: undefined, title: '', pages: '', summary: '', rows: [], slides: [], lines: [], error: undefined };
  // Without a named style the Host picks the workspace default, so the style is said only once it is named or known.
  // document_create / spreadsheet_create default to the gongwen / blank template; while the arguments stream, a template not yet written is left out.
  if (kind === 'create' && doc && (phase !== 'preparing' || str(args.template))) model.summary = format(t, 'template', { name: str(args.template) || (sheet ? 'blank' : 'gongwen') });
  else if (kind === 'create' && !doc && str(args.style)) model.summary = format(t, 'style', { name: str(args.style) });
  if (doc && kind === 'edit') {
    model.summary = firstLine(str(args.summary) ?? '');
    model.rows = sheet ? spreadsheetRows(t, args.operations) : documentRows(t, args.operations);
  } else if (kind === 'edit' || kind === 'undo') {
    const ops = kind === 'edit' ? args.operations : recorded?.ops;
    model.pages = pagesLabel(t, editPages(ops, recorded));
    model.summary = firstLine(str(kind === 'edit' ? args.summary : recorded?.summary) ?? '');
    if (kind === 'edit') model.rows = editRows(t, ops, recorded);
  }
  if (phase !== 'result') {
    model.state = phase === 'preparing' ? 'preparing' : 'running';
    model.title = format(t, kind === 'edit' && recorded?.awaiting ? 'edit.awaiting' : key(`${kind}.running`));
    return model;
  }
  if (block?.isError) {
    const failure = failureOf(block, recorded?.failure);
    model.error = failure.full || undefined;
    // A Word edit has no rollback: the row says how far it got.
    const progress = doc && kind === 'edit' ? documentProgress(t, failure.full) : undefined;
    if (progress) model.lines.push(progress);
    if (failure.outcome === 'stopped') return Object.assign(model, { state: 'stopped', tone: 'warn', title: format(t, 'stopped') });
    if (failure.outcome === 'declined') {
      model.lines.push(format(t, kind === 'save' ? 'notApprovedSave' : 'notApproved'));
      return Object.assign(model, { state: 'declined', tone: 'muted', title: format(t, key(`${kind}.declined`)), error: undefined });
    }
    // An unknown or broken style pack is reported as invalid-edit; for a new file that category would mislead.
    const category = failure.code && !(kind === 'create' && failure.code === 'invalid-edit') ? format(t, `reason.${failure.code}`) : '';
    const reason = category && failure.reason ? `${category}${format(t, 'reasonJoin')}${failure.reason}` : category || failure.reason;
    return Object.assign(model, { state: 'error', tone: 'error', title: format(t, key(`${kind}.error`)), summary: reason || model.summary });
  }
  const value = parse(resultText(block));
  const result = isObject(value) ? value : {};
  Object.assign(model, { state: 'ok', title: format(t, key(`${kind}.ok`)) });
  if (doc && kind === 'create') {
    if (typeof result.template === 'string') model.summary = format(t, 'template', { name: result.template });
    if (result.generating === true) model.lines.push(format(t, 'generatingNote'));
  } else if (sheet && kind === 'read') {
    const outline = isObject(result.workbook) && Array.isArray(result.workbook.sheets) ? { sheets: result.workbook.sheets.length, active: str((result.workbook.sheets.find(item => item?.active) ?? result.workbook.sheets[0])?.name) ?? '', used: str((result.workbook.sheets.find(item => item?.active) ?? result.workbook.sheets[0])?.used) ?? '' } : recorded?.outline ?? {};
    const unsaved = typeof result.unsaved === 'boolean' ? result.unsaved : recorded?.unsaved === true;
    model.summary = [Number.isInteger(outline.sheets) ? format(t, 'sheetCount', { n: outline.sheets }) : '', outline.active ? format(t, 'activeSheet', { name: outline.active }) : '', outline.used ? format(t, 'usedRange', { range: outline.used }) : '', unsaved ? format(t, 'unsaved') : ''].filter(Boolean).join(' · ');
  } else if (doc && kind === 'read') {
    const outline = isObject(result.document) ? { paragraphs: Array.isArray(result.document.paragraphs) ? result.document.paragraphs.length : 0, tables: Array.isArray(result.document.tables) ? result.document.tables.length : 0 } : recorded?.outline ?? {};
    const unsaved = typeof result.unsaved === 'boolean' ? result.unsaved : recorded?.unsaved === true;
    model.summary = [Number.isInteger(outline.paragraphs) ? format(t, 'docCount', { n: outline.paragraphs }) : '', outline.tables ? format(t, 'tableCount', { n: outline.tables }) : '', unsaved ? format(t, 'unsaved') : ''].filter(Boolean).join(' · ');
  } else if (doc && kind === 'edit') {
    if (result.declined === true) {
      Object.assign(model, { state: 'declined', tone: 'muted', title: format(t, key('edit.declined')) });
      model.lines.push(format(t, 'declined'));
    } else model.lines.push([typeof result.applied === 'number' ? format(t, 'applied', { n: result.applied }) : '', result.unsaved === true ? format(t, 'notSaved') : ''].filter(Boolean).join(' · '));
  } else if (doc && kind === 'save') {
    if (result.unchanged === true) {
      Object.assign(model, { state: 'unchanged', tone: 'muted', title: format(t, key('save.unchanged')) });
      model.lines.push(format(t, 'unchangedNote'));
    } else if (typeof result.path === 'string') model.lines.push(format(t, 'savedTo', { path: result.path }));
  } else if (kind === 'create') {
    if (typeof result.style?.name === 'string') model.summary = format(t, 'style', { name: result.style.name });
    if (result.generating === true) model.lines.push(format(t, 'generatingNote'));
  } else if (kind === 'read') {
    const outline = Array.isArray(result.deck?.slides) ? readOutline(result.deck) : recorded?.outline ?? [];
    const unsaved = typeof result.unsaved === 'boolean' ? result.unsaved : recorded?.unsaved === true;
    model.summary = [outline.length ? format(t, 'slideCount', { n: outline.length }) : '', unsaved ? format(t, 'unsaved') : ''].filter(Boolean).join(' · ');
    model.slides = outline.slice(0, MAX_LISTED).map((slide, index) => ({ page: format(t, 'page', { pages: index + 1 }), shapes: format(t, 'shapes', { n: slide.shapes }), text: slide.text }));
    if (outline.length > MAX_LISTED) model.lines.push(format(t, 'more', { n: outline.length - MAX_LISTED }));
  } else if (kind === 'edit') {
    if (result.declined === true) {
      Object.assign(model, { state: 'declined', tone: 'muted', title: format(t, 'edit.declined') });
      model.lines.push(format(t, 'declined'));
    } else {
      model.lines.push([typeof result.applied === 'number' ? format(t, 'applied', { n: result.applied }) : '', result.unsaved === true ? format(t, 'notSaved') : '', typeof result.editId === 'string' ? format(t, 'undoable') : ''].filter(Boolean).join(' · '));
    }
  } else if (kind === 'undo') {
    model.lines.push(format(t, 'undone'));
  } else if (kind === 'save') {
    if (result.unchanged === true) {
      Object.assign(model, { state: 'unchanged', tone: 'muted', title: format(t, 'save.unchanged') });
      model.lines.push(format(t, 'unchangedNote'));
    } else if (typeof result.path === 'string') model.lines.push(format(t, 'savedTo', { path: result.path }));
  }
  model.lines = model.lines.filter(Boolean);
  return model;
}

/** What the bridge recorded while it ran calls; rows read it by call id (an undo by the edit id it names). */
export function createToolViewStore({ max = 200 } = {}) {
  const entries = new Map();
  const listeners = new Set();
  let version = 0;
  return {
    get: key => entries.get(key),
    version: () => version,
    /** Stores `value` under every non-empty key; past `max` the oldest go. */
    record(keys, value) {
      for (const key of keys) if (key) { entries.delete(key); entries.set(key, value); }
      while (entries.size > max) entries.delete(entries.keys().next().value);
      version += 1;
      for (const listener of [...listeners]) listener();
    },
    subscribe(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  };
}

// 24×24 line icons.
const DECK = 'M3 4h18v12H3zM8 20h8M12 16v4M7 8h6M7 11h10';
const CHEVRON = 'm6 9 6 6 6-6';
const OPEN = 'M14 4h6v6M20 4l-8 8M18 14v6H4V6h6';

/**
 * The row component for `tool.call.toolview`, closed over the host React, the record store, and the deck opener
 * (`openDeck(sessionId, path)`, `canOpenDeck(sessionId)`, `subscribeDeck(listener)` for when that answer may change).
 * DSH passes the owner props (`callId`, `toolName`, `phase`, `block`, `inspect`, `useDisclosure`), `t` for this
 * namespace, and `shimoSessionId` from the entry's inject.
 */
export function createToolView({ React, store, openDeck, canOpenDeck = () => true, subscribeDeck = () => () => {} }) {
  const { createElement: h, useState, useSyncExternalStore } = React;
  const subscribe = store?.subscribe ?? (() => () => {});
  const snapshot = store?.version ?? (() => 0);
  const icon = (d, className) => h('svg', { className, width: 14, height: 14, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': true }, h('path', { d }));

  return function ShimoToolRow(props) {
    const { t, phase, inspect, shimoSessionId: sessionId } = props;
    const [localOpen, setLocalOpen] = useState(false);
    // DSH's disclosure resets when the turn folds; the local one is for hosts without it.
    const disclosure = props.useDisclosure ? props.useDisclosure() : { expanded: localOpen, toggle: () => setLocalOpen(open => !open) };
    useSyncExternalStore(subscribe, snapshot);
    const deckShown = useSyncExternalStore(subscribeDeck, () => typeof sessionId === 'string' && sessionId !== '' && canOpenDeck(sessionId));
    const model = toolViewModel(t, props, key => store?.get(key));
    const expandable = phase !== 'preparing' && (model.rows.length > 0 || model.slides.length > 0 || model.lines.length > 0 || model.error !== undefined || typeof inspect === 'function');
    const open = expandable && disclosure.expanded === true;
    const toggle = () => { if (expandable) disclosure.toggle(); };
    const onKeyDown = event => { if (expandable && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); toggle(); } };
    const busy = model.state === 'running' || model.state === 'preparing';
    const canOpen = model.path !== '' && deckShown && typeof openDeck === 'function';
    const sep = () => h('span', { className: 'sbt-sep', 'aria-hidden': true });
    const part = (text, className, title = text) => text ? [sep(), h('span', { className: `sbt-part ${className}`.trim(), title }, text)] : [];
    // The title carries a stop or failure; the summary is coloured only when it is the failure's reason.
    const titleClass = model.state === 'stopped' || model.state === 'error' ? `sbt-title sbt-${model.tone}` : 'sbt-title';
    const summaryClass = model.state === 'error' ? 'sbt-summary sbt-error' : 'sbt-summary';
    return h('div', { className: 'sbt', 'data-tool': props.toolName, 'data-state': model.state },
      h('div', { className: 'sbt-line' },
        h('div', { className: 'sbt-row', 'data-expandable': expandable || undefined, 'data-busy': busy || undefined,
          ...(expandable ? { role: 'button', tabIndex: 0, 'aria-expanded': open, onClick: toggle, onKeyDown } : {}) },
          h('span', { className: 'sbt-lead' }, open ? icon(CHEVRON) : icon(DECK)),
          h('span', { className: titleClass }, model.title),
          ...part(model.file, 'sbt-file', model.path),
          ...part(model.pages, 'sbt-pages'),
          ...part(model.summary, summaryClass)),
        canOpen ? h('button', { type: 'button', className: 'sbt-open', title: format(t, 'openTitle', { path: model.path }), 'aria-label': format(t, 'open'),
          onClick: event => { event.stopPropagation(); openDeck(sessionId, model.path); } }, icon(OPEN), h('span', null, format(t, 'open'))) : null),
      open ? h('div', { className: 'sbt-body' },
        model.rows.length ? h('ol', { className: 'sbt-list', 'aria-label': format(t, 'operations') }, model.rows.map((row, index) => h('li', { key: index },
          h('span', { className: 'sbt-op' }, row.label), row.where ? h('span', { className: 'sbt-where' }, row.where) : null, row.text ? h('span', { className: 'sbt-text' }, row.text) : null))) : null,
        model.slides.length ? h('ol', { className: 'sbt-list' }, model.slides.map((slide, index) => h('li', { key: index },
          h('span', { className: 'sbt-op' }, slide.page), h('span', { className: 'sbt-where' }, slide.shapes), slide.text ? h('span', { className: 'sbt-text' }, slide.text) : null))) : null,
        model.lines.map((line, index) => h('div', { key: `line${index}`, className: 'sbt-result' }, line)),
        model.error !== undefined ? h('pre', { className: `sbt-detail sbt-${model.tone ?? 'error'}`, 'aria-label': format(t, 'error') }, model.error) : null,
        typeof inspect === 'function' ? h('button', { type: 'button', className: 'sbt-inspect', onClick: inspect }, format(t, 'inspect')) : null) : null);
  };
}

/**
 * Registers the rows: the dictionaries, then one `tool.call.toolview` entry per tool under its own name (no
 * priority: nothing built in uses these keys). Each entry gets its Session id through `inject` (a plugin-specific
 * name, owner props win clashes), so opening the file uses the same address as the AI's tab.
 */
export function registerToolViews(ctx, React, { store, deck } = {}) {
  ctx.effect(() => ctx.locale.register(TOOL_VIEW_NS, TOOL_VIEW_MESSAGES));
  const Row = createToolView({ React, store, ...deck });
  for (const key of TOOL_VIEW_TOOLS) {
    ctx.effect(() => ctx.slots.inject('tool.call.toolview', () => ctx.slots.register({ name: 'tool.call.toolview', key, locale: TOOL_VIEW_NS, inject: sessionId => ({ shimoSessionId: sessionId }) }, Row)));
  }
  return Row;
}
