import { describeXlsxSteps, summarizeWorkbook, validateXlsxOps } from '../../../core/src/agent/xlsx.mjs';
import { callKey } from './tool-views.mjs';

const failure = (code, message) => Object.assign(new Error(message), { code: `officesdk/${code}` });

/** A read's outline for its chat row: worksheet count, the active sheet and its used range. */
export function workbookOutline(summary) {
  const sheets = Array.isArray(summary?.sheets) ? summary.sheets : [];
  const active = sheets.find(sheet => sheet?.active) ?? sheets[0];
  return { sheets: sheets.length, active: typeof active?.name === 'string' ? active.name : '', used: typeof active?.used === 'string' ? active.used : '' };
}

/**
 * Serves one Host command for a workbook (spec 261105 §6) through the tab's controller, like the Word path: read,
 * flush, and edit with the approval card unless the model is generating this file. No AI undo (no rollback).
 */
export async function runSpreadsheetCommand(bridge, sessionId, command, controller, signal) {
  if (command.kind === 'read') {
    const book = await controller.agent('inspect', {}, signal);
    const summary = summarizeWorkbook(book);
    const result = { format: 'xlsx', path: command.path, version: controller.savedVersion, unsaved: controller.state.dirty === true, editRevision: controller.editRevision, workbook: summary };
    if (command.callId) bridge.views?.record([callKey(command.callId)], { outline: workbookOutline(summary), unsaved: result.unsaved });
    return result;
  }
  if (command.kind === 'flush') { await controller.flushEditor(); return { editRevision: controller.editRevision }; }
  if (command.kind === 'undo') throw failure('unsupported-operation', 'Spreadsheet edits cannot be undone by the AI (there is no rollback). Change the cells back with another spreadsheet_edit.');
  if (command.kind !== 'edit') throw failure('unsupported-operation', 'Unknown document command.');
  const book = await controller.agent('inspect', {}, signal);
  let plan;
  try { plan = validateXlsxOps(command.ops, book); }
  catch (error) { throw failure(error.code ?? 'invalid-edit', error.message); }
  const auto = command.auto === true && !controller.state.generating?.stopped;
  const keys = [command.callId && callKey(command.callId)];
  const record = bridge.views && keys.some(Boolean) ? extra => bridge.views.record(keys, { ...bridge.views.get(keys[0]), summary: command.summary, path: command.path, ops: command.ops, ...extra }) : undefined;
  record?.({ awaiting: false });
  if (auto) bridge.generating(sessionId, command, controller);
  else {
    record?.({ awaiting: true });
    let approved;
    try { approved = await controller.propose({ kind: 'spreadsheet', summary: command.summary, rows: describeXlsxSteps(plan.steps), danger: plan.danger, reasons: plan.reasons }, signal); }
    finally { record?.({ awaiting: false }); }
    if (!approved) return { declined: true };
  }
  // The card named a worksheet for each operation without one (the active sheet then): the editor writes there even if
  // the user switches sheets while the card is open.
  const ops = command.ops.map((op, index) => op && typeof op === 'object' && op.sheet === undefined && typeof plan.steps[index]?.sheet === 'string' ? { ...op, sheet: plan.steps[index].sheet } : op);
  try {
    const result = await controller.agent('apply', { ops, editId: command.editId }, signal);
    record?.({ applied: result?.applied });
    return result;
  } finally { if (auto) bridge.generating(sessionId, command, controller, { idle: true }); }
}
