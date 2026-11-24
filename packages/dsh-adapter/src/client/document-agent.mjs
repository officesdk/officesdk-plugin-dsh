import { checkDocxToken, describeDocxSteps, summarizeDocument, validateDocxOps } from '../../../core/src/agent/docx.mjs';
import { callKey } from './tool-views.mjs';

const failure = (code, message) => Object.assign(new Error(message), { code: `officesdk/${code}` });

/** A read's outline for its chat row: paragraph count, table count, the first heading-like text. */
export function documentOutline(summary) {
  const paragraphs = Array.isArray(summary?.paragraphs) ? summary.paragraphs : [];
  const first = paragraphs.map(paragraph => typeof paragraph?.text === 'string' ? paragraph.text.trim() : '').find(Boolean) ?? '';
  return { paragraphs: paragraphs.length, tables: Array.isArray(summary?.tables) ? summary.tables.length : 0, first: first.length > 80 ? `${first.slice(0, 79)}…` : first };
}

/**
 * Serves one Host command for a Word document (spec 261105 §3.4) through the tab's controller, the way AgentBridge.run
 * serves presentations: read, flush, and edit with the approval card unless the model is generating this file. There
 * is no AI undo for Word (office-kit K3). `bridge` is the AgentBridge (its records, views and generating state).
 */
export async function runDocumentCommand(bridge, sessionId, command, controller, signal) {
  if (command.kind === 'read') {
    const doc = await controller.agent('inspect', command.editTarget ? { editTarget: command.editTarget } : {}, signal);
    let summary;
    if (command.editTarget) {
      const { paragraph, readToken } = command.editTarget;
      const entry = doc.paragraphs?.[paragraph];
      if (!entry) throw failure('source-changed', 'The targeted paragraph is no longer in the document.');
      summary = summarizeDocument({ paragraphs: [entry], tables: [] });
      summary.paragraphs[0].index = paragraph;
      summary.paragraphCount = doc.paragraphs.length;
      summary.scope = 'editTarget';
      // The model only sees this passage. Do not promote its authority to a newer
      // whole-document revision containing changes it did not read.
      summary.readToken = readToken;
    } else summary = summarizeDocument(doc);
    const result = { format: 'docx', path: command.path, version: controller.savedVersion, unsaved: controller.state.dirty === true, editRevision: controller.editRevision, document: summary };
    if (command.callId) bridge.views?.record([callKey(command.callId)], { outline: documentOutline(summary), unsaved: result.unsaved });
    return result;
  }
  if (command.kind === 'flush') { await controller.flushEditor(); return { editRevision: controller.editRevision }; }
  if (command.kind === 'undo') throw failure('unsupported-operation', 'Word edits cannot be undone by the AI (there is no rollback). Change the text back with another document_edit.');
  if (command.kind !== 'edit') throw failure('unsupported-operation', 'Unknown document command.');
  const doc = await controller.agent('inspect', { checkEdit: { ops: command.ops, readToken: command.readToken } }, signal);
  let plan;
  try {
    if (doc.checkedEdit !== true) checkDocxToken(command.ops, doc, command.readToken);
    plan = validateDocxOps(command.ops, doc);
  }
  catch (error) { throw failure(error.code ?? 'invalid-edit', error.message); }
  const auto = command.auto === true && !controller.state.generating?.stopped;
  const keys = [command.callId && callKey(command.callId)];
  const record = bridge.views && keys.some(Boolean) ? extra => bridge.views.record(keys, { ...bridge.views.get(keys[0]), summary: command.summary, path: command.path, ops: command.ops, ...extra }) : undefined;
  record?.({ awaiting: false });
  // Direct requests typed in the DSH composer can also target existing text without a selection ID.
  // Mixed/dependent batches retain their atomic approval. A pure text change in an existing paragraph
  // can use the same staged region and stale-text check as an explicit selection.
  if (!auto && !plan.danger && controller.documentUi &&
      plan.steps.length === 1 && plan.steps[0].op === 'setText') {
    const step = plan.steps[0], i = step.target.paragraph;
    const original = step.find ?? step.before;
    if (original && Number.isInteger(i) && !doc.paragraphs[i]?.table) {
      const start = step.find ? Array.from(step.before.slice(0,step.before.indexOf(step.find))).length : 0;
      const result = await controller.agent('rewrite', { id: command.editId, kind:'docx',
        anchor:{ paragraphs:[i,i],start,end:start+Array.from(original).length },original,text:step.replace ?? step.text },signal);
      record?.({ awaiting:result.pendingReview === true,applied:result.applied });
      return result;
    }
  }
  if (auto) bridge.generating(sessionId, command, controller);
  else {
    record?.({ awaiting: true });
    let approved;
    try { approved = await controller.propose({ kind: 'document', summary: command.summary, rows: describeDocxSteps(plan.steps), danger: plan.danger, reasons: plan.reasons }, signal); }
    finally { record?.({ awaiting: false }); }
    if (!approved) return { declined: true };
  }
  try {
    const result = await controller.agent('apply', { ops: command.ops, editId: command.editId, ...(typeof command.readToken === 'string' && { readToken: command.readToken }) }, signal);
    record?.({ applied: result?.applied });
    return result;
  } finally { if (auto) bridge.generating(sessionId, command, controller, { idle: true }); }
}
