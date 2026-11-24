import { randomBytes } from 'node:crypto';
import { SuiteError } from '../../../core/src/index.mjs';

/**
 * Selections the user asked to rewrite (spec 261109-selection-ai-edit §3, §4.3, §5), per Session, until the model's
 * selection_rewrite uses one. A selection is one-shot: `pending` → `applying` → `used`; the user can cancel it while it
 * waits, a document_edit landing on the same file makes the waiting ones stale, and after half an hour it is gone.
 * Kept in memory only: a Host restart forgets them, and the model is told the id is unknown.
 */
export const SELECTION_STATES = Object.freeze(['pending', 'applying', 'used', 'cancelled', 'stale', 'reverted', 'kept']);
/** What the user may report about their own selection from the page. */
export const USER_STATES = Object.freeze(['cancelled', 'kept', 'reverted']);

export class SelectionStore {
  constructor({ ttlMs = 30 * 60_000, maxPerSession = 20, now = () => Date.now(), random = () => randomBytes(4).toString('hex') } = {}) {
    Object.assign(this, { ttlMs, maxPerSession, now, random });
    this.bySession = new Map();
  }

  sweep(sessionId) {
    const records = this.bySession.get(sessionId);
    if (!records) return new Map();
    for (const [id, record] of records) if (this.now() - record.createdAt > this.ttlMs) records.delete(id);
    return records;
  }

  /** A new pending selection; the oldest ones go beyond `maxPerSession`. */
  create(sessionId, { path, realPath, name, request, lang = 'zh' }) {
    const records = this.sweep(sessionId);
    this.bySession.set(sessionId, records);
    let id;
    do id = `sel-${this.random()}`; while (records.has(id));
    const record = { id, sessionId, path, realPath, name, request, lang, state: 'pending', createdAt: this.now() };
    records.set(id, record);
    while (records.size > this.maxPerSession) records.delete(records.keys().next().value);
    return record;
  }

  /** The Session's selection `id`, or `selection-unknown` (another Session's, expired, or never made). */
  get(sessionId, id) {
    const record = typeof id === 'string' ? this.sweep(sessionId).get(id) : undefined;
    if (!record) throw new SuiteError('selection-unknown', `There is no selection ${JSON.stringify(id)} in this conversation (it may have expired after 30 minutes). Ask the user to select the text again and send the request from the editor.`);
    return record;
  }

  /**
   * Takes a pending selection for the model's rewrite. A selection rewritten once, cancelled by the user or changed by
   * another edit cannot be taken again.
   */
  take(sessionId, id) {
    const record = this.get(sessionId, id);
    if (record.state === 'cancelled') throw new SuiteError('cancelled', 'The user cancelled this change in the editor. Do not change the document; tell them it was not changed.');
    if (record.state === 'stale') throw new SuiteError('selection-stale', 'The document changed since the user selected this text. Ask them to select it again and send the request from the editor.');
    if (record.state === 'applying') throw new SuiteError('busy', 'This selection is being rewritten right now.');
    if (record.state !== 'pending') throw new SuiteError('selection-used', 'This selection was already rewritten once. To change it again, ask the user to select the text again (or use document_edit / spreadsheet_edit, which asks them first).');
    record.state = 'applying';
    return record;
  }

  set(sessionId, id, state) {
    const record = this.get(sessionId, id);
    record.state = state;
    return record;
  }

  /** The user's own report from the page: cancel a waiting selection, or keep / revert a rewritten one. */
  report(sessionId, id, state) {
    if (!USER_STATES.includes(state)) throw new SuiteError('invalid-edit', `state must be one of ${USER_STATES.join(', ')}.`);
    const record = this.get(sessionId, id);
    if (state === 'cancelled' ? record.state === 'pending' : ['used', 'kept'].includes(record.state)) record.state = state;
    return { id, state: record.state };
  }

  /** Another edit landed on `realPath`: the selections still waiting there no longer hold what the user selected. */
  staleFor(sessionId, realPath) {
    let count = 0;
    for (const record of this.sweep(sessionId).values()) if (record.state === 'pending' && !record.tracked && record.realPath === realPath) { record.state = 'stale'; count++; }
    return count;
  }
}
