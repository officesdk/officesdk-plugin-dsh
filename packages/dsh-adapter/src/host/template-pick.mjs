/**
 * The inspiration library's pick, one per Session, in memory (a Host restart drops it; the Client then shows no tag).
 * The input box's panel names a Word, Excel or PPT template in the draft and records the pick here, so each model step
 * of the turn that follows gets it as a structured line in its runtime context (`pickContext`). A pick ends when:
 *   - the Client clears it (the user removed the tag, or the turn that carried it ended);
 *   - PICK_CONTEXT_MS after the first model step that read it, whatever the Client did (a closed window, a reload).
 */
export const PICK_KINDS = Object.freeze(['word', 'excel', 'ppt']);
/** Built-in Word and Excel templates: the live templates the bundled skills ship (skills/<skill>/templates): Word gongwen, no Excel one. */
export const BUILTIN_TEMPLATES = Object.freeze({
  word: Object.freeze([
    { id: 'gongwen', name: { zh: '公文版式', en: 'Official document' } },
  ]),
  excel: Object.freeze([]),
});
/** How long a pick stays in the runtime context after the first step that read it. */
export const PICK_CONTEXT_MS = 20 * 60_000;

export class TemplatePicks {
  constructor({ now = () => Date.now() } = {}) {
    this.now = now;
    /** sessionId → { kind, id, setAt, usedAt? } */
    this.bySession = new Map();
  }
  /** Record (or with null clear) the Session's pick; returns what is recorded now. */
  set(sessionId, pick) {
    if (!pick) { this.bySession.delete(sessionId); return null; }
    this.bySession.set(sessionId, { kind: pick.kind, id: pick.id, setAt: this.now() });
    return { kind: pick.kind, id: pick.id };
  }
  /** The Session's pick as the Client shows it, or null once it ended. */
  get(sessionId) {
    const pick = this.live(sessionId);
    return pick ? { kind: pick.kind, id: pick.id } : null;
  }
  /** Called on every model step: the first read starts the pick's context window. Returns the pick or undefined. */
  forStep(sessionId) {
    const pick = this.live(sessionId);
    if (pick && pick.usedAt === undefined) pick.usedAt = this.now();
    return pick;
  }
  live(sessionId) {
    const pick = this.bySession.get(sessionId);
    if (!pick) return undefined;
    if (pick.usedAt !== undefined && this.now() - pick.usedAt > PICK_CONTEXT_MS) { this.bySession.delete(sessionId); return undefined; }
    return pick;
  }
}

/**
 * The runtime-context line naming the pick; `nameOf(pick)` gives its display name. It counts as the user naming the
 * template, so it wins over the workspace default line above it.
 */
export function pickContext(pick, nameOf = () => '') {
  if (!pick) return '';
  const name = nameOf(pick);
  const named = name && name !== pick.id ? ` (${name})` : '';
  const lead = 'Inspiration library pick: the user picked this template in the input box for the file they ask for now; it counts as the user naming it and overrides any workspace default.';
  if (pick.kind === 'ppt') return `${lead} PPT style pack "${pick.id}"${named}: pass style "${pick.id}" to presentation_create (offline: --style ${pick.id}).`;
  if (pick.kind === 'word') return `${lead} Word template "${pick.id}"${named}: pass template "${pick.id}" to document_create.`;
  return `${lead} Excel template "${pick.id}"${named}: create the workbook with spreadsheet_create (its template is ${pick.id}).`;
}
