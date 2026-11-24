import { randomUUID } from 'node:crypto';

/**
 * Style recommendations the model makes with presentation_suggest_styles, one per Session, in memory (a Host restart
 * drops them; the model can recommend again). The input box's recommendation strip (conversation.input.dock) asks for
 * the Session's open one while the Session is idle; that marks it shown. A recommendation ends when:
 *   - the model recommends again (replaced) or creates a presentation in that Session;
 *   - the Client dismisses it (the user picked, closed the strip, or a new turn started under the strip);
 *   - a model step runs after it was shown: the user has answered in a new turn, whatever the Client saw. That step's
 *     runtime context still lists it (`contextFor`), so an answer such as "the second one" can be read.
 */
export const SUGGEST_MIN = 1;
export const SUGGEST_MAX = 4;
const REASON_MAX = 140;
const TOPIC_MAX = 80;
/** How long a recommendation answered in a new turn stays in the model's runtime context. */
export const ANSWERED_CONTEXT_MS = 10 * 60_000;

const clip = (text, size) => {
  const value = typeof text === 'string' ? text.replace(/\s+/g, ' ').trim() : '';
  return value.length > size ? `${value.slice(0, size - 1)}…` : value;
};

/**
 * The model's candidates against the packs there are: known names only, each once, at most SUGGEST_MAX, in its order.
 * Returns `{ candidates: [{ style, reason }], dropped: [names] }`.
 */
export function normalizeCandidates(candidates, available) {
  const known = new Set(available);
  const seen = new Set();
  const kept = [], dropped = [];
  for (const item of Array.isArray(candidates) ? candidates : []) {
    const style = typeof item?.style === 'string' ? item.style.trim() : '';
    if (!known.has(style)) { if (style) dropped.push(style); continue; }
    if (seen.has(style) || kept.length >= SUGGEST_MAX) continue;
    seen.add(style);
    kept.push({ style, reason: clip(item.reason, REASON_MAX) });
  }
  return { candidates: kept, dropped };
}

export class StyleSuggestions {
  constructor({ now = () => Date.now(), newId = randomUUID } = {}) {
    Object.assign(this, { now, newId });
    /** sessionId → { id, topic, candidates, created, state: 'open' | 'shown' | 'answered', answeredAt? } */
    this.bySession = new Map();
  }
  /** Record the Session's recommendation, replacing any earlier one. Throws when no candidate names a pack. */
  propose(sessionId, { candidates, topic } = {}, available = []) {
    const normalized = normalizeCandidates(candidates, available);
    if (normalized.candidates.length < SUGGEST_MIN) {
      const named = normalized.dropped.length ? ` (${normalized.dropped.map(name => JSON.stringify(name)).join(', ')} are not style packs)` : '';
      throw Object.assign(new Error(`No candidate names a style pack${named}; available: ${available.join(', ') || 'none'}.`), { code: 'invalid-edit' });
    }
    const suggestion = { id: this.newId(), topic: clip(topic, TOPIC_MAX), candidates: normalized.candidates, created: this.now(), state: 'open' };
    this.bySession.set(sessionId, suggestion);
    return { suggestion, dropped: normalized.dropped };
  }
  /** The open recommendation for the strip, marking it shown; undefined once answered or ended. */
  show(sessionId) {
    const suggestion = this.bySession.get(sessionId);
    if (!suggestion || suggestion.state === 'answered') return undefined;
    suggestion.state = 'shown';
    return suggestion;
  }
  /** End the Session's recommendation; with `id`, only when it is still that one. Returns whether one ended. */
  dismiss(sessionId, id) {
    const suggestion = this.bySession.get(sessionId);
    if (!suggestion || (id !== undefined && id !== null && suggestion.id !== id)) return false;
    this.bySession.delete(sessionId);
    return true;
  }
  /**
   * Called on every model step of the Session (runtime context). A shown recommendation means this step follows the
   * user's answer: it turns answered (no strip any more) and stays in context for ANSWERED_CONTEXT_MS. Returns the
   * recommendation the step's context should list, if any.
   */
  forStep(sessionId) {
    const suggestion = this.bySession.get(sessionId);
    if (!suggestion) return undefined;
    if (suggestion.state === 'shown') Object.assign(suggestion, { state: 'answered', answeredAt: this.now() });
    if (suggestion.state !== 'answered') return undefined;
    if (this.now() - suggestion.answeredAt > ANSWERED_CONTEXT_MS) { this.bySession.delete(sessionId); return undefined; }
    return suggestion;
  }
}

/** The runtime-context lines naming an answered recommendation; `nameOf(style)` gives a pack's display name. */
export function suggestionContext(suggestion, nameOf = style => style) {
  if (!suggestion) return '';
  const list = suggestion.candidates.map((item, index) => `${index + 1}) ${item.style}${nameOf(item.style) && nameOf(item.style) !== item.style ? ` (${nameOf(item.style)})` : ''}`).join(', ');
  return `PPT styles you recommended last turn (presentation_suggest_styles): ${list}. If the user's reply picks one by number, name or alias, use that pack; "use the default style" means the workspace default; anything else, ask again rather than choosing.`;
}
