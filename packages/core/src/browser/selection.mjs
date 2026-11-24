/**
 * The v1.1 selection gesture every agent target goes through before anything is written (handoff
 * first-generation interaction handoff v1.1, 01/02; spec 261105 §7.2, §7.6): the cursor appears at the target's top-left,
 * then pulls the selection open to the bottom-right (width and height together, smoothstep), and only then does the
 * real write start. Only these two parts are timed; the rest follows the real writes. Reduced motion skips both.
 */
export const SELECTION_TIMING = Object.freeze({ leadMs: 150, dragMs: 500, tailMs: 200 });

// Tracks are appended independently. Every cursor (arrow and label) paints above every selection.
export const AGENT_LAYER = Object.freeze({ selection: '2147483646', cursor: '2147483647' });

export const smoothstep = value => { const p = Math.max(0, Math.min(1, value)); return p * p * (3 - 2 * p); };

/**
 * Where a selection begun at `start` is at `now`: `cursor` (only the cursor, at the top-left), `selecting` (the outline
 * grows; `progress` 0–1, eased) or `done` (full outline). `animate: false` is always `done`.
 */
export function selectionAt(start, now, { animate = true, leadMs = SELECTION_TIMING.leadMs, dragMs = SELECTION_TIMING.dragMs } = {}) {
  if (!animate || start === undefined) return { phase: 'done', progress: 1 };
  const age = now - start;
  if (age < leadMs) return { phase: 'cursor', progress: 0 };
  if (age < leadMs + dragMs) return { phase: 'selecting', progress: smoothstep((age - leadMs) / dragMs) };
  return { phase: 'done', progress: 1 };
}

/** How long a selection begun now takes before writing may start. */
export const selectionMs = ({ animate = true } = {}) => animate ? SELECTION_TIMING.leadMs + SELECTION_TIMING.dragMs : 0;

/**
 * A text written in up to `max` real writes (spec §7.6: at most 4, each at least `min` characters): the prefixes to
 * write one after another, the last one the whole text. Short texts are written at once.
 */
export function textChunks(text, { max = 4, min = 12 } = {}) {
  const value = String(text ?? '');
  const parts = Math.max(1, Math.min(max, Math.floor(value.length / min)));
  if (parts === 1) return [value];
  const chars = Array.from(value);
  return Array.from({ length: parts }, (_, index) => chars.slice(0, Math.round(chars.length * (index + 1) / parts)).join(''));
}
