/** Points the 「本次更新」 card shows before 展开全部. */
export const WHATS_NEW_VISIBLE = 4;

/** How long the card stays on screen before it counts as seen without being closed. */
export const WHATS_NEW_SEEN_MS = 8_000;

/** Failed asks (a session still starting, a dropped connection) before a load gives up and shows nothing. */
export const WHATS_NEW_TRIES = 3;

/**
 * The 「本次更新」 notes for every open tab, asked through `fetch()` (any tab's session) until the Host answers once per
 * plugin load. `fetch` answers undefined while no tab can ask yet, so the next tab asks again; a failed ask is tried
 * again by the next tab up to WHATS_NEW_TRIES times, and an older Host without the method only ever shows nothing.
 * The Host records the version as seen once, when the user closes the card (`dismiss()`, every tab) or after a card has
 * stayed on screen for `seenMs`: `shown()` starts the clock and returns the call that stops it when that card goes away,
 * so a card that only flashes past (its tab reloading or closing, a menu covering it) is not lost.
 */
export function createWhatsNewWatch({ fetch, mark, tries = WHATS_NEW_TRIES, seenMs = WHATS_NEW_SEEN_MS, scheduler = globalThis }) {
  const listeners = new Set();
  let value, asking, answered = false, marked = false, failures = 0, timer, onScreen = 0;
  const markNow = () => {
    if (!value || marked) return;
    marked = true;
    scheduler.clearTimeout(timer);
    const { current } = value;
    void Promise.resolve().then(() => mark(current)).catch(() => {});
  };
  const publish = next => { value = next; for (const listener of listeners) listener(); };
  return {
    snapshot: () => value,
    subscribe(listener) { listeners.add(listener); return () => listeners.delete(listener); },
    check() {
      if (answered) return Promise.resolve();
      asking ??= Promise.resolve().then(fetch).then(next => {
        if (next === undefined) return;
        answered = true;
        if (next?.entries?.length) publish(next);
      }, () => { answered = ++failures >= tries; }).finally(() => { asking = undefined; });
      return asking;
    },
    shown() {
      if (!value || marked) return () => {};
      onScreen++;
      timer ??= scheduler.setTimeout(markNow, seenMs);
      let gone = false;
      return () => {
        if (gone) return;
        gone = true;
        if (--onScreen === 0 && timer !== undefined && !marked) { scheduler.clearTimeout(timer); timer = undefined; }
      };
    },
    dismiss() { if (!value) return; markNow(); publish(undefined); },
  };
}

/** The card's language from DSH's interface language (`t('lang')`, as the rest of the tab): English or Chinese. */
export const notesLanguage = lang => String(lang ?? '').toLowerCase().startsWith('en') ? 'en' : 'zh';

/** One version's points in `lang`: English when the notes have them for that version, else the Chinese ones. */
const pointsOf = (entry, lang) => (lang === 'en' && Array.isArray(entry.highlightsEn) && entry.highlightsEn.length ? entry.highlightsEn : entry.highlights);

/**
 * What the card shows: its title (message key and values), the points grouped by version (a version label only when
 * there are several), cut to `limit` points unless expanded. `lang` (notesLanguage) picks each version's English points.
 */
export function whatsNewView(notes, expanded = false, limit = WHATS_NEW_VISIBLE, lang = 'zh') {
  const entries = notes.entries.map(entry => ({ version: entry.version, highlights: pointsOf(entry, notesLanguage(lang)) }));
  const total = entries.reduce((sum, entry) => sum + entry.highlights.length, 0);
  let left = expanded ? total : limit;
  const groups = [];
  for (const entry of entries) {
    if (left <= 0) break;
    const highlights = entry.highlights.slice(0, left);
    left -= highlights.length;
    groups.push({ version: entry.version, highlights });
  }
  const shown = groups.reduce((sum, group) => sum + group.highlights.length, 0);
  const title = notes.previous ? 'whatsNewRange' : 'whatsNewFresh';
  return { title, values: { current: notes.current, previous: notes.previous }, groups, grouped: notes.entries.length > 1, total, hidden: total - shown };
}

/** `{name}` placeholders filled from `values`. */
export const fill = (text, values) => String(text ?? '').replace(/\{(\w+)\}/g, (match, key) => values[key] ?? match);
