/**
 * The one style card and gallery every PPT style surface uses: the input box's "PPT style" popover, the first-screen
 * template gallery and the recommendation strip above the input box. Cards come from the Host's listStyles /
 * styleSuggestion (style-catalog.mjs styleCard: thumbnail and page previews as data URIs). Stateless: the surface
 * passes the filter it keeps and what a pick does. Styling: the plugin's own `sbs-` classes (styles.mjs).
 */

/** A card's display name in `lang`, falling back to the other language, then the id. */
export const styleTitle = (card, lang = 'zh') => card?.name?.[lang] || card?.name?.zh || card?.name?.en || card?.id || '';

const localizedOf = (value, lang) => value?.[lang] || value?.zh || value?.en || '';
const scenariosOf = (card, lang) => {
  const list = card?.scenarios?.[lang]?.length ? card.scenarios[lang] : card?.scenarios?.zh?.length ? card.scenarios.zh : card?.scenarios?.en;
  return Array.isArray(list) ? list.filter(item => typeof item === 'string' && item) : [];
};

export const FILTER_ALL = '*';
const FILTERS_MAX = 8;

/** Scenario filters for these cards: the most common first (ties keep first-seen order), at most eight. */
export function galleryFilters(cards, lang = 'zh') {
  const counts = new Map();
  for (const card of cards ?? []) for (const item of new Set(scenariosOf(card, lang))) counts.set(item, (counts.get(item) ?? 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, FILTERS_MAX).map(([name]) => name);
}

/** The cards a filter keeps; the all-filter (or one no card has any more) keeps every card. */
export function filterCards(cards, filter, lang = 'zh') {
  const list = cards ?? [];
  if (!filter || filter === FILTER_ALL) return list;
  const kept = list.filter(card => scenariosOf(card, lang).includes(filter));
  return kept.length ? kept : list;
}

/** Arrow keys between cards in a grid of `columns`; Home/End jump. Undefined for other keys. */
export function gridKey(key, index, count, columns = 2) {
  if (!count) return undefined;
  const move = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: columns, ArrowUp: -columns }[key];
  if (move !== undefined) return Math.max(0, Math.min(count - 1, index + move));
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return undefined;
}

/** Columns of a laid-out grid: the cards on the first card's row (2 when there is no layout to measure). */
export function columnsOf(cards) {
  const top = cards[0]?.offsetTop;
  if (typeof top !== 'number') return 2;
  const n = cards.filter(card => card.offsetTop === top).length;
  return n > 0 ? n : 2;
}

/** Move the keyboard between the `.sbs-card` buttons of `container` for an arrow / Home / End key event. */
export function moveFocus(event, container, doc) {
  const cards = [...(container?.querySelectorAll?.('.sbs-card') ?? [])];
  const next = gridKey(event.key, cards.indexOf(doc?.activeElement), cards.length, columnsOf(cards));
  if (next === undefined) return false;
  event.preventDefault?.();
  cards[next]?.focus?.();
  return true;
}

const format = (t, key, params = {}) => Object.entries(params).reduce((text, [name, value]) => text.split(`{${name}}`).join(String(value)), String((typeof t === 'function' ? t(key) : undefined) ?? key));

/** Card and gallery components on the host's React. */
export function createStyleCards(React) {
  const { createElement: h } = React;

  /**
   * One style. `full`: thumbnail, page previews, name, id, badge, one-line summary, scenarios, prerequisite.
   * `compact`: thumbnail, name and the reason it was recommended.
   */
  function StyleCard({ card, lang, t, variant = 'full', selected = false, badge, reason, disabled, onPick }) {
    const name = styleTitle(card, lang);
    const thumb = card.thumb ? h('img', { className: 'sbs-thumb', src: card.thumb, alt: '', draggable: false })
      : h('span', { className: 'sbs-thumb sbs-swatch', style: { background: card.primary ?? '#6b7178' }, 'aria-hidden': true }, name);
    const common = { type: 'button', role: 'option', 'aria-selected': selected, className: `sbs-card sbs-card-${variant}`, 'data-style': card.id, disabled, onClick: () => onPick?.(card) };
    if (variant === 'compact') {
      const why = reason ?? card.reason;
      return h('button', common, thumb,
        h('span', { className: 'sbs-name' }, h('span', { className: 'sbs-title' }, name)),
        why ? h('span', { className: 'sbs-meta sbs-reason' }, why) : null);
    }
    const previews = Array.isArray(card.previews) ? card.previews.filter(Boolean) : [];
    const summary = localizedOf(card.summary, lang);
    const scenarios = scenariosOf(card, lang);
    const requires = localizedOf(card.requires?.note, lang);
    return h('button', common, thumb,
      previews.length ? h('span', { className: 'sbs-previews', 'aria-hidden': true }, previews.map((src, index) => h('img', { key: index, className: 'sbs-preview', src, alt: '', draggable: false }))) : null,
      h('span', { className: 'sbs-name' }, h('span', { className: 'sbs-title' }, name), h('span', { className: 'sbs-id' }, card.id),
        badge ? h('span', { className: 'sbs-badge' }, badge) : null),
      summary ? h('span', { className: 'sbs-meta' }, summary) : null,
      scenarios.length ? h('span', { className: 'sbs-meta sbs-scenarios' }, format(t, 'scenarios', { list: scenarios.join(lang === 'en' ? ', ' : '、') })) : null,
      requires ? h('span', { className: 'sbs-meta sbs-requires' }, requires) : null);
  }

  /**
   * Scenario filter chips (when there are two cards or more and any scenario) above a grid of cards. `filter` and
   * `onFilter` belong to the surface; `selectedId` marks a card, `badgeId` gets `badge`.
   */
  function StyleGallery({ cards, lang, t, label, filter = FILTER_ALL, onFilter, selectedId, badgeId, badge, busy, onPick, onKeyDown, gridRef, variant = 'full' }) {
    const filters = (cards?.length ?? 0) >= 2 ? galleryFilters(cards, lang) : [];
    const active = filters.includes(filter) ? filter : FILTER_ALL;
    const shown = filterCards(cards, active, lang);
    const chip = (value, text) => h('button', { key: value, type: 'button', className: 'sbs-chip', 'aria-pressed': active === value, onClick: () => onFilter?.(value) }, text);
    return h('div', { className: 'sbs-gallery' },
      filters.length && onFilter ? h('div', { className: 'sbs-filters', role: 'toolbar', 'aria-label': format(t, 'filtersSuggest') }, chip(FILTER_ALL, format(t, 'all')), filters.map(item => chip(item, item))) : null,
      h('div', { ref: gridRef, className: 'sbs-grid', role: 'listbox', 'aria-label': label, onKeyDown },
        shown.map(card => h(StyleCard, { key: card.id, card, lang, t, variant, selected: selectedId === card.id, badge: badgeId === card.id ? badge : undefined, disabled: busy, onPick }))));
  }

  return { StyleCard, StyleGallery };
}
