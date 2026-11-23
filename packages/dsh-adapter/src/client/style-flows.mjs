import { createStyleCards, moveFocus, styleTitle } from './style-cards.mjs';
import { STYLE_PICKER_NS, focusComposer, format, languageOf } from './style-picker.mjs';

/**
 * What shows right above the input box (`conversation.input.dock`): after the model called presentation_suggest_styles,
 * its 2–4 picks with the reason for each, "none of these, use the default" and "see all styles", while the Session is
 * idle. The inspiration library itself is the full page behind the sidebar entry (inspiration-library.mjs, handoff
 * v2.0): the input box no longer has a library button, panel or first-screen chip.
 */
export const STYLE_DOCK_ID = 'officesdk-ppt-dock';

/** Whether the Session is running a turn: DSH's session status, else the owner's Session snapshot; undefined if unknown. */
export function runningOf({ sessionId, useSessionStatus, session }) {
  const status = useSessionStatus?.(statuses => statuses?.get?.(sessionId)?.running);
  if (typeof status === 'boolean') return status;
  return typeof session?.running === 'boolean' ? session.running : undefined;
}

/** What the dock shows: the recommendation while idle, else nothing. */
export function dockMode(state, running) {
  if (state.suggestion && running !== true) return state.suggestView === 'all' ? 'suggest-all' : 'suggest';
  return 'none';
}

export function createStyleFlows({ React, stores, document: doc = globalThis.document, defer }) {
  const { createElement: h, useEffect, useRef, useSyncExternalStore } = React;
  const { StyleGallery } = createStyleCards(React);
  const useStore = sessionId => {
    const store = stores.get(sessionId);
    return [store, useSyncExternalStore(listener => store.subscribe(listener), () => store.snapshot())];
  };

  /** `conversation.input.dock`: the recommendation strip. */
  function StyleDock({ sessionId, inputActions, useInput, useSessionStatus, session, input: zoneInput, t }) {
    const [store, state] = useStore(sessionId);
    const lang = languageOf(t);
    const input = useInput?.(value => value) ?? zoneInput;
    const running = runningOf({ sessionId, useSessionStatus, session });
    const ref = useRef(null);
    const wasRunning = useRef(running);
    // Idle (or unknown): ask for the recommendation. A turn that starts while one is shown is the user answering.
    // A turn starting (the user sent) takes the template prefix away; when it ends the Host forgets the pick.
    useEffect(() => {
      const before = wasRunning.current;
      wasRunning.current = running;
      if (running === true) {
        // Mounted mid-turn (a reload, a new pane): whatever pick the Host holds belongs to this turn and is cleared when
        // it ends; picks made now wait for that.
        if (!store.running && before === true) store.clearAfterTurn = true;
        store.running = true;
        if (before === false && store.state.suggestion) void store.dismissSuggestion();
        if (before !== true) store.turnStarted();
        return;
      }
      if (before === true || store.running) void store.turnEnded();
      void store.fetchSuggestion();
    }, [store, running]);
    const mode = dockMode(state, running);
    if (mode === 'none' || !inputActions) return null;
    const after = () => focusComposer(ref.current);

    const onKeyDown = event => { moveFocus(event, ref.current, doc); };
    const cards = state.list?.styles ?? [];
    const loading = state.phase === 'error' && !state.list
      ? h('div', { className: 'sbs-state', role: 'alert' }, format(t, 'failed', { message: state.error ?? '' }), h('button', { type: 'button', className: 'sbs-link', onClick: () => store.load() }, format(t, 'retry')))
      : !state.list ? h('div', { className: 'sbs-state', role: 'status' }, format(t, 'loading')) : null;
    const suggestion = state.suggestion;
    const answer = (card, fallback = false) => {
      const how = store.answer(card, { actions: inputActions, input, lang, fallback, running: running === true, defer });
      if (how !== 'sent') after();
    };
    const fallbackCard = suggestion.defaultStyle ?? { id: 'default', name: { zh: 'default', en: 'default' } };
    const head = h('div', { className: 'sbs-head' },
      h('span', { className: 'sbs-heading' }, format(t, 'suggestTitle')),
      suggestion.topic ? h('span', { className: 'sbs-hint' }, format(t, 'suggestTopic', { topic: suggestion.topic })) : h('span', { className: 'sbs-hint' }),
      h('button', { type: 'button', className: 'sbs-x', 'aria-label': format(t, 'suggestClose'), title: format(t, 'suggestClose'), onClick: () => { void store.dismissSuggestion(); } }, '×'));
    const body = mode === 'suggest-all'
      ? (loading ?? h(StyleGallery, { cards, lang, t, label: format(t, 'suggestAll'), filter: state.filter, onFilter: filter => store.setFilter(filter), onPick: card => answer(card), onKeyDown }))
      // Compact cards: thumbnail, name and the model's reason for each.
      : h(StyleGallery, { cards: suggestion.candidates ?? [], lang, t, variant: 'compact', label: format(t, 'suggestTitle'), onPick: card => answer(card), onKeyDown });
    const actions = h('div', { className: 'sbs-actions' },
      h('button', { type: 'button', className: 'sbs-action', 'data-shimo-style-default': true, onClick: () => answer(fallbackCard, true) }, format(t, 'suggestDefault', { name: styleTitle(fallbackCard, lang) })),
      mode === 'suggest-all'
        ? h('button', { type: 'button', className: 'sbs-action', onClick: () => store.setSuggestView('cards') }, format(t, 'suggestBack'))
        : h('button', { type: 'button', className: 'sbs-action', 'data-shimo-style-all': true, onClick: () => { void store.setSuggestView('all'); } }, format(t, 'suggestAll')));
    return h('div', { ref, className: 'sbs-dock sbs-dock-suggest', role: 'region', 'aria-label': format(t, 'suggestTitle'), 'data-shimo-style-dock': 'suggest' }, head, body, actions);
  }

  return { StyleDock };
}

/** Registers the recommendation dock beside the input box's prefix, sharing its per-Session stores and dictionary. */
export function registerStyleFlows(ctx, React, { stores, document, defer } = {}) {
  const { StyleDock } = createStyleFlows({ React, stores, document, defer });
  // After DSH's own docks (queue, to-do), which register at the default order.
  ctx.effect(() => ctx.slots.inject('conversation.input.dock', () => ctx.slots.register({ name: 'conversation.input.dock', id: STYLE_DOCK_ID, order: 100, locale: STYLE_PICKER_NS }, StyleDock)));
  return { StyleDock };
}
