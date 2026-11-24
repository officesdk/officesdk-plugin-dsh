import { STYLE_PICKER_NS, bulb, format, languageOf } from './style-picker.mjs';
import { styleTitle } from './style-cards.mjs';
import {
  CARD_WIDTH, FRAME_RATIO, HOVER_DELAY_MS, HOVER_EVERY_MS, HOVER_PAGES, KINDS, KIND_LABEL, SCOPE_FEATURED, SCOPE_MINE, UPLOAD_EXTENSIONS,
  categoriesOf, checkUpload, columnCount, featurePage, isKind, kindsOf, libraryTemplates, masonry, modeKey, pageKey, scopeTemplates, summaryLines,
} from './inspiration.mjs';
import { FILE_ICONS, OUTLINE_ICONS } from './inspiration-icons.mjs';

/**
 * The inspiration library page (design handoff OfficeSDK-DSH-Inspiration-v2.0, docs/03-最终交互规格.md), on DSH's
 * public contracts (checked in the 0.2.1-rc.3 bundle): a `sidebar.panellist` entry 「灵感库」 between 自动化任务
 * (order 10) and 更多, and a `main` entry under the same key that DSH shows while that panel is selected
 * (`layout.selectPanel(id)`; null returns to the conversation). The page has the centred title, the Word / Excel / PPT
 * glass switch (it sets the creation mode only), the 「我的模板｜精选｜场景」 filters over every kind mixed, the
 * fixed-width masonry, the template detail and the page lightbox.
 *
 * DSH does not let a plugin page hold its real input box (slot names are declared once; the composer chain replaces
 * the composer), and the user sends with DSH's own button: so 「添加并使用」 records a pending use for the Session the
 * user came from and goes back to it; that Session's input-box slot (style-picker.mjs) binds the template on the Host,
 * writes the kind's prompt into DSH's draft and focuses it. After the user sends, the AI's file opens in the existing
 * editor in DSH's right sidebar (AgentBridge), which DSH slides in leftwards while the chat narrows.
 * Styling: the plugin's own `sbi-` classes (styles.mjs); colours follow DSH's theme, never set it.
 */
export const INSPIRATION_PANEL_ID = 'officesdk-inspiration';
/** Between DSH's 自动化任务 (10) and the More launcher, which always comes last; 插件 is 0. */
export const INSPIRATION_PANEL_ORDER = 20;
/** How long a pending use waits for a Session that had none (a new conversation) to mount its input box. */
export const PENDING_MS = 30_000;

const emptyUploads = () => ({ word: { state: 'idle' }, excel: { state: 'idle' }, ppt: { state: 'idle' } });

/**
 * The page's state, one per window: the creation mode, the filter, the open detail, "my templates" entries, and the
 * pending use (`{ template, target, at }`, `target` = the Session the user came from, or null when there was none).
 * Mode and filter survive closing and reopening the page; drafts and bindings live with each Session.
 */
export class InspirationLibrary {
  constructor({ now = () => Date.now(), currentSession = () => undefined } = {}) {
    this.now = now;
    // The Session DSH's main view shows: a use with no target lands only there (not in another pane's input box).
    this.currentSession = currentSession;
    this.state = { mode: 'ppt', scope: SCOPE_FEATURED, detail: null, pending: null, uploads: emptyUploads(), dragKind: null };
    this.listeners = new Set();
  }
  subscribe(listener) { this.listeners.add(listener); return () => this.listeners.delete(listener); }
  snapshot() { return this.state; }
  set(patch) { this.state = { ...this.state, ...patch }; for (const listener of this.listeners) listener(); }
  /** The mode switch: only the creation mode; the list and the filter stay as they are (§2). */
  setMode(mode) { if (isKind(mode) && mode !== this.state.mode) this.set({ mode }); }
  setScope(scope) { const next = scope || SCOPE_FEATURED; if (next !== this.state.scope) this.set({ scope: next }); }
  /** A card's body: the detail, nothing else changes (AC-026). */
  openDetail(template) { this.set({ detail: template }); }
  closeDetail() { if (this.state.detail) this.set({ detail: null }); }
  /**
   * 「添加并使用」 (cards and detail alike): the mode follows the template's kind, the detail closes, and the use waits
   * for `target`'s input box (null: the next Session that mounts one within PENDING_MS).
   */
  use(template, { target = null } = {}) {
    this.set({ mode: template.kind, detail: null, pending: { template, target: target ?? null, at: this.now() } });
  }
  /** The pending use for `sessionId`, taken once; an expired one is dropped; one meant for another Session stays. */
  takePending(sessionId) {
    const pending = this.state.pending;
    if (!pending) return undefined;
    if (this.now() - pending.at > PENDING_MS) { this.set({ pending: null }); return undefined; }
    if (pending.target && pending.target !== sessionId) return undefined;
    if (!pending.target) { const current = this.currentSession?.(); if (current && current !== sessionId) return undefined; }
    this.set({ pending: null });
    return pending.template;
  }
  /**
   * A file chosen or dropped on "my templates" of `kind`: checked by name, size and count only. A good one ends at
   * "not connected": there is no upload or making service yet, so nothing is uploaded, no progress shown, nothing saved.
   */
  receiveFiles(kind, files) {
    const checked = checkUpload(files, kind);
    const entry = checked.error ? { state: 'error', error: checked.error } : { state: 'unavailable', fileName: checked.file.name };
    this.set({ uploads: { ...this.state.uploads, [kind]: entry }, dragKind: null });
    return entry;
  }
  resetUpload(kind) { this.set({ uploads: { ...this.state.uploads, [kind]: { state: 'idle' } } }); }
  setDrag(kind) { if (this.state.dragKind !== kind) this.set({ dragKind: kind }); }
}

const NO_STORE = { subscribe: () => () => {}, snapshot: () => undefined, getSnapshot: () => undefined };
const reducedMotion = win => Boolean(win?.matchMedia?.('(prefers-reduced-motion: reduce)')?.matches);
const ICON_PATHS = { close: 'M6 6l12 12M18 6L6 18', left: 'm14 6-6 6 6 6', right: 'm10 6 6 6-6 6', back: 'M15 6l-6 6 6 6', plus: 'M12 5v14M5 12h14' };

export function createInspirationLibrary({ React, stores, library, layout = () => undefined, workspace = () => undefined, document: doc = globalThis.document }) {
  const { createElement: h, Fragment, useEffect, useLayoutEffect = React.useEffect, useRef, useState, useSyncExternalStore } = React;
  const win = () => doc?.defaultView;
  const icon = (name, size = 16) => h('svg', { width: size, height: size, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.6, strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': true }, h('path', { d: ICON_PATHS[name] }));
  const useLibrary = () => useSyncExternalStore(listener => library.subscribe(listener), () => library.snapshot());
  /** The Session DSH's main view holds (the one the user goes back to): `uiWorkspace.selection`. */
  const selectionOf = () => workspace()?.selection ?? NO_STORE;
  const useSessionId = () => {
    const selection = selectionOf();
    return useSyncExternalStore(listener => selection.subscribe?.(listener) ?? (() => {}), () => selection.getSnapshot?.()?.sessionId);
  };
  const close = () => { layout()?.selectPanel?.(null); };

  /** The type mark shared by the switch and every card (§2): 12px icon, 2px gap, 12/14px 400 label. */
  function FileTypeContent({ kind, selected = true }) {
    return h('span', { className: 'sbi-type', 'data-kind': kind },
      h('span', { className: 'sbi-type-icon' }, selected
        ? h('img', { src: FILE_ICONS[kind], width: 12, height: 12, alt: '', draggable: false })
        : h('span', { className: 'sbi-type-outline', style: { '--sbi-outline': `url("${OUTLINE_ICONS[kind]}")` }, 'aria-hidden': true })),
      h('span', { className: 'sbi-type-label' }, KIND_LABEL[kind]));
  }

  /** The Word / Excel / PPT glass switch: 520ms jelly move, a new press cancels the running one, none with reduced motion. */
  function TypeSwitch({ mode, onChange, t, kinds = KINDS }) {
    const index = Math.max(0, kinds.indexOf(mode));
    const slider = useRef(null);
    const previous = useRef(index);
    const motion = useRef(null);
    useLayoutEffect(() => {
      const node = slider.current;
      const from = previous.current;
      previous.current = index;
      if (!node || from === index) return;
      const running = motion.current?.playState === 'running';
      const style = win()?.getComputedStyle?.(node);
      const width = parseFloat(style?.width) || 0;
      const start = running && style?.transform && style.transform !== 'none' ? new (win().DOMMatrixReadOnly)(style.transform).m41 : from * width;
      motion.current?.cancel?.();
      motion.current = null;
      if (reducedMotion(win()) || !node.animate) return;
      const end = index * width;
      const direction = Math.sign(index - from);
      motion.current = node.animate([
        { transform: `translateX(${start}px)`, offset: 0 },
        { transform: `translateX(${start + (end - start) * 0.82}px) scale(1.12,.91)`, offset: 0.44 },
        { transform: `translateX(${end + direction * 3}px) scale(.95,1.055)`, offset: 0.7 },
        { transform: `translateX(${end - direction}px) scale(1.025,.98)`, offset: 0.86 },
        { transform: `translateX(${end}px) scale(1,1)`, offset: 1 },
      ], { duration: 520, easing: 'cubic-bezier(.25,.65,.3,1)' });
    }, [index]);
    useEffect(() => () => motion.current?.cancel?.(), []);
    return h('div', { className: 'sbi-switch', role: 'radiogroup', 'aria-label': format(t, 'modes'), style: { '--sbi-index': index, '--sbi-count': kinds.length }, 'data-shimo-inspiration': 'modes',
      onKeyDown: event => {
        const next = modeKey(event.key, index, kinds.length);
        if (next === undefined) return;
        event.preventDefault?.();
        onChange(kinds[next]);
        event.currentTarget?.querySelectorAll?.('button')?.[next]?.focus?.();
      } },
    h('span', { ref: slider, className: 'sbi-switch-slider', 'aria-hidden': true }),
    ...kinds.map(value => h('button', { key: value, type: 'button', role: 'radio', 'aria-checked': mode === value, tabIndex: mode === value ? 0 : -1, 'data-kind': value, onClick: () => onChange(value) },
      h(FileTypeContent, { kind: value, selected: mode === value }))));
  }

  /** A page image, or the honest missing face (file icon + text); never another page in its place. */
  function PageImage({ src, kind, alt, className, t }) {
    const [failed, setFailed] = useState(false);
    useEffect(() => { setFailed(false); }, [src]);
    if (src && !failed) return h('img', { className, src, alt, draggable: false, onError: () => setFailed(true) });
    return h('span', { className: 'sbi-no-preview', style: { aspectRatio: FRAME_RATIO[kind] } },
      h('img', { src: FILE_ICONS[kind], width: 32, height: 32, alt: '', draggable: false }), h('span', null, format(t, 'noPreview')));
  }

  /**
   * One card: the first page in its kind's frame; hovering 1s turns to page 2, then every 3s over the first three pages
   * (650ms fades); leaving, the page going to the background, an open detail or reduced motion stop it back on page 1.
   * The body opens the detail; 「添加并使用」 (hover or keyboard focus) uses it directly.
   */
  function CoverCard({ item, onPreview, onUse, active, t }) {
    const frames = item.pages.slice(0, HOVER_PAGES);
    const [hovered, setHovered] = useState(false);
    const [page, setPage] = useState(0);
    const [failed, setFailed] = useState({});
    const [visible, setVisible] = useState(() => !doc?.hidden);
    useEffect(() => {
      const onVisibility = () => setVisible(!doc?.hidden);
      doc?.addEventListener?.('visibilitychange', onVisibility);
      return () => doc?.removeEventListener?.('visibilitychange', onVisibility);
    }, []);
    useEffect(() => {
      if (!hovered || !active || !visible || reducedMotion(win()) || frames.length < 2) { setPage(0); return undefined; }
      let timer;
      const delay = setTimeout(() => {
        setPage(1);
        timer = setInterval(() => setPage(current => (current + 1) % frames.length), HOVER_EVERY_MS);
      }, HOVER_DELAY_MS);
      return () => { clearTimeout(delay); clearInterval(timer); };
    }, [hovered, active, visible, frames.length]);
    const missing = !frames.length || failed[page];
    return h('article', { className: 'sbi-card', 'data-kind': item.kind, 'data-template': `${item.kind}:${item.id}`,
      onMouseEnter: () => setHovered(true), onMouseLeave: () => { setHovered(false); setPage(0); } },
    h('button', { type: 'button', className: 'sbi-card-open', 'aria-label': format(t, 'openDetail', { name: item.title }), onClick: () => onPreview(item) },
      h('span', { className: 'sbi-cover', style: { aspectRatio: FRAME_RATIO[item.kind] } },
        ...frames.map((src, index) => h('img', { key: `${index}`, src, alt: '', draggable: false, className: index === page && !failed[index] ? 'sbi-current' : '',
          onError: () => setFailed(current => ({ ...current, [index]: true })) })),
        missing ? h('span', { className: 'sbi-cover-empty' }, h('img', { src: FILE_ICONS[item.kind], width: 32, height: 32, alt: '', draggable: false }), h('span', null, format(t, 'noCover'))) : null),
      h('span', { className: 'sbi-card-title', title: `${KIND_LABEL[item.kind]} · ${item.title}` },
        h(FileTypeContent, { kind: item.kind }),
        h('span', { className: 'sbi-card-divider', 'aria-hidden': true }),
        h('span', { className: 'sbi-card-name' }, item.title))),
    h('div', { className: 'sbi-card-actions' },
      h('button', { type: 'button', className: 'sbi-quick-use', 'aria-label': format(t, 'useNamed', { name: item.title }), onClick: () => onUse(item) }, format(t, 'use'))));
  }

  /** Fixed 280px cards, each into the shortest column; the container decides how many columns fit (§3). */
  function MasonryGrid({ items, onPreview, onUse, active, t }) {
    const container = useRef(null);
    const [size, setSize] = useState({ width: CARD_WIDTH, gap: 20 });
    useLayoutEffect(() => {
      const node = container.current;
      const W = win();
      if (!node || !W) return undefined;
      const measure = () => {
        const style = W.getComputedStyle(node);
        const width = node.clientWidth - (parseFloat(style.paddingLeft) || 0) - (parseFloat(style.paddingRight) || 0);
        const gap = parseFloat(style.getPropertyValue('--sbi-gap')) || 20;
        setSize(previous => (Math.abs(previous.width - width) < 0.5 && previous.gap === gap ? previous : { width, gap }));
      };
      measure();
      if (!W.ResizeObserver) return undefined;
      const observer = new W.ResizeObserver(measure);
      observer.observe(node);
      return () => observer.disconnect();
    }, []);
    const count = columnCount(size.width, size.gap, items.length);
    const columns = masonry(items, count, size.gap);
    return h('div', { ref: container, className: 'sbi-grid', style: { '--sbi-columns': count, '--sbi-card-width': `${CARD_WIDTH}px` }, 'data-columns': count },
      ...columns.map((list, index) => h('div', { key: `${index}`, className: 'sbi-column' },
        ...list.map(item => h(CoverCard, { key: `${item.kind}:${item.id}`, item, onPreview, onUse, active, t })))));
  }

  /** "My templates" with none yet: choose or drop one file; checked by name, size and count; ends at "not connected". */
  function CreateEntry({ state, kind, t }) {
    const input = useRef(null);
    const entry = state.uploads?.[kind] ?? { state: 'idle' };
    const dragging = state.dragKind === kind;
    const errorText = entry.state === 'error'
      ? format(t, { multiple: 'uploadMultiple', empty: 'uploadEmpty', type: 'uploadType', none: 'uploadNone' }[entry.error] ?? 'uploadNone', { list: UPLOAD_EXTENSIONS[kind].join(' ') }) : '';
    const stop = event => { event.preventDefault?.(); event.stopPropagation?.(); };
    const body = entry.state === 'unavailable'
      ? [h('span', { key: 'f', className: 'sbi-create-file' }, entry.fileName), h('span', { key: 'u', className: 'sbi-create-note', role: 'status' }, format(t, 'uploadUnavailable')), h('span', { key: 'r', className: 'sbi-link' }, format(t, 'reselect'))]
      : [h('span', { key: 'p', className: 'sbi-create-plus', 'aria-hidden': true }, icon('plus', 20)),
        h('strong', { key: 't', className: 'sbi-create-title' }, format(t, 'createTitle')),
        h('span', { key: 'h', className: 'sbi-create-hint' }, format(t, dragging ? 'dropHere' : 'createHint')),
        errorText ? h('span', { key: 'e', className: 'sbi-create-error', role: 'alert' }, errorText) : h('span', { key: 'o', className: 'sbi-create-offline' }, format(t, 'createOffline'))];
    return h('div', { className: 'sbi-create-wrap' },
      h('button', { type: 'button', className: `sbi-create${dragging ? ' sbi-drag' : ''}`, 'data-shimo-inspiration': 'create', 'data-state': entry.state,
        onClick: () => { if (entry.state !== 'idle') library.resetUpload(kind); try { input.current?.click?.(); } catch {} },
        onDragEnter: event => { stop(event); library.setDrag(kind); },
        onDragOver: event => { stop(event); if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy'; library.setDrag(kind); },
        onDragLeave: event => { if (!event.currentTarget?.contains?.(event.relatedTarget)) library.setDrag(null); },
        onDrop: event => { stop(event); library.receiveFiles(kind, event.dataTransfer?.files); } }, ...body),
      h('input', { ref: input, type: 'file', accept: UPLOAD_EXTENSIONS[kind].join(','), hidden: true, tabIndex: -1, 'aria-hidden': true,
        onChange: event => { library.receiveFiles(kind, event.target?.files); try { event.target.value = ''; } catch {} } }));
  }

  /** The page's bottom line: the workspace-default box (a PPT the current Session carries) and what is added or remembered. */
  function Footer({ store, picked, mode, t, lang }) {
    const list = picked?.list;
    const pick = picked?.pick;
    const kind = pick?.kind ?? mode;
    const remembered = list?.remembered ?? null;
    const card = remembered ? list?.styles?.find(item => item.id === remembered) : undefined;
    const defaultName = remembered ? (card ? styleTitle(card, lang) : format(t, 'missing', { id: remembered })) : undefined;
    const folder = list?.rememberedFrom ? list.rememberedFrom.split(/[\\/]/).filter(Boolean).at(-1) ?? list.rememberedFrom : undefined;
    const pptPicked = pick?.kind === 'ppt';
    const inheritedPick = pptPicked && folder && remembered === pick.id;
    let why;
    if (kind !== 'ppt') why = format(t, 'rememberOther');
    else if (list && !list.canRemember) why = format(t, 'rememberGuest');
    else if (inheritedPick) why = format(t, 'rememberInherited', { folder });
    const enabled = Boolean(store) && pptPicked && Boolean(list?.canRemember) && !inheritedPick && !picked?.busy;
    const checked = pptPicked && remembered === pick.id;
    let status;
    if (picked?.notice) status = h('span', { className: `sbi-status sbi-${picked.notice.tone}`, role: picked.notice.tone === 'error' ? 'alert' : 'status' }, format(t, picked.notice.key, picked.notice.params));
    else {
      const parts = [];
      if (pick) parts.push(format(t, 'statusPicked', { kind: KIND_LABEL[pick.kind], name: pick.title }));
      if (kind === 'ppt' && remembered && !(pptPicked && pick.id === remembered)) parts.push(format(t, folder ? 'statusDefaultFrom' : 'statusDefault', { name: defaultName, folder }));
      if (kind !== 'ppt') parts.push(format(t, 'rememberOther'));
      else if (!parts.length) parts.push(format(t, 'status'));
      status = h('span', { className: 'sbi-status', role: 'status' }, parts.join('　'));
    }
    return h('div', { className: 'sbi-foot' },
      h('label', { className: `sbi-check${enabled ? '' : ' sbi-off'}`, title: why },
        h('input', { type: 'checkbox', checked, disabled: !enabled, onChange: event => { void store?.setDefault(event.target.checked); } }), format(t, 'remember')),
      status);
  }

  /** Native modal dialog lifecycle: open on mount, close on unmount. */
  const useModal = ref => useLayoutEffect(() => {
    const node = ref.current;
    try { node?.showModal?.(); } catch { /* already open */ }
    return () => { try { node?.close?.(); } catch {} };
  }, []);

  /** The lightbox: the whole page on a dark backdrop; ←/→/Home/End; Esc closes only this layer; keys never reach the page. */
  function PageLightbox({ template, index, onChange, onClose, t }) {
    const dialog = useRef(null);
    const total = template.pages.length;
    const move = next => onChange(Math.max(0, Math.min(total - 1, next)));
    useModal(dialog);
    return h('dialog', { ref: dialog, className: 'sbi-lightbox', 'aria-label': format(t, 'lightboxTitle', { name: template.title }), 'data-shimo-inspiration': 'lightbox',
      onCancel: event => { event.preventDefault?.(); onClose(); },
      onClick: event => { if (event.target === event.currentTarget) onClose(); },
      onKeyDown: event => {
        event.stopPropagation?.();
        const next = pageKey(event.key, index, total);
        if (next !== undefined) { event.preventDefault?.(); move(next); }
      } },
    h('button', { type: 'button', className: 'sbi-lightbox-close', autoFocus: true, 'aria-label': format(t, 'lightboxClose'), onClick: onClose }, icon('close', 20)),
    h('button', { type: 'button', className: 'sbi-lightbox-arrow sbi-prev', 'aria-label': format(t, 'prev'), disabled: index === 0, hidden: total < 2, onClick: () => move(index - 1) }, icon('left', 26)),
    h('div', { className: 'sbi-lightbox-image' }, h(PageImage, { src: template.pages[index], kind: template.kind, alt: format(t, 'pageAlt', { name: template.title, n: index + 1 }), t })),
    h('button', { type: 'button', className: 'sbi-lightbox-arrow sbi-next', 'aria-label': format(t, 'next'), disabled: index >= total - 1, hidden: total < 2, onClick: () => move(index + 1) }, icon('right', 26)),
    h('span', { className: 'sbi-lightbox-counter', 'aria-live': 'polite' }, `${index + 1} / ${Math.max(1, total)}`));
  }

  /**
   * The detail (§5): category, title and a summary clamped to the preview's height (「更多」 only when it overflows),
   * the carousel, the example page, every page as numbered thumbnails; body scrolls, the footer keeps 「添加并使用」.
   */
  function TemplateDetail({ template, onClose, onUse, t }) {
    const dialog = useRef(null);
    const copy = useRef(null);
    const previewBox = useRef(null);
    const summaryArea = useRef(null);
    const summaryMeasure = useRef(null);
    const [page, setPage] = useState(0);
    const [viewer, setViewer] = useState(null);
    const [expanded, setExpanded] = useState(false);
    const [layoutState, setLayoutState] = useState({ overflow: false, lines: 1 });
    const total = template.pages.length;
    const category = template.categories?.[0] || format(t, `category_${template.kind}`);
    const pageLabel = format(t, `page_${template.kind}`);
    const summary = template.summary || format(t, 'summaryFallback');
    useModal(dialog);
    useLayoutEffect(() => {
      const W = win();
      if (!W || !previewBox.current) return undefined;
      let alive = true;
      const measure = () => {
        if (!alive || !previewBox.current || !summaryArea.current || !copy.current || !summaryMeasure.current) return;
        const available = previewBox.current.getBoundingClientRect().height - (summaryArea.current.getBoundingClientRect().top - copy.current.getBoundingClientRect().top);
        const lineHeight = parseFloat(W.getComputedStyle(summaryMeasure.current).lineHeight);
        const overflow = summaryMeasure.current.getBoundingClientRect().height > available + 0.5;
        const lines = summaryLines(available, lineHeight, overflow);
        setLayoutState(previous => (previous.overflow === overflow && previous.lines === lines ? previous : { overflow, lines }));
      };
      measure();
      const observer = W.ResizeObserver ? new W.ResizeObserver(measure) : undefined;
      for (const node of [previewBox.current, copy.current, summaryMeasure.current]) if (node) observer?.observe(node);
      doc?.fonts?.ready?.then?.(measure);
      return () => { alive = false; observer?.disconnect(); };
    }, [summary, template.title]);
    const pageAlt = n => format(t, 'pageAlt', { name: template.title, n });
    const preview = h('div', { ref: previewBox, className: 'sbi-carousel', 'data-kind': template.kind, style: { '--sbi-ratio': FRAME_RATIO[template.kind] } },
      h('button', { type: 'button', className: 'sbi-carousel-image', 'aria-label': format(t, 'zoom'), disabled: !total, onClick: () => setViewer(page) },
        h(PageImage, { src: template.pages[page], kind: template.kind, alt: pageAlt(page + 1), t })),
      total > 1 ? h('button', { type: 'button', className: 'sbi-carousel-arrow sbi-prev', 'aria-label': format(t, 'prevPage', { page: pageLabel }), disabled: page === 0, onClick: () => setPage(value => Math.max(0, value - 1)) }, icon('left', 18)) : null,
      total > 1 ? h('button', { type: 'button', className: 'sbi-carousel-arrow sbi-next', 'aria-label': format(t, 'nextPage', { page: pageLabel }), disabled: page === total - 1, onClick: () => setPage(value => Math.min(total - 1, value + 1)) }, icon('right', 18)) : null,
      total > 0 ? h('span', { className: 'sbi-carousel-counter', 'aria-live': 'polite' }, `${page + 1} / ${total}`) : null);
    const feature = featurePage(template.pages);
    return h(Fragment, null,
      h('dialog', { ref: dialog, className: 'sbi-detail', 'aria-labelledby': 'sbi-detail-title', 'data-shimo-inspiration': 'detail',
        onCancel: event => { event.preventDefault?.(); onClose(); },
        onKeyDown: event => { if (event.key === 'Escape') event.stopPropagation?.(); },
        onClick: event => { if (event.target === event.currentTarget) onClose(); } },
      h('div', { className: 'sbi-detail-layout' },
        h('button', { type: 'button', className: 'sbi-detail-close', 'aria-label': format(t, 'detailClose'), autoFocus: true, onClick: onClose }, icon('close', 16)),
        h('div', { className: 'sbi-detail-scroll' },
          h('header', { className: 'sbi-detail-intro' },
            h('div', { ref: copy, className: 'sbi-detail-copy' },
              h('span', { className: 'sbi-detail-category' }, category),
              h('h2', { id: 'sbi-detail-title' }, template.title),
              h('div', { ref: summaryArea, className: 'sbi-summary-area' },
                h('blockquote', { id: 'sbi-summary', className: `sbi-summary${!expanded && layoutState.overflow ? ' sbi-clamped' : ''}`, style: { '--sbi-lines': layoutState.lines } }, summary),
                h('div', { className: 'sbi-summary-measure', 'aria-hidden': true }, h('blockquote', { ref: summaryMeasure, className: 'sbi-summary' }, summary)),
                layoutState.overflow ? h('button', { type: 'button', className: 'sbi-summary-more', 'aria-expanded': expanded, 'aria-controls': 'sbi-summary', onClick: () => setExpanded(value => !value) }, format(t, expanded ? 'less' : 'more')) : null)),
            preview),
          feature ? h('figure', { className: 'sbi-feature', 'data-kind': template.kind }, h(PageImage, { src: feature, kind: template.kind, alt: format(t, 'featureAlt', { name: template.title }), t })) : null,
          total > 0 ? h('section', { className: 'sbi-pages', 'aria-labelledby': 'sbi-pages-title' },
            h('div', { className: 'sbi-pages-heading' }, h('h3', { id: 'sbi-pages-title' }, format(t, 'allPages', { page: pageLabel })), h('span', null, format(t, `pageCount_${template.kind}`, { n: total }))),
            h('div', { className: 'sbi-thumbs' }, ...template.pages.map((src, index) => h('button', { key: `${index}`, type: 'button', className: 'sbi-thumb', 'aria-label': format(t, 'zoomPage', { n: index + 1 }), onClick: () => setViewer(index) },
              h('img', { src, alt: format(t, 'thumbAlt', { n: index + 1 }), loading: 'lazy', draggable: false }),
              h('span', { className: 'sbi-thumb-badge', 'aria-hidden': true }, String(index + 1).padStart(2, '0')))))) : null),
        h('footer', { className: 'sbi-detail-foot' },
          h('button', { type: 'button', className: 'sbi-detail-use', onClick: () => onUse(template) }, format(t, 'use'))))),
      viewer !== null ? h(PageLightbox, { template, index: viewer, onChange: setViewer, onClose: () => setViewer(null), t }) : null);
  }

  /** The `main` entry: the whole page. */
  function InspirationPage({ t }) {
    const state = useLibrary();
    const lang = languageOf(t);
    const sessionId = useSessionId();
    const store = sessionId ? stores.get(sessionId) : undefined;
    const picked = useSyncExternalStore(listener => store?.subscribe(listener) ?? (() => {}), () => store?.snapshot());
    // Opening (or the Session changing) refreshes the PPT list for that Session's workspace.
    useEffect(() => { void store?.load(); }, [store]);
    // Esc on the page (no detail open) goes back to the conversation, as the back button and the entry do.
    useEffect(() => {
      if (!doc?.addEventListener) return undefined;
      const onKey = event => {
        if (event.key !== 'Escape' || event.isComposing || event.defaultPrevented || library.state.detail) return;
        event.preventDefault?.();
        close();
      };
      doc.addEventListener('keydown', onKey);
      return () => doc.removeEventListener('keydown', onKey);
    }, []);
    const templates = libraryTemplates(picked?.list?.styles, lang, picked?.list?.templates);
    // A kind this copy has no skill for (the public build: Excel) has no tab and no create entry; its mode falls back to PPT.
    const kinds = kindsOf(picked?.list?.templates);
    const mode = kinds.includes(state.mode) ? state.mode : 'ppt';
    const scenes = categoriesOf(templates, 10);
    const scope = state.scope === SCOPE_MINE || state.scope === SCOPE_FEATURED || scenes.includes(state.scope) ? state.scope : SCOPE_FEATURED;
    const shown = scopeTemplates(templates, scope);
    const use = template => {
      library.use(template, { target: sessionId ?? null });
      if (!sessionId) workspace()?.startSession?.();
      close();
    };
    const scopes = [SCOPE_MINE, SCOPE_FEATURED, ...scenes];
    const chipText = value => (value === SCOPE_MINE ? format(t, 'scopeMine') : value === SCOPE_FEATURED ? format(t, 'scopeFeatured') : value);
    const filters = h('div', { className: 'sbi-filters', role: 'toolbar', 'aria-label': format(t, 'filters') },
      ...scopes.map(value => h('button', { key: value, type: 'button', className: 'sbi-chip', 'aria-pressed': value === scope, 'data-scope': value, onClick: () => library.setScope(value) }, chipText(value))));
    let pptState = null;
    if (!store) pptState = h('div', { className: 'sbi-state', role: 'status' }, format(t, 'pptNeedsSession'));
    else if (!picked?.list && picked?.phase === 'error') pptState = h('div', { className: 'sbi-state', role: 'alert' }, format(t, 'failed', { message: picked.error ?? '' }), ' ', h('button', { type: 'button', className: 'sbi-link', onClick: () => store.load() }, format(t, 'retry')));
    else if (!picked?.list) pptState = h('div', { className: 'sbi-state', role: 'status' }, format(t, 'loading'));
    let body;
    if (scope === SCOPE_MINE && !shown.length) body = h('div', { className: 'sbi-create-area' }, h(CreateEntry, { state, kind: mode, t }));
    else if (!shown.length) body = h('div', { className: 'sbi-empty' }, format(t, 'noResults'));
    else body = h(MasonryGrid, { items: shown, onPreview: item => library.openDetail(item), onUse: use, active: !state.detail, t });
    const title = format(t, 'heroTitle');
    return h('div', { className: 'sbi-page', role: 'region', 'aria-label': format(t, 'title'), 'data-shimo-inspiration': 'page',
      onDragOver: event => event.preventDefault?.(), onDrop: event => event.preventDefault?.() },
    h('button', { type: 'button', className: 'sbi-back', 'aria-label': format(t, 'back'), title: format(t, 'back'), onClick: close }, icon('back', 16), h('span', null, format(t, 'back'))),
    h('section', { className: 'sbi-compose' },
      // The full-width question mark's empty right half is dropped so the words centre on the switch.
      h('h1', { className: 'sbi-title' }, title.replace(/[？?]$/, ''), /[？?]$/.test(title) ? h('span', { className: 'sbi-title-mark' }, title.slice(-1)) : null),
      h(TypeSwitch, { mode, kinds, onChange: next => library.setMode(next), t })),
    h('section', { className: 'sbi-gallery' },
      h('div', { className: 'sbi-toolbar' }, filters),
      h('div', { className: 'sbi-scroll' }, pptState && scope !== SCOPE_MINE ? pptState : null, body),
      h(Footer, { store, picked, mode, t, lang })),
    state.detail ? h(TemplateDetail, { template: state.detail, onClose: () => library.closeDetail(), onUse: use, t }) : null);
  }

  /**
   * `sidebar.panellist`: the bulb. DSH's row button selects the panel; pressing it while the page is already open goes
   * back to the conversation (the row has no toggle of its own).
   */
  function PanelGlyph({ size = 18, active }) {
    const ref = useRef(null);
    const activeRef = useRef(active);
    activeRef.current = active;
    useEffect(() => {
      const button = ref.current?.closest?.('button');
      if (!button?.addEventListener) return undefined;
      // The button's own listener runs before DSH's (delegated) onClick, so it still sees the state before this press;
      // Enter and Space fire `click` too.
      const onClick = () => { if (activeRef.current === true) setTimeout(close, 0); };
      button.addEventListener('click', onClick);
      return () => button.removeEventListener('click', onClick);
    }, []);
    return h('span', { ref, className: 'sbi-glyph', 'data-shimo-inspiration': 'entry' }, bulb(h, size, 'sbi-bulb'));
  }

  return { InspirationPage, PanelGlyph, TemplateDetail, PageLightbox, MasonryGrid, CoverCard, TypeSwitch, FileTypeContent, CreateEntry, Footer };
}

/** Registers the sidebar entry and the page; `library` is shared with the input box's slot (style-picker.mjs). */
export function registerInspirationLibrary(ctx, React, { stores, library, document } = {}) {
  const layout = () => ctx.get?.('layout');
  const workspace = () => ctx.get?.('uiWorkspace');
  library.currentSession = () => workspace()?.selection?.getSnapshot?.()?.sessionId;
  const { InspirationPage, PanelGlyph } = createInspirationLibrary({ React, stores, library, layout, workspace, document });
  const label = () => ctx.locale.bind(STYLE_PICKER_NS)('panel');
  ctx.effect(() => ctx.slots.inject('main', () => ctx.slots.register({ name: 'main', key: INSPIRATION_PANEL_ID, locale: STYLE_PICKER_NS }, InspirationPage)));
  ctx.effect(() => ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({ name: 'sidebar.panellist', id: INSPIRATION_PANEL_ID, order: INSPIRATION_PANEL_ORDER, locale: STYLE_PICKER_NS, label }, PanelGlyph)));
  return { InspirationPage, PanelGlyph };
}
