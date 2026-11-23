import { unwrapRemote } from './address.mjs';
import { FILTER_ALL, gridKey, styleTitle } from './style-cards.mjs';
import { BUILTIN, BULB_PATHS, KIND_LABEL, composeDraft, promptOf, stripPrompt } from './inspiration.mjs';

export { gridKey, styleTitle };

/**
 * The inspiration library's per-Session half (design handoff OfficeSDK-DSH-Inspiration-v2.0). The library itself is a
 * full page behind the sidebar entry (inspiration-library.mjs); a template's 「添加并使用」 comes back here, to the input
 * box of the Session the user returns to: the Host records the template for that Session first (`selectTemplate`, the
 * data-layer template + skill association its runtime context names), and only then is the kind's fixed prompt written at
 * the start of DSH's own draft and the input box focused. DSH's own send button sends it; nothing here ever sends.
 * The input box's toolbar (`conversation.input.left`) shows the borderless 「bulb + template name」 prefix with a × on
 * hover; Backspace at the very start of the draft removes it too (DSH has no slot inside the text area). PPT templates
 * are the shimo-pptx style packs, whose workspace default the Host keeps (`rememberStyle`). Styling: `sbs-` classes.
 */
export const STYLE_PICKER_NS = 'shimoStyles';
export const STYLE_PICKER_ID = 'officesdk-ppt-style';
const ZH = {
  lang: 'zh', panel: '灵感库', title: '灵感库', heroTitle: '今天，你想从哪开始？', back: '返回', close: '关闭',
  modes: '选择创作类型', filters: '按使用场景筛选', scopeMine: '我的模板', scopeFeatured: '精选', noResults: '这个场景下还没有模板',
  use: '添加并使用', useNamed: '添加并使用 {name}', openDetail: '查看 {name} 模板详情', noCover: '封面暂不可用', noPreview: '暂无页面预览',
  more: '更多', less: '收起', detailClose: '关闭模板详情', zoom: '点击放大', prevPage: '上一张{page}', nextPage: '下一张{page}',
  allPages: '全部{page}', pageCount_word: '{n} 页', pageCount_excel: '{n} 张工作表', pageCount_ppt: '{n} 张幻灯片',
  page_word: '文档页面', page_excel: '工作表', page_ppt: '幻灯片', category_word: '文档模板', category_excel: '表格模板', category_ppt: '演示模板',
  pageAlt: '{name} · 第 {n} 页', featureAlt: '{name} · 版式示例', lightboxTitle: '{name} 全屏预览', lightboxClose: '关闭放大预览', prev: '上一页', next: '下一页',
  zoomPage: '放大第 {n} 页', thumbAlt: '第 {n} 页缩略图', summaryFallback: '保留你喜欢的页面样式，从已有模板开始创作。',
  loading: '正在读取模板…', failed: '读取模板失败：{message}', retry: '重试', pptNeedsSession: '打开一个工作区后才能读取 PPT 模板',
  defaultBadge: '默认', remember: '在这个工作区默认使用', rememberOther: 'Word、Excel 的工作区默认模板暂未接入', rememberInherited: '这个默认来自上层文件夹 {folder}，请在那个工作区修改',
  rememberGuest: '只有工作区所有者可以设置默认模板',
  status: '添加并使用后，模板与提示词会带入对话框，由你按发送开始', statusPicked: '已带入 · {kind} · {name}', statusDefault: '本工作区默认：{name}', statusDefaultFrom: '本工作区默认：{name}（沿用上层文件夹 {folder}）',
  remembered: '已设为本工作区默认：{name}', forgotten: '已取消本工作区默认样式', rememberFailed: '没记住：{message}', useFailed: '模板未能带入：{message}', missing: '{id}（已不存在）',
  tagRemove: '移除模板 {name}', tagTitle: '已带入模板：{kind} · {name}',
  createTitle: '创建我的模板', createHint: '添加喜欢的样式、图片', createOffline: '上传与制作服务暂未接入', dropHere: '松开即可上传',
  uploadMultiple: '一次只能选一个文件', uploadEmpty: '文件是空的，请换一个', uploadType: '这里只接受 {list} 文件', uploadNone: '没有收到文件，请重新选择',
  uploadUnavailable: '模板上传与制作服务尚未接入，暂时不能创建我的模板', reselect: '重新选择',
  all: '全部', filtersSuggest: '按场景筛选', scenarios: '适用：{list}',
  suggestTitle: '按你的内容推荐的版式', suggestTopic: '「{topic}」', suggestDefault: '都不要，用默认（{name}）', suggestAll: '看全部样式', suggestBack: '返回推荐', suggestClose: '收起推荐',
};
const EN = {
  lang: 'en', panel: 'Inspiration', title: 'Inspiration', heroTitle: 'Where would you like to start today?', back: 'Back', close: 'Close',
  modes: 'Choose what to create', filters: 'Filter by occasion', scopeMine: 'My templates', scopeFeatured: 'Featured', noResults: 'No templates for this occasion yet',
  use: 'Add and use', useNamed: 'Add and use {name}', openDetail: 'Details of the {name} template', noCover: 'Cover unavailable', noPreview: 'No page preview',
  more: 'More', less: 'Less', detailClose: 'Close template details', zoom: 'Click to enlarge', prevPage: 'Previous {page}', nextPage: 'Next {page}',
  allPages: 'All {page}s', pageCount_word: '{n} pages', pageCount_excel: '{n} sheets', pageCount_ppt: '{n} slides',
  page_word: 'page', page_excel: 'sheet', page_ppt: 'slide', category_word: 'Document template', category_excel: 'Spreadsheet template', category_ppt: 'Presentation template',
  pageAlt: '{name} · page {n}', featureAlt: '{name} · layout example', lightboxTitle: '{name} full-screen preview', lightboxClose: 'Close the enlarged preview', prev: 'Previous page', next: 'Next page',
  zoomPage: 'Enlarge page {n}', thumbAlt: 'Page {n} thumbnail', summaryFallback: 'Keep the page style you like and start from an existing template.',
  loading: 'Loading templates…', failed: 'Could not load templates: {message}', retry: 'Retry', pptNeedsSession: 'Open a workspace to load the PPT templates',
  defaultBadge: 'Default', remember: 'Use as this workspace\'s default', rememberOther: 'Workspace defaults for Word and Excel are not available yet', rememberInherited: 'This default comes from the enclosing folder {folder}; change it in that workspace',
  rememberGuest: 'Only the workspace owner can set the default template',
  status: 'Add and use puts the template and a prompt in the message box; you send it', statusPicked: 'Added · {kind} · {name}', statusDefault: 'Workspace default: {name}', statusDefaultFrom: 'Workspace default: {name} (from the enclosing folder {folder})',
  remembered: 'Now this workspace\'s default: {name}', forgotten: 'This workspace has no default style now', rememberFailed: 'Not remembered: {message}', useFailed: 'The template was not added: {message}', missing: '{id} (no longer there)',
  tagRemove: 'Remove template {name}', tagTitle: 'Template added: {kind} · {name}',
  createTitle: 'Create my template', createHint: 'Add a style or picture you like', createOffline: 'Upload and template making are not connected yet', dropHere: 'Drop to upload',
  uploadMultiple: 'Choose one file at a time', uploadEmpty: 'The file is empty; choose another one', uploadType: 'Only {list} files here', uploadNone: 'No file arrived; choose again',
  uploadUnavailable: 'Template upload and making are not connected yet, so my templates cannot be created', reselect: 'Choose again',
  all: 'All', filtersSuggest: 'Filter by occasion', scenarios: 'For: {list}',
  suggestTitle: 'Layouts suggested for your content', suggestTopic: '"{topic}"', suggestDefault: 'None of these, use the default ({name})', suggestAll: 'See all styles', suggestBack: 'Back to suggestions', suggestClose: 'Close suggestions',
};
export const STYLE_PICKER_MESSAGES = Object.freeze({ zh: ZH, en: EN });

export const format = (t, key, params = {}) => Object.entries(params).reduce((text, [name, value]) => text.split(`{${name}}`).join(String(value)), String((typeof t === 'function' ? t(key) : undefined) ?? key));
export const languageOf = t => (typeof t === 'function' && t('lang') === 'en' ? 'en' : 'zh');

/** The answer a recommendation pick sends: the pack named the way SKILL.md reads it; `fallback` for "none of these". */
export function answerText(card, lang, { fallback = false } = {}) {
  const title = styleTitle(card, lang);
  if (lang === 'en') return fallback ? `None of these; use the default style ${card.id} (${title})` : `Use the ${card.id} (${title}) style`;
  return fallback ? `都不要，用默认样式 ${card.id}（${title}）` : `用 ${card.id}（${title}）样式`;
}

/** `phrase` after `draft`, with a space only between two Latin words. */
export const appendText = (draft, phrase) => (/[A-Za-z0-9.,;:!?)]$/.test(draft) && /^[A-Za-z0-9(]/.test(phrase) ? `${draft} ${phrase}` : `${draft}${phrase}`);

/** `phrase` into a draft that has text: at `span`, else at the current caret, else appended. */
export function insertPhrase(actions, draft, phrase, span) {
  for (const at of [span, actions.captureInsertion?.()]) {
    if (at && actions.insertText?.(phrase, at)) return 'inserted';
  }
  actions.setDraft(appendText(draft, phrase));
  return 'appended';
}

/** DSH's composer editor (Lexical marks its root), tried first at each level; a plain field only as a fallback. */
const COMPOSER = ['[data-lexical-editor="true"]', '[contenteditable="true"], textarea'];

/**
 * Give the keyboard back to the input box that holds `from`: the nearest composer editor going up from it, so a second
 * pane of the Session keeps its own. inputActions has no focus verb; this does what DSH's own input shell does — focus
 * the editor's root, then Lexical's `focus()`, which keeps the caret where the draft left it. Returns whether an editor
 * was found.
 */
export function focusComposer(from) {
  for (let node = from?.parentElement; node; node = node.parentElement) {
    const root = COMPOSER.map(selector => node.querySelector?.(selector)).find(Boolean);
    if (!root) continue;
    try {
      root.focus?.({ preventScroll: true });
      root.__lexicalEditor?.focus?.();
    } catch { /* focus is a courtesy: the draft is already written */ }
    return true;
  }
  return false;
}

const messageOf = error => error?.message ?? String(error);

/**
 * One Session's inspiration state, outside React and shared by every mount (main view, split panes): the PPT style list
 * (fetched once per Session, refreshed when the library opens), the template the Session carries and the prompt it
 * wrote, a failed use (shown with a retry), the last notice, and the model's style recommendation.
 */
export class StylePickerStore {
  constructor({ rpc, sessionId }) {
    Object.assign(this, { rpc, sessionId });
    // `filter`: the scenario chip of the recommendation's "all styles" view.
    this.state = {
      phase: 'idle', list: undefined, error: undefined, notice: undefined, busy: false, filter: FILTER_ALL,
      pick: null, inserted: undefined, binding: undefined, failed: undefined, suggestion: undefined, suggestView: 'cards',
    };
    this.listeners = new Set();
    // Set when a turn took the pick away from the screen: the Host is told once that turn ends (`turnEnded`).
    this.clearAfterTurn = false;
    // Whether a turn is running (StyleDock reports it). A pick or removal made meanwhile reaches the Host only when
    // that turn ends (`pendingHost`: undefined = nothing waiting, null = clear, { kind, id } = pick), so the
    // running turn's context never changes under it.
    this.running = false;
    this.pendingHost = undefined;
  }
  /** Tell the Host the pick (or null), now when idle, else once the running turn ends. */
  async sendPick(pick) {
    if (this.running) { this.pendingHost = pick; return undefined; }
    return this.call('selectTemplate', pick ? { kind: pick.kind, id: pick.id } : { kind: 'ppt', id: null });
  }
  subscribe(listener) { this.listeners.add(listener); return () => this.listeners.delete(listener); }
  snapshot() { return this.state; }
  set(patch) { this.state = { ...this.state, ...patch }; for (const listener of this.listeners) listener(); }
  call(method, args = {}) {
    return this.rpc.call('/api', `shimoPreview/${method}`, { args: { workspaceFileScopeId: this.sessionId, ...args } }).then(unwrapRemote);
  }
  /** Fetch the PPT list; a refresh keeps the shown list until the new one arrives. */
  async load() {
    if (this.loading) return this.loading;
    if (!this.state.list) this.set({ phase: 'loading', error: undefined });
    this.loading = this.call('listStyles').then(
      list => this.set({ phase: 'ready', list, error: undefined }),
      error => this.set(this.state.list ? { notice: { tone: 'error', key: 'failed', params: { message: messageOf(error) } } } : { phase: 'error', error: messageOf(error) }),
    ).finally(() => { this.loading = undefined; });
    return this.loading;
  }
  setFilter(filter) { this.set({ filter: filter || FILTER_ALL }); }
  /** The workspace default as the list reports it (only the fields the footer needs). */
  choice(patch) { this.set({ list: { ...this.state.list, ...patch } }); }
  /**
   * Use `template` ({ id, kind, title }) in this Session: the Host records it first; only once it did is the kind's
   * fixed prompt written at the start of the draft (replacing the previous use's prompt while still there word for
   * word, keeping what the user wrote) and the prefix shown. A Host that refuses leaves the draft and the prefix
   * untouched and shows the failure with a retry, so a failed binding never looks like a success. Never sent.
   */
  async use(template, { actions, draft, lang = 'zh' }) {
    const prompt = promptOf(template.kind, lang);
    this.set({ binding: template.id, failed: undefined, notice: undefined });
    try { await this.sendPick({ kind: template.kind, id: template.id }); }
    catch (error) {
      this.set({ binding: undefined, failed: { template, message: messageOf(error) } });
      return { written: false, error };
    }
    actions?.setDraft?.(composeDraft(typeof draft === 'string' ? draft : '', this.state.inserted, prompt));
    this.clearAfterTurn = false;
    this.set({ pick: { kind: template.kind, id: template.id, title: template.title }, inserted: prompt, binding: undefined });
    return { written: true };
  }
  /** Forget a failed use (its × in the toolbar). */
  dismissFailure() { if (this.state.failed) this.set({ failed: undefined }); }
  /** The prefix's × or Backspace at the start: no pick; the prompt leaves the draft only while still at its start word for word. */
  async clearPick({ actions, draft } = {}) {
    const pick = this.state.pick;
    if (!pick) return false;
    const kept = stripPrompt(draft, this.state.inserted);
    if (typeof draft === 'string' && kept !== draft) actions?.setDraft?.(kept);
    this.clearAfterTurn = false;
    this.set({ pick: null, inserted: undefined });
    try { await this.sendPick(null); } catch { /* the Host drops it by itself after the turn */ }
    return true;
  }
  /** The user sent: the prefix goes now; the Host keeps the pick until that turn ends, so every step of it sees it. */
  turnStarted() {
    this.running = true;
    if (!this.state.pick) return;
    this.clearAfterTurn = true;
    this.set({ pick: null, inserted: undefined });
  }
  /** The turn ended: the Host forgets the pick it carried, then gets whatever the user picked or removed meanwhile. */
  async turnEnded() {
    this.running = false;
    const clear = this.clearAfterTurn;
    const pending = this.pendingHost;
    this.clearAfterTurn = false;
    this.pendingHost = undefined;
    if (pending !== undefined) {
      try { await this.sendPick(pending); return; }
      catch (error) {
        // The Host must not keep the finished turn's pick while the screen shows another: clear it and say so.
        if (pending) this.set({ notice: { tone: 'error', key: 'useFailed', params: { message: messageOf(error) } } });
        try { await this.sendPick(null); } catch { /* the Host's own window ends it */ }
        return;
      }
    }
    if (clear) { try { await this.sendPick(null); } catch { /* the Host's own window ends it */ } }
  }
  /**
   * On mount: show the prefix for the pick the Host still holds for this Session (a reload, a second pane); a pick the
   * screen cannot name is cleared on the Host so the context never names a template the user does not see.
   */
  async restore(lang = 'zh', { running = false } = {}) {
    if (this.restored) return this.restored;
    // Mounted during a turn: that turn may still be using the pick; showing it again would carry it into the next one,
    // so it is cleared on the Host when the turn ends instead.
    if (running) { this.clearAfterTurn = true; return undefined; }
    this.restored = (async () => {
      let held;
      try { held = await this.call('templatePick'); } catch { return; }
      if (!held || this.state.pick) return;
      let title;
      if (held.kind === 'ppt') {
        if (!this.state.list) await this.load();
        const card = this.state.list?.styles?.find(item => item.id === held.id);
        title = card ? styleTitle(card, lang) : undefined;
      } else {
        const item = BUILTIN[held.kind]?.find(entry => entry.id === held.id);
        title = item ? (item.name[lang] || item.name.zh) : undefined;
      }
      if (this.state.pick || this.state.binding) return;   // the user used one while this was loading: theirs stands
      // `inserted` is the prompt the use would have written: it is replaced or removed only while still there word for word.
      if (title) this.set({ pick: { kind: held.kind, id: held.id, title }, inserted: promptOf(held.kind, lang) });
      else { try { await this.call('selectTemplate', { kind: held.kind, id: null }); } catch {} }
    })();
    return this.restored;
  }
  /** The library's footer box: tick = remember the PPT pack this Session carries as the workspace's default; untick = forget it. */
  async setDefault(checked) {
    const pick = this.state.pick;
    if (checked && pick?.kind !== 'ppt') return;
    this.set({ busy: true });
    try {
      this.choice(await this.call('rememberStyle', { style: checked ? pick.id : null }));
      this.set({ busy: false, notice: checked ? { tone: 'ok', key: 'remembered', params: { name: pick.title } } : { tone: 'ok', key: 'forgotten' } });
    } catch (error) {
      this.set({ busy: false, notice: { tone: 'error', key: 'rememberFailed', params: { message: messageOf(error) } } });
    }
  }
  /**
   * The recommendation strip: ask the Host for this Session's open recommendation (it then counts as shown). A failed
   * ask keeps what is shown. Returns the recommendation or undefined.
   */
  async fetchSuggestion() {
    if (this.fetching) return this.fetching;
    this.fetching = this.call('styleSuggestion').then(
      suggestion => {
        const fresh = suggestion && suggestion.id !== this.state.suggestion?.id;
        this.set({ suggestion: suggestion ?? undefined, suggestView: suggestion && !fresh ? this.state.suggestView : 'cards' });
        return suggestion ?? undefined;
      },
      () => this.state.suggestion,
    ).finally(() => { this.fetching = undefined; });
    return this.fetching;
  }
  /** Stop showing the recommendation and tell the Host it is over (only if it is still that one there). */
  async dismissSuggestion(id = this.state.suggestion?.id) {
    if (!id) return false;
    if (this.state.suggestion?.id === id) this.set({ suggestion: undefined, suggestView: 'cards' });
    try { return (await this.call('dismissStyleSuggestion', { id }))?.dismissed === true; } catch { return false; }
  }
  /** Show the recommended cards, or every style ('all', loading the list). */
  setSuggestView(view) {
    this.set({ suggestView: view === 'all' ? 'all' : 'cards' });
    return view === 'all' ? this.load() : undefined;
  }
  /**
   * Answer the recommendation with `card`. When the draft has no text and no attachment and the Session is idle, the
   * answer becomes the message and is sent (`defer` runs the send after the editor took the text); otherwise it goes
   * into the draft beside what the user has there, unsent. The recommendation ends either way.
   */
  answer(card, { actions, input, lang, fallback = false, running = false, span, defer = run => setTimeout(run, 0) }) {
    const text = answerText(card, lang, { fallback });
    const draft = typeof input?.draft === 'string' ? input.draft : '';
    const attachments = Array.isArray(input?.attachmentIds) ? input.attachmentIds.length : 0;
    let how;
    if (!draft.trim() && !attachments && !running && typeof actions.submit === 'function') {
      actions.setDraft(text);
      defer(() => actions.submit());
      how = 'sent';
    } else how = draft.trim() ? insertPhrase(actions, draft, text, span) : (actions.setDraft(text), 'set');
    void this.dismissSuggestion();
    return how;
  }
}

/** Stores per Session, shared by every mount of the input box (main view, split panes) and the library page. */
export function createStylePickerStores(rpc) {
  const stores = new Map();
  return { get: sessionId => { let store = stores.get(sessionId); if (!store) stores.set(sessionId, store = new StylePickerStore({ rpc, sessionId })); return store; }, stores };
}

/** The bulb with the thin bolt, as the handoff draws it: one stroke in the text colour. */
export const bulb = (h, size = 16, className = 'sbs-bulb') => h('svg', { width: size, height: size, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.5, strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': true, className },
  ...BULB_PATHS.map(d => h('path', { key: d, d })));

/** The caret is at the very start of `root` (nothing but whitespace-free emptiness before it) and nothing is selected. */
export function caretAtStart(root, selection) {
  if (!root || !selection || selection.rangeCount < 1 || !selection.isCollapsed) return false;
  const range = selection.getRangeAt(0);
  if (!root.contains?.(range.startContainer)) return false;
  try {
    const before = root.ownerDocument.createRange();
    before.selectNodeContents(root);
    before.setEnd(range.startContainer, range.startOffset);
    return before.toString() === '';
  } catch { return false; }
}

/** Whether a keydown is a plain Backspace outside an IME composition (§7: composing must never remove the template). */
export const plainBackspace = event => event?.key === 'Backspace' && !event.isComposing && event.keyCode !== 229 && !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey;

/** The composer editor root of the input box that holds `from` (the same search focusComposer makes). */
export function composerOf(from) {
  for (let node = from?.parentElement; node; node = node.parentElement) {
    const root = COMPOSER.map(selector => node.querySelector?.(selector)).find(Boolean);
    if (root) return root;
  }
  return undefined;
}

/**
 * The slot component: the prefix of the template this Session carries (or a failed use with its retry). It also takes
 * the library's pending use meant for this Session (`library.takePending`), binds it, writes the draft and focuses the
 * input box.
 */
export function createStylePicker({ React, stores, library }) {
  const { createElement: h, useEffect, useRef, useSyncExternalStore } = React;
  const noLibrary = { subscribe: () => () => {}, snapshot: () => null };
  return function StylePicker({ sessionId, inputActions, useInput, useSessionStatus, session, t }) {
    const store = stores.get(sessionId);
    const state = useSyncExternalStore(listener => store.subscribe(listener), () => store.snapshot());
    const shared = library ?? noLibrary;
    const pending = useSyncExternalStore(listener => shared.subscribe(listener), () => shared.snapshot()?.pending ?? null);
    const lang = languageOf(t);
    const input = useInput?.(value => value);
    const root = useRef(null);
    const latest = useRef({});
    latest.current = { draft: typeof input?.draft === 'string' ? input.draft : '', actions: inputActions, lang };
    const status = useSessionStatus?.(statuses => statuses?.get?.(sessionId)?.running);
    const running = typeof status === 'boolean' ? status : session?.running === true;
    useEffect(() => { if (!state.list && state.phase === 'idle') void store.load(); void store.restore(lang, { running }); }, [store]);
    const use = template => {
      void store.use(template, { actions: latest.current.actions, draft: latest.current.draft, lang: latest.current.lang }).then(result => {
        if (result?.written) focusComposer(root.current);
      });
    };
    // A use started on the library page lands in the input box of the Session it was meant for, once.
    useEffect(() => {
      if (!pending || !library) return;
      const template = library.takePending(sessionId);
      if (template) use(template);
    }, [pending, sessionId]);
    // Backspace at the very start of the draft removes the template and its prompt (never while composing). Listened
    // for on the input card (stable while DSH re-renders its editor); the editor is looked up when the key arrives.
    useEffect(() => {
      const host = root.current?.closest?.('[data-composer-card]') ?? root.current?.ownerDocument;
      if (!host?.addEventListener) return undefined;
      const onKey = event => {
        if (!store.state.pick || !plainBackspace(event)) return;
        const editor = composerOf(root.current);
        if (!editor?.contains?.(event.target)) return;
        if (!caretAtStart(editor, editor.ownerDocument?.getSelection?.())) return;
        event.preventDefault?.();
        event.stopPropagation?.();
        void store.clearPick({ actions: latest.current.actions, draft: latest.current.draft });
      };
      host.addEventListener('keydown', onKey, true);
      return () => host.removeEventListener('keydown', onKey, true);
    }, [store]);
    const pick = state.pick;
    const failed = state.failed;
    const remove = () => {
      void store.clearPick({ actions: inputActions, draft: latest.current.draft });
      focusComposer(root.current);
    };
    const x = (label, onClick) => h('button', { type: 'button', className: 'sbs-tag-x', 'aria-label': label, title: label, onClick }, '×');
    let tag = null;
    if (failed) {
      tag = h('span', { className: 'sbs-tag sbs-tag-failed', role: 'alert', 'data-shimo-inspiration': 'failed' },
        h('span', { className: 'sbs-tag-name' }, format(t, 'useFailed', { message: failed.message })),
        h('button', { type: 'button', className: 'sbs-link', onClick: () => use(failed.template) }, format(t, 'retry')),
        x(format(t, 'close'), () => store.dismissFailure()));
    } else if (pick) {
      tag = h('span', { className: 'sbs-tag', 'data-shimo-inspiration': 'tag', title: format(t, 'tagTitle', { kind: KIND_LABEL[pick.kind], name: pick.title }) },
        bulb(h, 14, 'sbs-bulb sbs-tag-bulb'),
        h('span', { className: 'sbs-tag-name' }, pick.title),
        x(format(t, 'tagRemove', { name: pick.title }), remove));
    }
    return h('span', { ref: root, className: 'sbs', 'data-shimo-inspiration': 'toolbar' }, tag);
  };
}

export function registerStylePicker(ctx, React, { rpc = ctx.connection?.rpc, library } = {}) {
  ctx.effect(() => ctx.locale.register(STYLE_PICKER_NS, STYLE_PICKER_MESSAGES));
  const stores = createStylePickerStores(rpc);
  const StylePicker = createStylePicker({ React, stores, library });
  ctx.effect(() => ctx.slots.inject('conversation.input.left', () => ctx.slots.register({ name: 'conversation.input.left', id: STYLE_PICKER_ID, order: 100, locale: STYLE_PICKER_NS }, StylePicker)));
  return { StylePicker, stores };
}
