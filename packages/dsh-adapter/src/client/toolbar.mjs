/** What the tab header shows; pure so the rules are tested without React. */
import { isAbsolute } from './paths.mjs';

/**
 * The update entry (design package V3.0 P区): it exists only for a confirmed newer release this caller may install,
 * while that install runs, and once it is installed and waiting to take effect. No release, a check still running or
 * failed, a fixed install and a caller who may not install show nothing and take no room.
 * @returns {'available' | 'updating' | 'ready' | 'hidden'}
 */
export function updateEntry(update) {
  if (update?.updating) return 'updating';
  if (update?.canUpdate !== true) return 'hidden';
  if (update.installed) return 'ready';
  return update.available === true ? 'available' : 'hidden';
}

/**
 * What the update bubble shows for an entry state: the version and release date come from the Host's check (never
 * sample values); `notes` is the release's highlights when the Host has them. `action` is the button after the notes.
 */
export function updatePanelView(update, entry, unsaved = false, lang = 'zh') {
  const version = entry === 'ready' ? update?.installed ?? update?.latest : update?.latest;
  // releasedAt is when `latest` was published: shown only for that version (a copy waiting to take effect may be older).
  const date = version === update?.latest && typeof update?.releasedAt === 'string' && !Number.isNaN(Date.parse(update.releasedAt)) ? update.releasedAt : undefined;
  const list = lang === 'en' ? update?.highlightsEn : update?.highlights;
  const notes = version && version === update?.latest && Array.isArray(list) ? list.filter(text => typeof text === 'string' && text.trim()) : [];
  if (entry === 'updating') return { version, date, notes, action: { label: 'updatePreparing', disabled: true, busy: true }, message: { key: 'updatePreparingNote', tone: 'progress' } };
  if (entry === 'ready') return { version, date, notes: [], action: { label: 'updateGotIt', close: true }, message: { key: update?.restartRequired === false ? 'updateReadyLoaded' : 'updateReadyRestart', tone: 'progress' } };
  // A refused or failed install says why inside the bubble too (the toolbar status sits behind it).
  const message = unsaved ? { key: 'updateUnsaved', tone: 'warning' } : update?.failed ? { text: String(update.failed), tone: 'warning' } : undefined;
  return { version, date, notes, action: { label: 'updateNow', disabled: unsaved || !notes.length, install: true }, message };
}

/**
 * Any open tab with edits not yet written (or not yet handed to DSH): an update reloads every tab, so it waits for them.
 * A save that failed leaves its tab dirty, so the update stays blocked and the edits stay.
 */
export const anyUnsaved = controllers => [...controllers.values()].some(controller => controller.state?.dirty === true || Boolean(controller.state?.unstaged));

/** The update panel follows edits in every tab and stops observing a tab as soon as it closes. */
export function watchTabEdits(controllers, tabWatchers, notify) {
  let subscriptions = [];
  const unsubscribe = () => { for (const off of subscriptions) off(); };
  const updateTabs = () => {
    unsubscribe();
    subscriptions = Array.from(controllers.values(), controller => controller.subscribe(notify));
    notify();
  };
  tabWatchers.add(updateTabs);
  updateTabs();
  return () => { tabWatchers.delete(updateTabs); unsubscribe(); };
}

/**
 * The pen (edit) and eye (read) pair of a Kit tab. Leaving edit writes the edits first, so the eye waits
 * while edits cannot be written: before authorization, or in a read-only Session. `saving` is the header's own
 * save or authorization in progress.
 */
export function modeButtons(state, saving = false) {
  if (state?.backend !== 'kit' || state.fallback || !['edit', 'view'].includes(state.mode)) return undefined;
  const busy = saving === true || state.loading === true || state.switching === true;
  const editing = state.mode === 'edit';
  const unsavable = editing && state.dirty === true && (state.authorized !== true || state.access === 'denied');
  return {
    edit: { pressed: editing, disabled: busy },
    view: { pressed: !editing, disabled: busy || unsavable, ...unsavable && { reason: 'unsaved' } },
  };
}

/**
 * The P toolbar's ten modules in their fixed order (design package v2.9, docs/01 §2). A module appears only once this
 * plugin provides it; the three-dot button follows them and cannot be hidden; the account (个人中心) comes after it.
 */
export const MODULES = Object.freeze(['undo', 'redo', 'save', 'saveAs', 'folder', 'print', 'layout', 'focus', 'slideshow', 'mode', 'update', 'account']);
const STORAGE_KEY = 'officesdk:toolbar:compact';
const NEWLY_HIDDEN = ['saveAs', 'mode'];
const DEFAULT_HIDDEN = [...NEWLY_HIDDEN, 'print', 'update'];

/**
 * Busy for the P buttons: the header's own work (authorization, export), the toolbar's Save, an AI edit writing, a
 * mode switch or a reload. Kit autosave is not busy: the toolbar would flicker at every pause.
 */
export const toolbarBusy = (busy, state) => busy === true || state?.switching === true || state?.loading === true
  || state?.aiBusy === true || state?.saveState === 'saving';

// Formats whose Kit editor prints every page through the browser's print dialog. Writer keeps a print-ready copy
// of its pages; Sheet and PPT build theirs only in their own print flows, which office-kit leaves disabled (a blank
// page otherwise). Checked in a real DSH, 2026-10-30.
const PRINT_FORMATS = new Set(['docx']);

/**
 * Undo/redo, Save and Print for this tab: whether each is offered, and when offered whether it is unavailable now
 * and why. Undo/redo exist for PPT only, once its editor announced its history; read mode keeps them off.
 */
export function editButtons(state, busy = false) {
  const kit = state?.backend === 'kit';
  const live = Boolean(state?.capabilities) && !state.fallback && !state.error;
  const blocked = toolbarBusy(busy, state);
  const editing = state?.mode === 'edit';
  const history = kit && live && state.format === 'pptx' && state.history ? state.history : undefined;
  const step = can => ({ disabled: blocked || !editing || can !== true, ...blocked && { reason: 'busy' } });
  const denied = kit && state.access === 'denied';
  const saveReason = blocked && state.saveState !== 'saving' ? 'busy' : !editing ? 'view' : denied ? 'readOnly' : undefined;
  const kitLive = kit && live;
  return {
    // Save As writes a copy of what the tab shows, editing or reading; Open is offered even on a fallback preview,
    // to leave a file the editor could not open.
    slideshow: kitLive && state.format === 'pptx' && ['edit', 'view'].includes(state.mode)
      ? { disabled: blocked || state.ready !== true } : undefined,
    saveAs: kitLive ? { disabled: blocked, ...blocked && { reason: 'busy' } } : undefined,
    folder: state?.capabilities ? { disabled: blocked, ...blocked && { reason: 'busy' } } : undefined,
    undo: history && step(history.canUndo),
    redo: history && step(history.canRedo),
    save: live && state.capabilities.save !== false ? { disabled: saveReason !== undefined || state.saveState === 'saving', ...saveReason && { reason: saveReason },
      look: state.saveState === 'saving' ? 'saving' : state.saveState === 'failed' ? 'failed' : state.saveState === 'saved' ? 'saved' : editing && (state.dirty || state.unstaged) ? 'dirty' : 'clean' } : undefined,
    print: kit && live && PRINT_FORMATS.has(state.format) ? { disabled: blocked, ...blocked && { reason: 'busy' } } : undefined,
  };
}

/**
 * Modules this plugin keeps but does not show for now, with or without the tab's support.
 * `folder` (打开本地文件夹) is withdrawn from the plugin (0.0.37): a file opened from it, in a new tab or in place of
 * this one, is not part of the chat's context, so the model loses track of it; opening local files only makes sense in
 * the 工作台 or a desktop client. The open flow behind it (openFile, openPicked, BrowsePanel, OPEN-01, the Host's
 * probe/browse) stays, so it can come back by removing the id here.
 */
export const WITHDRAWN_MODULES = Object.freeze(['folder', 'undo', 'redo', 'account', 'skills']);

/** The modules this tab offers, in order: a module must be provided here, apply to this tab, and not be withdrawn. */
export const offeredModules = applies => MODULES.filter(id => applies[id] === true && !WITHDRAWN_MODULES.includes(id));

/**
 * Shown/hidden choices shared by every tab and suite. Stored as the hidden list so modules added later start visible;
 * storage that refuses or holds junk leaves the choices in memory for this page.
 */
export function createToolbarPrefs(storage = () => globalThis.localStorage) {
  const listeners = new Set();
  const read = () => {
    try {
      const store = storage();
      const stored = store?.getItem(STORAGE_KEY);
      const previous = store?.getItem('officesdk:toolbar:v3.1');
      const hidden = JSON.parse(stored ?? previous ?? store?.getItem('officesdk:toolbar') ?? '{}')?.hidden;
      // Ids this version does not know stay: a newer version (another tab, before a downgrade) may own them.
      if (!Array.isArray(hidden)) return Object.freeze([...DEFAULT_HIDDEN]);
      const ids = hidden.filter(id => typeof id === 'string');
      if (stored != null) return Object.freeze(ids);
      return Object.freeze([...new Set([...(previous == null ? DEFAULT_HIDDEN : NEWLY_HIDDEN), ...ids])]);
    } catch { return Object.freeze([...DEFAULT_HIDDEN]); }
  };
  let hidden = read();
  const publish = next => { hidden = Object.freeze(next); for (const listener of listeners) listener(); };
  return {
    key: STORAGE_KEY,
    hidden: () => hidden,
    subscribe(listener) { listeners.add(listener); return () => listeners.delete(listener); },
    toggle(id) {
      const next = hidden.includes(id) ? hidden.filter(other => other !== id) : [...hidden, id];
      try { storage()?.setItem(STORAGE_KEY, JSON.stringify({ hidden: next })); } catch {}
      publish(next);
    },
    /** Another window changed the stored choices. */
    reload() { publish(read()); },
  };
}

/**
 * Where a popover goes, in viewport pixels: its right edge on the trigger's right edge, 5px below it, kept 8px inside
 * both sides; a popover taller than the room below scrolls inside. `measured` is where a previous placement actually
 * landed: an ancestor with a transform moves fixed boxes, so the next placement corrects by that offset.
 */
export function placeMenu({ trigger, width, height, viewport, target, measured }) {
  const margin = 8;
  const maxWidth = Math.max(0, viewport.width - 2 * margin);
  const w = Math.min(width, maxWidth);
  let left = Math.min(trigger.right - w, viewport.width - margin - w);
  left = Math.max(margin, left);
  const top = trigger.bottom + 5;
  const maxHeight = Math.max(0, viewport.height - margin - top);
  const place = { left, top, maxWidth, maxHeight, height: Math.min(height, maxHeight) };
  if (!target || !measured) return place;
  return { ...place, left: place.left + target.left - measured.left, top: place.top + target.top - measured.top };
}

/** Keyboard in a menu of `count` items with `index` focused: the next index, and whether the key closes the menu. */
export function menuKey(key, index, count) {
  if (key === 'ArrowDown') return { index: (index + 1) % count };
  if (key === 'ArrowUp') return { index: (index - 1 + count) % count };
  if (key === 'Home') return { index: 0 };
  if (key === 'End') return { index: count - 1 };
  if (key === 'Enter' || key === ' ') return { index, toggle: true, close: 'restore' };
  if (key === 'Escape') return { index, close: 'restore' };
  if (key === 'Tab') return { index, close: 'leave' };
  return undefined;
}

/** Left/Right/Home/End across the toolbar's visible buttons; other keys do nothing. */
export function rovingKey(key, index, count) {
  if (key === 'ArrowRight') return (index + 1) % count;
  if (key === 'ArrowLeft') return (index - 1 + count) % count;
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return undefined;
}

/** Cancel keys for a popover: Esc, and ⌘. on macOS. */
export const cancelKey = event => event.key === 'Escape' || (event.key === '.' && event.metaKey === true);

// Names a file cannot take here: the Save As panel says so in place (macOS and Windows reserved characters).
const BAD_NAME = /[\/\\:*?"<>|\u0000-\u001f]/;

/** "report.docx" → "report - 副本.docx": the Save As panel's first suggestion, extension kept. */
export function copyName(name, suffix = ' - 副本') {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? `${name.slice(0, dot)}${suffix}${name.slice(dot)}` : `${name}${suffix}`;
}

/** Why a Save As name cannot be used ('empty' | 'characters' | 'dot'), or undefined when it can. */
export function nameProblem(stem) {
  const trimmed = String(stem ?? '').trim();
  if (!trimmed) return 'empty';
  if (BAD_NAME.test(trimmed)) return 'characters';
  if (trimmed.startsWith('.')) return 'dot';
  return undefined;
}

/** A path the desktop's file chooser gave: absolute on the Host's `platform` (a Windows drive or share there), an Office extension, no control characters. */
export const pickedPath = (path, platform) => typeof path === 'string' && isAbsolute(path, platform) && !/[\u0000-\u001f]/.test(path) && /\.(docx|xlsx|pptx)$/i.test(path) ? path : undefined;

/**
 * What a switch to another file needs first (OPEN-01): nothing without an editing session; a Kit tab not allowed to
 * write needs authorization when it has edits (a read-only Session can only Save As); otherwise a real save, which
 * also takes the editor's pending input (a clean-looking tab included, Core PPTX too).
 */
export function switchPlan(state, leased, held = false) {
  if (!leased || state?.mode !== 'edit') return 'switch';
  // The account holds edits back: a tab with edits stays (they cannot be saved now); a clean one may go.
  if (held) return state.dirty || state.unstaged ? 'account' : 'switch';
  if (state.backend === 'kit' && !state.authorized) {
    if (!state.dirty && !state.unstaged) return 'switch';
    return state.access === 'denied' ? 'readOnly' : 'authorize';
  }
  return 'save';
}

/** Saved for a switch: the save went through and nothing new arrived meanwhile. */
export const savedForSwitch = (state, revisionBefore, revisionAfter) => !state.dirty && !state.unstaged && revisionBefore === revisionAfter;

/** How often open document tabs look for a newer plugin release in the background. */
export const UPDATE_CHECK_MS = 600_000;

/**
 * One background release check for every open tab: the first header that subscribes starts it, the last one to go
 * stops it. A check answers through `fetch({ refresh })`; a failed check keeps the last known answer (a failure is
 * never "up to date"). Nothing here opens a popover or takes focus (N09): headers only re-render.
 */
export function createUpdateWatch({ fetch, scheduler = globalThis, intervalMs = UPDATE_CHECK_MS }) {
  const listeners = new Set();
  let value, timer, running;
  const publish = next => { value = next; for (const listener of listeners) listener(); };
  const check = (refresh = false) => {
    running ??= Promise.resolve().then(() => fetch({ refresh })).then(next => { if (next) publish(next); }, () => {}).finally(() => { running = undefined; });
    return running;
  };
  const tick = () => { void check(true); timer = scheduler.setTimeout(tick, intervalMs); };
  return {
    snapshot: () => value,
    check,
    subscribe(listener) {
      listeners.add(listener);
      if (listeners.size === 1) { void check(); timer = scheduler.setTimeout(tick, intervalMs); }
      return () => {
        listeners.delete(listener);
        if (!listeners.size) { scheduler.clearTimeout(timer); timer = undefined; }
      };
    },
  };
}
