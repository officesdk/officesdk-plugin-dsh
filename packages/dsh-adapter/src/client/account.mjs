/**
 * Account and credits for the P area (design package v2.9 §8 AUTH-01, CREDIT-01). The real account and billing
 * services are not defined yet (docs/04 Q01–Q06): this is the rules and a mock source behind a preview switch, off
 * by default, so the interface can be built and tried before it is wired to the Host.
 */
export const PREVIEW_KEY = 'officesdk:account-preview';
export const MOCK_KEY = 'officesdk:account-mock';
const STATUSES = new Set(['connected', 'disconnected', 'expired']);
const DEFAULT = Object.freeze({ status: 'connected', email: 'user@example.com', credits: 670 });

/** A mock account from its stored JSON; junk falls back to the default, unknown states to disconnected. */
export function parseAccount(text, warn = () => {}) {
  if (text === null || text === undefined || text === '') return DEFAULT;
  let value;
  try { value = JSON.parse(text); } catch { warn('account mock is not JSON; using the default'); return DEFAULT; }
  if (!value || typeof value !== 'object') return DEFAULT;
  return Object.freeze({
    status: STATUSES.has(value.status) ? value.status : 'disconnected',
    email: typeof value.email === 'string' ? value.email : undefined,
    credits: typeof value.credits === 'number' && Number.isFinite(value.credits) && value.credits >= 0 ? value.credits : null,
  });
}

/**
 * What the account allows. Not connected (or expired): no editing and no saving, whatever the credits — the account
 * condition comes first. Connected with 0 credits: saving existing edits stays possible; only paid features stop.
 * Unknown usage is not 0.
 */
export function accountGate(account) {
  if (account?.status !== 'connected') return { edit: false, save: false, paid: 'no', reason: 'account' };
  const paid = account.credits === null || account.credits === undefined ? 'unknown' : account.credits > 0 ? 'yes' : 'no';
  return { edit: true, save: true, paid };
}

/**
 * The account every tab shares: the preview switch and the mock, read live. Changes made here reach this window's
 * tabs at once (a `storage` event never fires in the window that wrote) and other windows through storage.
 */
export function createAccountSource(storage = () => globalThis.localStorage, warn = () => {}) {
  const listeners = new Set();
  const read = key => { try { return storage()?.getItem(key) ?? null; } catch { return null; } };
  let current;
  const load = () => { current = Object.freeze({ preview: read(PREVIEW_KEY) === '1', account: parseAccount(read(MOCK_KEY), warn), version: (current?.version ?? 0) + 1 }); };
  load();
  const publish = () => { load(); for (const listener of listeners) listener(current); };
  return {
    keys: [PREVIEW_KEY, MOCK_KEY],
    snapshot: () => current,
    subscribe(listener) { listeners.add(listener); return () => listeners.delete(listener); },
    /** Another window wrote one of the keys. */
    reload: publish,
    /** The mock's "connect": the account becomes connected here and in other windows. */
    connect() {
      try { storage()?.setItem(MOCK_KEY, JSON.stringify({ ...current.account, status: 'connected' })); } catch {}
      publish();
    },
    refresh: publish,
    /** The three-dot menu's 个人中心（预览） checkbox. */
    setPreview(on) {
      try { on ? storage()?.setItem(PREVIEW_KEY, '1') : storage()?.removeItem(PREVIEW_KEY); } catch {}
      publish();
    },
    /** A demo state chosen in the account menu. */
    setMock(account) {
      try { storage()?.setItem(MOCK_KEY, JSON.stringify(account)); } catch {}
      publish();
    },
  };
}

/** The account menu's demo states, in order; `matches` tells which one the current mock is. */
export const DEMO_STATES = Object.freeze([
  { id: 'connected', account: { status: 'connected', email: 'user@example.com', credits: 670 } },
  { id: 'disconnected', account: { status: 'disconnected', email: 'user@example.com', credits: 670 } },
  { id: 'expired', account: { status: 'expired', email: 'user@example.com', credits: 670 } },
  { id: 'zero', account: { status: 'connected', email: 'user@example.com', credits: 0 } },
  { id: 'unknown', account: { status: 'connected', email: 'user@example.com', credits: null } },
]);
export function demoState(account) {
  if (account?.status !== 'connected') return account?.status === 'expired' ? 'expired' : 'disconnected';
  return account.credits === null || account.credits === undefined ? 'unknown' : account.credits === 0 ? 'zero' : 'connected';
}

/** Whether the account holds this tab's edits back: only with the preview on and the account not connected. */
export const accountHeld = snapshot => snapshot?.preview === true && accountGate(snapshot.account).reason === 'account';

/**
 * The toolbar while the account holds edits back: Save, Save As and undo/redo off, the pen off, the eye off while
 * there are edits to write (leaving edit saves them); Open and Print stay. Outside a hold the buttons are unchanged.
 */
export function heldButtons(edits, modes, state, held) {
  if (!held) return { edits, modes };
  const off = button => button && { ...button, disabled: true, reason: 'account' };
  const dirty = state?.dirty === true || Boolean(state?.unstaged);
  return {
    edits: { ...edits, undo: off(edits.undo), redo: off(edits.redo), save: edits.save && { ...off(edits.save), look: edits.save.look }, saveAs: off(edits.saveAs) },
    modes: modes && {
      edit: modes.edit.pressed ? modes.edit : off(modes.edit),
      view: !modes.view.pressed && dirty ? off(modes.view) : modes.view,
    },
  };
}

/** Keys in the account menu: move between rows, activate one, or close. */
export function accountMenuKey(key, index, count) {
  if (key === 'ArrowDown') return { index: (index + 1) % count };
  if (key === 'ArrowUp') return { index: (index - 1 + count) % count };
  if (key === 'Home') return { index: 0 };
  if (key === 'End') return { index: count - 1 };
  if (key === 'Enter' || key === ' ') return { index, activate: true };
  if (key === 'Escape') return { index, close: 'restore' };
  if (key === 'Tab') return { index, close: 'leave' };
  return undefined;
}
