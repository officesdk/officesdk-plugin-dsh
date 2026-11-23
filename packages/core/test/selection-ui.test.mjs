import test from 'node:test';
import assert from 'node:assert/strict';
import { SELECTION_TEXTS, createSelectionFeature, isSend, isShortcut, placeNear } from '../src/browser/selection-ui.mjs';
import { replaceInParagraph } from '../src/browser/selection-edit.mjs';
import { textChunks } from '../src/browser/selection.mjs';

/** The drawing, recorded: what is shown and the handlers the feature installs. */
function fakeUi() {
  const ui = { handlers: {}, texts: SELECTION_TEXTS.zh, calls: [], chipShown: false, panelOpen: false, resultShown: false, pending: undefined, hintText: undefined, draftOffered: false, resultText: undefined };
  Object.assign(ui, {
    owns: target => target === 'ui',
    showChip(point, options = {}) { ui.chipShown = true; ui.chip = { point, ...options }; ui.calls.push('chip'); },
    hideChip() { ui.chipShown = false; },
    openPanel() { ui.panelOpen = true; ui.chipShown = false; ui.hintText = undefined; ui.calls.push('panel'); },
    closePanel() { ui.panelOpen = false; },
    busy(on) { ui.busyNow = on; },
    hint(text, { draft = false } = {}) { ui.hintText = text; ui.draftOffered = draft; },
    showPending(box) { ui.pending = box; },
    showResult(point, text = SELECTION_TEXTS.zh.done) { ui.resultShown = true; ui.resultText = text; },
    hideResult() { ui.resultShown = false; },
    dispose() { ui.disposed = true; },
  });
  return ui;
}
function fakeDocument() {
  const listeners = [];
  return {
    listeners,
    addEventListener: (type, run) => listeners.push({ type, run }),
    removeEventListener: (type, run) => { const at = listeners.findIndex(item => item.type === type && item.run === run); if (at >= 0) listeners.splice(at, 1); },
    fire(type, event = {}) { const full = { type, target: 'doc', preventDefault() { full.prevented = true; }, stopImmediatePropagation() {}, ...event }; for (const item of listeners.filter(entry => entry.type === type)) item.run(full); return full; },
  };
}
/** Timers run by hand. */
function clock() {
  const timers = new Map(); let next = 1;
  return { schedule: (run, ms) => { timers.set(next, { run, ms }); return next++; }, unschedule: id => timers.delete(id), async flush(max = Infinity) { for (const [id, timer] of [...timers]) if (timer.ms <= max) { timers.delete(id); await timer.run(); } await new Promise(resolve => setImmediate(resolve)); } };
}
const SELECTION = { kind: 'docx', anchor: { paragraphs: [1, 1], start: 0, end: 3 }, original: '第一句', view: { first: 1, texts: ['第一句。第二句。'] } };

function setup({ selection = SELECTION, agent: extra = {} } = {}) {
  const ui = fakeUi(), document = fakeDocument(), time = clock(), sent = [];
  const agent = { applying: false, readSelection: async () => selection, locate: () => ({ left: 10, top: 20, width: 100, height: 20 }), rewriteSelection: async () => ({ reply: { applied: true }, revert: { anchor: 'after' }, located: {} }), revertSelection: async () => {}, ...extra };
  const feature = createSelectionFeature({ ui, agent, document, window: {}, notify: (event, payload) => sent.push({ event, ...payload }), schedule: time.schedule, unschedule: time.unschedule });
  return { ui, document, time, sent, agent, feature };
}

test('a selection shows the button; it opens the box; Esc and the box closing send nothing', async () => {
  const { ui, document, time, sent } = setup();
  document.fire('pointerup', { clientX: 300, clientY: 200 });
  await time.flush();
  assert.equal(ui.chipShown, true);
  assert.deepEqual(ui.chip.point.x, 300);
  ui.handlers.open();
  assert.equal(ui.panelOpen, true);
  ui.handlers.close();
  assert.equal(ui.panelOpen, false);
  // An empty request is not sent either.
  ui.handlers.open();
  await ui.handlers.submit('   ');
  assert.deepEqual(sent, []);
});

test('⌘J opens the box only while a selection can be rewritten; Enter during IME composition does not send', async () => {
  const { ui, document, time } = setup();
  assert.equal(document.fire('keydown', { key: 'j', metaKey: true }).prevented, undefined, 'no selection yet: the editor keeps the key');
  document.fire('pointerup', { clientX: 1, clientY: 1 });
  await time.flush();
  assert.equal(document.fire('keydown', { key: 'j', metaKey: true }).prevented, true);
  assert.equal(ui.panelOpen, true);
  assert.ok(isShortcut({ ctrlKey: true, key: 'J' }) && !isShortcut({ key: 'j' }) && !isShortcut({ metaKey: true, altKey: true, key: 'j' }));
  assert.ok(isSend({ key: 'Enter' }) && !isSend({ key: 'Enter', isComposing: true }) && !isSend({ key: 'Enter', keyCode: 229 }) && !isSend({ key: 'Enter', shiftKey: true }));
});

test('sending hands the selection to the parent and marks it waiting; × cancels it', async () => {
  const { ui, document, time, sent, feature } = setup();
  document.fire('pointerup', { clientX: 1, clientY: 1 });
  await time.flush();
  ui.handlers.open();
  const sending = ui.handlers.submit('改得更正式');
  assert.equal(sent.length, 1);
  const [request] = sent;
  assert.equal(request.event, 'selection-request');
  assert.deepEqual({ kind: request.kind, anchor: request.anchor, original: request.original, instruction: request.instruction }, { kind: 'docx', anchor: SELECTION.anchor, original: '第一句', instruction: '改得更正式' });
  assert.equal(request.view, undefined, 'drawing data stays in the page');
  feature.status({ requestId: request.requestId, ok: true, id: 'sel-00000001' });
  await sending;
  assert.equal(ui.panelOpen, false);
  assert.deepEqual(ui.pending, { left: 10, top: 20, width: 100, height: 20 });
  assert.equal(feature.pending.id, 'sel-00000001');
  // A dropped send for another selection leaves this one waiting; its own takes the mark down.
  feature.status({ dropped: 'sel-0000ffff' });
  assert.equal(feature.pending.id, 'sel-00000001');
  ui.handlers.cancelPending();
  assert.deepEqual(sent.at(-1), { event: 'selection-update', selection: 'sel-00000001', state: 'cancelled' });
  assert.equal(ui.pending, undefined);
});

test('a send that fails after the page heard it takes the waiting mark down', async () => {
  const { ui, document, time, sent, feature } = setup();
  document.fire('pointerup', {});
  await time.flush();
  ui.handlers.open();
  const sending = ui.handlers.submit('更简洁');
  feature.status({ requestId: sent.at(-1).requestId, ok: true, id: 'sel-00000009' });
  await sending;
  assert.ok(ui.pending);
  feature.status({ dropped: 'sel-00000009' });
  assert.equal(ui.pending, undefined);
  assert.equal(feature.pending, undefined);
});

test('a refused send keeps the box open with the reason; a draft in the chat box offers to put the request there', async () => {
  const { ui, document, time, sent, feature } = setup();
  document.fire('pointerup', {});
  await time.flush();
  ui.handlers.open();
  let sending = ui.handlers.submit('更简洁');
  feature.status({ requestId: sent.at(-1).requestId, ok: false, reason: 'draft' });
  await sending;
  assert.equal(ui.panelOpen, true);
  assert.equal(ui.hintText, SELECTION_TEXTS.zh.reasons.draft);
  assert.equal(ui.draftOffered, true);
  assert.equal(feature.pending, undefined);
  sending = ui.handlers.submit('更简洁', { draftOnly: true });
  assert.equal(sent.at(-1).draftOnly, true);
  feature.status({ requestId: sent.at(-1).requestId, ok: true, id: 'sel-00000002', drafted: true });
  await sending;
  assert.equal(ui.resultText, SELECTION_TEXTS.zh.drafted);
  assert.equal(feature.pending.id, 'sel-00000002');
  // No answer at all: the request times out as failed.
  document.fire('pointerup', {});
  await time.flush(1000);
  ui.handlers.open();
  sending = ui.handlers.submit('翻成英文');
  await time.flush();
  await sending;
  assert.equal(ui.hintText, SELECTION_TEXTS.zh.reasons.failed);
});

test('while the conversation runs the button is shown disabled; unsupported selections say why', async () => {
  const running = setup();
  running.feature.composer({ canSend: false, reason: 'running' });
  running.document.fire('pointerup', {});
  await running.time.flush();
  assert.equal(running.ui.chip.disabled, true);
  running.ui.handlers.open();
  assert.equal(running.ui.panelOpen, false);
  const table = setup({ selection: { unsupported: 'table' } });
  table.document.fire('pointerup', {});
  await table.time.flush();
  assert.deepEqual([table.ui.chip.disabled, table.ui.chip.reason], [true, 'table']);
  table.ui.handlers.open();
  assert.equal(table.ui.hintText, SELECTION_TEXTS.zh.reasons.table);
  await table.ui.handlers.submit('x');
  assert.deepEqual(table.sent, []);
  // While the AI writes nothing is offered.
  const busy = setup({ agent: { applying: true } });
  busy.document.fire('pointerup', {});
  await busy.time.flush();
  assert.equal(busy.ui.chipShown, false);
});

test('a rewrite clears the waiting mark and offers Keep / Undo; Undo reverts once, typing keeps', async () => {
  let reverted = 0;
  const { ui, feature, sent, document, time } = setup({ agent: { revertSelection: async revert => { assert.equal(revert.anchor, 'after'); reverted++; } } });
  const reply = await feature.rewrite({ id: 'sel-00000003', anchor: SELECTION.anchor, original: '第一句', text: '首句' });
  assert.deepEqual(reply, { applied: true });
  assert.equal(ui.resultShown, true);
  await ui.handlers.revert();
  assert.equal(reverted, 1);
  assert.deepEqual(sent.at(-1), { event: 'selection-update', selection: 'sel-00000003', state: 'reverted' });
  assert.equal(ui.resultText, SELECTION_TEXTS.zh.reverted);
  await ui.handlers.revert();
  assert.equal(reverted, 1, 'only once');
  await time.flush();
  await feature.rewrite({ id: 'sel-00000004' });
  document.fire('keydown', { key: 'a' });
  assert.deepEqual(sent.at(-1), { event: 'selection-update', selection: 'sel-00000004', state: 'kept' });
  assert.equal(ui.resultShown, false);
});

test('an undo the document no longer allows says so; a failed rewrite offers no undo', async () => {
  const failing = setup({ agent: { revertSelection: async () => { throw Object.assign(new Error('changed'), { code: 'source-changed' }); } } });
  await failing.feature.rewrite({ id: 'sel-00000005' });
  await failing.ui.handlers.revert();
  assert.equal(failing.ui.resultText, SELECTION_TEXTS.zh.revertFailed);
  const stopped = setup({ agent: { rewriteSelection: async () => { throw Object.assign(new Error('Stopped by the user.'), { code: 'cancelled' }); } } });
  await assert.rejects(stopped.feature.rewrite({ id: 'sel-00000006' }), error => error.code === 'cancelled');
  assert.equal(stopped.ui.resultShown, false);
  assert.equal(stopped.feature.result, undefined);
});

test('boxes stay inside the window', () => {
  assert.deepEqual(placeNear({ x: 990, y: 750 }, { width: 120, height: 30 }, { innerWidth: 1000, innerHeight: 768 }), { left: 872, top: 712 });
  assert.deepEqual(placeNear({ x: 100, y: 100 }, { width: 120, height: 30 }, { innerWidth: 1000, innerHeight: 768 }), { left: 108, top: 108 });
});

/** One paragraph's text and its Writer subranges (code points), every write recorded. */
function fakeWordParagraph(text) {
  const state = { text, writes: [] };
  const chars = () => Array.from(state.text);
  const subrange = (start, length) => ({
    insertText(value, where) {
      const all = chars();
      if (where === 'Replace') all.splice(start, length, ...Array.from(value));
      else if (where === 'After') all.splice(start + length, 0, ...Array.from(value));
      else all.splice(start, 0, ...Array.from(value));
      state.text = all.join('');
      state.writes.push([where, value]);
    },
    delete() { const all = chars(); all.splice(start, length); state.text = all.join(''); state.writes.push(['Delete']); },
  });
  const paragraph = { getRange: () => ({ getSubrange: subrange, getRange: () => subrange(0, 0) }) };
  const Word = { run: async fn => fn({ document: { body: { paragraphs: { load() {}, items: [paragraph] } } }, sync: async () => {} }) };
  return { Word, state };
}

test('an in-paragraph rewrite replaces only the selection, in growing parts, counting emoji as one', async () => {
  const { Word, state } = fakeWordParagraph('A😀B中文加粗尾巴');
  const seen = [];
  const text = '这是一段比二十四个字还要更长一些的新内容😀，用来分段写入结尾';
  await replaceInParagraph(Word, { index: 0, start: 3, length: 2, text, chunk: textChunks, onWrite: part => seen.push(part) });
  assert.equal(state.text, `A😀B${text}加粗尾巴`);
  assert.ok(seen.length > 1 && seen.at(-1) === text);
  // Taking it back writes the original in place again.
  await replaceInParagraph(Word, { index: 0, start: 3, length: Array.from(text).length, text: '中文' });
  assert.equal(state.text, 'A😀B中文加粗尾巴');
  // Rewriting to nothing deletes; taking that back inserts at the position.
  await replaceInParagraph(Word, { index: 0, start: 3, length: 2, text: '' });
  assert.equal(state.text, 'A😀B加粗尾巴');
  await replaceInParagraph(Word, { index: 0, start: 3, length: 0, text: '中文' });
  assert.equal(state.text, 'A😀B中文加粗尾巴');
  await assert.rejects(replaceInParagraph(Word, { index: 0, start: 0, length: 1, text: 'x', signal: AbortSignal.abort() }), error => error.code === 'cancelled');
});
