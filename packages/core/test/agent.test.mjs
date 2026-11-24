import test from 'node:test';
import assert from 'node:assert/strict';
import { PPTX_LIMITS, describePptxSteps, summarizeDeck, validatePptxOps } from '../src/index.mjs';
import { createFrameChannel, frameCommand, serveAgent } from '../src/browser/index.mjs';

const deck = {
  slides: [
    { id: 's1', shapes: [{ id: 'title', name: 'Title 1', type: 'GeometricShape', left: 10, top: 10, width: 300, height: 40, text: 'Roadmap 2026' }, { id: 'pic', name: 'Picture', type: 'Image', left: 0, top: 0, width: 50, height: 50 }] },
    { id: 's2', shapes: [{ id: 'body', name: '', type: 'GeometricShape', left: 20, top: 80, width: 400, height: 200, text: 'alpha beta alpha' }] },
    { id: 's3', shapes: [] }, { id: 's4', shapes: [] },
  ],
};

test('text operations resolve against the live deck and chain within one batch', () => {
  const { steps, danger } = validatePptxOps([
    { op: 'replaceText', slideId: 's1', shapeId: 'title', find: '2026', replace: '2027' },
    { op: 'replaceText', slideId: 's1', shapeId: 'title', find: 'Roadmap', replace: '路线图' },
    { op: 'setText', slideId: 's2', shapeId: 'body', text: 'one\ntwo' },
  ], deck);
  assert.equal(danger, false);
  assert.deepEqual(steps.map(step => [step.before, step.text]), [['Roadmap 2026', 'Roadmap 2027'], ['Roadmap 2027', '路线图 2027'], ['alpha beta alpha', 'one\ntwo']]);
  assert.deepEqual(describePptxSteps(steps, deck)[0], { op: 'setText', where: '#1 · Title 1', before: '2026', after: '2027' });
});

test('targets must exist, quotes must be unique, values must be well formed', () => {
  const reject = (ops, pattern) => assert.throws(() => validatePptxOps(ops, deck), error => error.code === 'invalid-edit' && pattern.test(error.message));
  reject([], /at least one/);
  reject([{ op: 'eval', slideId: 's1' }], /unknown operation/);
  reject([{ op: 'setText', slideId: 'nope', shapeId: 'title', text: 'x' }], /slide "nope"/);
  reject([{ op: 'setText', slideId: 's1', shapeId: 'body', text: 'x' }], /shape "body" is not on slide s1/);
  reject([{ op: 'setText', slideId: 's1', shapeId: 'pic', text: 'x' }], /has no text/);
  reject([{ op: 'replaceText', slideId: 's2', shapeId: 'body', find: 'alpha', replace: 'x' }], /more than once/);
  reject([{ op: 'replaceText', slideId: 's2', shapeId: 'body', find: 'gamma', replace: 'x' }], /does not occur/);
  reject([{ op: 'setFill', slideId: 's1', shapeId: 'title', color: 'red' }], /#RRGGBB/);
  reject([{ op: 'setFont', slideId: 's1', shapeId: 'title' }], /at least one/);
  reject([{ op: 'setFont', slideId: 's1', shapeId: 'title', size: 0 }], /size/);
  reject([{ op: 'setGeometry', slideId: 's1', shapeId: 'title', width: -1 }], /width/);
  reject([{ op: 'deleteShape', slideId: 's1', shapeId: 'title' }, { op: 'setText', slideId: 's1', shapeId: 'title', text: 'x' }], /Operation 2: shape "title"/);
  reject([{ op: 'moveSlide', slideId: 's1', index: 4 }], /index/);
  reject([{ op: 'setText', slideId: 's1', shapeId: 'title', text: 'x'.repeat(PPTX_LIMITS.maxText + 1) }], /longer/);
  reject(Array.from({ length: PPTX_LIMITS.maxOps + 1 }, () => ({ op: 'addSlide' })), /At most/);
  assert.throws(() => validatePptxOps([{ op: 'deleteSlide', slideId: 's1' }, { op: 'deleteSlide', slideId: 's2' }, { op: 'deleteSlide', slideId: 's3' }, { op: 'deleteSlide', slideId: 's4' }], deck), /at least one slide/);
});

test('deleting content or touching most slides needs a deliberate confirmation', () => {
  assert.deepEqual(validatePptxOps([{ op: 'deleteShape', slideId: 's1', shapeId: 'pic' }], deck).reasons, ['deletes']);
  const wide = validatePptxOps(['s1', 's2', 's3'].map(slideId => ({ op: 'addTextBox', slideId, text: 'x', left: 0, top: 0, width: 10, height: 10 })), deck);
  assert.equal(wide.danger, true);
  assert.deepEqual(wide.reasons, ['wide']);
  assert.equal(validatePptxOps([{ op: 'setFill', slideId: 's1', shapeId: 'title', color: '#FFAA00' }], deck).danger, false);
});

test('the model reads ids, boxes and text within a character budget', () => {
  const summary = summarizeDeck(deck, { maxChars: 15 });
  assert.equal(summary.truncated, true);
  assert.deepEqual(summary.slides[0].shapes[0], { id: 'title', name: 'Title 1', type: 'GeometricShape', box: [10, 10, 300, 40], text: 'Roadmap 2026' });
  assert.equal(summary.slides[0].shapes[1].text, undefined, 'shapes without text carry no text field');
  assert.equal(summary.slides[1].shapes[0].text, 'alp');
  assert.equal(summarizeDeck(deck).truncated, undefined);
});

function frames() {
  const parent = { posted: [], postMessage(message) { this.posted.push(message); } };
  const listeners = new Set();
  const window = { parent, location: { origin: 'http://h' }, addEventListener: (_, fn) => listeners.add(fn), removeEventListener: (_, fn) => listeners.delete(fn) };
  const channel = createFrameChannel({ window, id: 'L', revision: 1, parentOrigin: 'http://h' });
  const send = message => { for (const fn of listeners) fn({ source: parent, origin: 'http://h', data: message }); };
  return { channel, parent, send };
}
const tick = () => new Promise(resolve => setTimeout(resolve, 0));

test('agent commands are validated on both sides and answered once per request', async () => {
  assert.throws(() => frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'bad id', op: 'inspect' }), /Invalid agent/);
  assert.throws(() => frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'r1', op: 'run-script' }), /Invalid agent/);
  const { channel, parent, send } = frames();
  let started;
  serveAgent(channel, {
    inspect: async () => ({ slides: [] }),
    apply: async (_, signal) => { started = true; await new Promise((resolve, reject) => signal.addEventListener('abort', () => reject(Object.assign(new Error('stopped'), { code: 'cancelled' })))); },
  });
  send(frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'r1', op: 'inspect' }));
  send({ ...frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'r1', op: 'inspect' }), op: 'eval' });
  send(frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'r2', op: 'preview' }));
  send(frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'r3', op: 'apply', args: { ops: [] } }));
  await tick();
  assert.equal(started, true);
  send(frameCommand({ id: 'L', revision: 1, command: 'agent', requestId: 'r4', op: 'cancel', args: { target: 'r3' } }));
  await tick();
  const results = parent.posted.filter(message => message.type === 'officesdk:agent-result').map(({ requestId, ok, result, code }) => ({ requestId, ok, ...(result && { result }), ...(code && { code }) })).sort((a, b) => a.requestId.localeCompare(b.requestId));
  assert.deepEqual(results, [
    { requestId: 'r1', ok: true, result: { slides: [] } },
    { requestId: 'r2', ok: false, code: 'unsupported-operation' },
    { requestId: 'r3', ok: false, code: 'cancelled' },
  ]);
});
