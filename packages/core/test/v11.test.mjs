import test from 'node:test';
import assert from 'node:assert/strict';
import { SELECTION_TIMING, selectionAt, selectionMs, smoothstep, textChunks } from '../src/browser/selection.mjs';
import { STATUS_TEXTS, agentLabel, conclusionIndex, docxStatus, pptStatus, xlsxStatus } from '../src/browser/agent-status.mjs';
import { createDocOverlay } from '../src/browser/doc-overlay.mjs';

test('the selection gesture: cursor first, then the outline pulled open with smoothstep, then writing', () => {
  const { leadMs, dragMs } = SELECTION_TIMING;
  assert.deepEqual([leadMs, dragMs, SELECTION_TIMING.tailMs], [150, 500, 200]);
  assert.deepEqual(selectionAt(1000, 1000), { phase: 'cursor', progress: 0 });
  assert.deepEqual(selectionAt(1000, 1000 + leadMs - 1), { phase: 'cursor', progress: 0 });
  const half = selectionAt(1000, 1000 + leadMs + dragMs / 2);
  assert.equal(half.phase, 'selecting');
  assert.equal(half.progress, 0.5, 'half way through the drag the outline is half open (width and height together)');
  assert.ok(selectionAt(1000, 1000 + leadMs + dragMs / 4).progress < 0.25, 'eased in');
  assert.deepEqual(selectionAt(1000, 1000 + leadMs + dragMs), { phase: 'done', progress: 1 });
  assert.deepEqual(selectionAt(1000, 1000, { animate: false }), { phase: 'done', progress: 1 }, 'reduced motion: the full outline at once');
  assert.equal(selectionMs(), leadMs + dragMs);
  assert.equal(selectionMs({ animate: false }), 0);
  assert.equal(smoothstep(-1), 0); assert.equal(smoothstep(2), 1);
});

test('text is written in at most four growing parts, short text at once', () => {
  assert.deepEqual(textChunks('短文字'), ['短文字']);
  const long = '一二三四五六七八九十'.repeat(6);
  const parts = textChunks(long);
  assert.equal(parts.length, 4);
  assert.equal(parts.at(-1), long);
  assert.ok(parts.every((part, index) => long.startsWith(part) && (index === 0 || part.length > parts[index - 1].length)));
  assert.equal(textChunks('x'.repeat(30)).length, 2, 'each part at least 12 characters');
  assert.deepEqual(textChunks(''), ['']);
});

test('status texts come from what a step does, in the handoff\'s words', () => {
  assert.equal(agentLabel('zh', 'selecting'), 'Agent · 框选中');
  assert.equal(agentLabel('en', 'fill'), 'Agent · Filling data');
  assert.equal(STATUS_TEXTS.zh.content, '生成内容');
  assert.equal(pptStatus({ op: 'addSlide' }), 'page');
  assert.equal(pptStatus({ op: 'setText' }, { shape: { placeholder: 'Title' } }), 'title');
  assert.equal(pptStatus({ op: 'addTextBox', font: { size: 36 } }), 'title');
  assert.equal(pptStatus({ op: 'addTextBox' }), 'content');
  assert.equal(pptStatus({ op: 'addTextBox' }, { conclusion: true }), 'conclusion');
  assert.equal(pptStatus({ op: 'addShape' }), 'shape');
  assert.equal(pptStatus({ op: 'deleteShape' }), 'delete');
  assert.equal(docxStatus({ op: 'insertParagraph', style: 'heading2' }), 'title');
  assert.equal(docxStatus({ op: 'insertParagraph', list: 'bullet' }), 'points');
  assert.equal(docxStatus({ op: 'insertParagraph' }), 'body');
  assert.equal(xlsxStatus({ op: 'setValues', values: [['a', 1]] }), 'fill');
  assert.equal(xlsxStatus({ op: 'setValues', values: [['=SUM(A1:A2)', null]] }), 'formula');
  assert.equal(xlsxStatus({ op: 'addSheet' }), 'sheet');
  assert.equal(xlsxStatus({ op: 'formatRange' }), 'format');
  const steps = [{ op: 'setText', slideId: 's', geometry: { top: 10 } }, { op: 'addTextBox', slideId: 's', geometry: { top: 300 } }, { op: 'addTextBox', slideId: 's', geometry: { top: 200 } }, { op: 'addTextBox', slideId: 't', geometry: { top: 900 } }];
  assert.equal(conclusionIndex(steps, 's', step => step.op === 'setText' ? { placeholder: 'Title' } : undefined), 1, 'the lowest non-title text on that slide');
});

/** Just enough DOM for the overlay's own elements; no target is ever located, so no outline is drawn. */
function fakePage() {
  const element = () => ({ style: {}, dataset: {}, append() {}, remove() {}, setAttribute() {}, getBoundingClientRect: () => ({ width: 0 }) });
  return { defaultView: { requestAnimationFrame: () => 1, cancelAnimationFrame() {}, performance: { now: () => Date.now() }, addEventListener() {}, removeEventListener() {} },
    body: { append() {} }, createElement: element, createElementNS: element, addEventListener() {}, removeEventListener() {} };
}

test('a track resolves its selection only after the gesture (or at once without motion) and goes after its tail', async () => {
  const waits = [];
  const overlay = createDocOverlay(fakePage(), { animate: () => true, wait: ms => { waits.push(ms); return Promise.resolve(); } });
  await overlay.select('a', () => undefined, { label: 'Agent · 撰写正文', selecting: 'Agent · 框选中' });
  await overlay.select('b', () => undefined, { label: 'Agent · 梳理要点', selecting: 'Agent · 框选中' });
  assert.deepEqual(waits, [650, 650], 'writing waits for cursor + drag');
  assert.deepEqual(overlay.tracks, ['a', 'b'], 'several tracks at once');
  overlay.release('a');
  await new Promise(resolve => setTimeout(resolve, 260));
  assert.deepEqual(overlay.tracks, ['b']);
  const still = createDocOverlay(fakePage(), { animate: () => false, wait: () => { throw new Error('no wait without motion'); } });
  await still.select('a', () => undefined, { label: 'x' });
  overlay.busy(false);
  assert.deepEqual(overlay.tracks, [], 'ending the edit clears every track');
  overlay.dispose(); still.dispose();
});
