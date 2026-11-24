import test from 'node:test';
import assert from 'node:assert/strict';
import { acceptFrameMessage, createFrameChannel, frameCommand, resolveMeta } from '../src/browser/index.mjs';

const meta = { v: 1, id: 'abc', routePrefix: '/shimo', contentUrl: '/shimo/doc/abc/content', resourcesBaseUrl: '/shimo/doc/abc/resources/', runtimeBaseUrl: '/shimo/runtime/', sdkApiBaseUrl: '/shimo/doc/abc/sdk/v2/api' };

test('meta URLs keep the application mount and stay same-origin', () => {
  const resolved = resolveMeta(meta, 'dsh-app://app/mount/shimo/doc/abc/index.html?revision=1');
  assert.equal(resolved.contentUrl, 'dsh-app://app/mount/shimo/doc/abc/content');
  assert.equal(resolved.stageUrl, undefined);
  const plain = resolveMeta({ ...meta, stageUrl: '/shimo/doc/abc/stage/' }, 'http://127.0.0.1:5/shimo/doc/abc/index.html');
  assert.equal(plain.stageUrl, 'http://127.0.0.1:5/shimo/doc/abc/stage/');
  assert.throws(() => resolveMeta({ ...meta, routePrefix: '/other' }, 'http://h/shimo/doc/abc/index.html'), /prefix/);
  assert.throws(() => resolveMeta({ ...meta, contentUrl: '//evil.example/x' }, 'http://h/shimo/doc/abc/index.html'), /External/);
});

test('parent accepts only its frame, origin, id, current revision and known v1 events', () => {
  const frame = { contentWindow: {} };
  const ok = { source: frame.contentWindow, origin: 'http://h', data: { type: 'officesdk:ready', v: 1, id: 'abc', revision: 2 } };
  const expected = { frame, origin: 'http://h', id: 'abc', revision: 2 };
  assert.equal(acceptFrameMessage(ok, expected).event, 'ready');
  assert.equal(acceptFrameMessage({ ...ok, source: {} }, expected), undefined);
  assert.equal(acceptFrameMessage({ ...ok, origin: 'http://evil' }, expected), undefined);
  for (const data of [{ ...ok.data, v: 2 }, { ...ok.data, revision: 1 }, { ...ok.data, id: 'x' }, { ...ok.data, type: 'legacy:ready' }, { ...ok.data, type: 'officesdk:unknown' }]) {
    assert.equal(acceptFrameMessage({ ...ok, data }, expected), undefined, JSON.stringify(data));
  }
  // Only the officesdk: prefix is heard; commands go out under it.
  assert.equal(acceptFrameMessage({ ...ok, data: { ...ok.data, type: 'other:ready' } }, expected), undefined);
  assert.equal(frameCommand({ id: 'abc', revision: 2, command: 'save' }).type, 'officesdk:command');
});

test('iframe channel applies rebinds and ignores commands for other revisions or senders', () => {
  const listeners = [];
  const posted = [];
  const parent = { postMessage: (message, origin) => posted.push([message, origin]) };
  const window = { parent, location: { origin: 'http://h' }, addEventListener: (_, fn) => listeners.push(fn), removeEventListener() {} };
  const channel = createFrameChannel({ window, id: 'abc', revision: 1, parentOrigin: 'http://h' });
  const seen = [];
  channel.onCommand(message => seen.push(message.command));
  const send = (data, source = parent, origin = 'http://h') => listeners.forEach(fn => fn({ data, source, origin }));
  send(frameCommand({ id: 'abc', revision: 1, command: 'save' }));
  send(frameCommand({ id: 'abc', revision: 1, command: 'save' }), {});
  send(frameCommand({ id: 'abc', revision: 1, command: 'save' }), parent, 'http://evil');
  send(frameCommand({ id: 'abc', revision: 1, command: 'rebind', newRevision: 5 }));
  send(frameCommand({ id: 'abc', revision: 1, command: 'flush' }));
  send(frameCommand({ id: 'abc', revision: 5, command: 'flush' }));
  assert.deepEqual(seen, ['save', 'rebind', 'flush']);
  assert.equal(channel.revision, 5);
  channel.notify('saved', { version: 'v2' });
  assert.deepEqual(posted.at(-1), [{ version: 'v2', type: 'officesdk:saved', v: 1, id: 'abc', revision: 5 }, 'http://h']);
  assert.throws(() => channel.notify('bogus'), /Unknown/);
  assert.throws(() => frameCommand({ id: 'abc', revision: 1, command: 'delete' }), /Unknown/);
});

const fakeTimers = () => {
  let now = 0, next = 1;
  const tasks = new Map();
  return {
    now: () => now,
    setTimeout(fn, ms) { const id = next++; tasks.set(id, { fn, at: now + ms }); return id; },
    clearTimeout(id) { tasks.delete(id); },
    setInterval(fn, ms) { const id = next++; tasks.set(id, { fn, at: now + ms, every: ms }); return id; },
    clearInterval(id) { tasks.delete(id); },
    async advance(ms) {
      const end = now + ms;
      for (;;) {
        const [id, task] = [...tasks].filter(([, t]) => t.at <= end).sort((a, b) => a[1].at - b[1].at)[0] ?? [];
        if (!task) break;
        now = task.at;
        if (task.every) task.at += task.every; else tasks.delete(id);
        task.fn();
        await new Promise(resolve => setImmediate(resolve));
      }
      now = end;
    },
  };
};

test('hand edits are staged right away, at most one stage per gap, one at a time', async () => {
  const { createEditSync } = await import('../src/browser/presentation-agent.mjs');
  const timers = fakeTimers();
  let revision = 0, manual = 0;
  const staged = [];
  const sync = createEditSync({ revision: () => revision, stage: async () => { staged.push(revision); }, onManual: () => manual++, timers, now: timers.now });
  await timers.advance(1000);
  assert.deepEqual(staged, [], 'nothing changed, nothing staged');
  revision = 1; await timers.advance(150);
  assert.deepEqual(staged, [1], 'the first change is staged on the next poll');
  revision = 2; await timers.advance(200);
  assert.deepEqual(staged, [1], 'the next waits for the gap');
  await timers.advance(300);
  assert.deepEqual([staged, manual], [[1, 2], 2]);
  revision = 3;
  assert.equal(await sync.flush(), true);
  assert.deepEqual(staged, [1, 2, 3], 'flush stages a change the poll has not seen yet');
  assert.equal(await sync.flush(), true);
  assert.deepEqual(staged, [1, 2, 3], 'nothing pending');
  sync.dispose();
  revision = 4; await timers.advance(2000);
  assert.deepEqual(staged, [1, 2, 3]);
});

test('an AI edit runs after pending hand edits, and its own revisions are not the user\'s', async () => {
  const { createEditSync } = await import('../src/browser/presentation-agent.mjs');
  const timers = fakeTimers();
  let revision = 0, busy = false, manual = 0;
  const order = [];
  const sync = createEditSync({ revision: () => revision, stage: async () => { order.push(`stage@${revision}`); }, onManual: () => manual++, paused: () => busy, timers });
  revision = 1;
  busy = true;
  const result = await sync.exclusive(async () => { order.push('ai'); revision = 5; await timers.advance(1000); return 'done'; });
  busy = false;
  assert.equal(result, 'done');
  assert.deepEqual(order, ['stage@1', 'ai'], 'the hand edit was staged first');
  await timers.advance(1000);
  assert.deepEqual([order.length, manual, sync.changed()], [2, 1, false], 'the AI\'s revision 5 is not a hand edit');
  revision = 6;
  assert.equal(sync.changed(), true);
  sync.dispose();
});

test('a failed stage is reported, stays pending and is retried with backoff', async () => {
  const { createEditSync } = await import('../src/browser/presentation-agent.mjs');
  const timers = fakeTimers();
  let revision = 0, fail = true;
  const errors = [], staged = [];
  const sync = createEditSync({ revision: () => revision, stage: async () => { if (fail) throw new Error('offline'); staged.push(revision); }, onError: error => errors.push(error.message), timers, now: timers.now });
  revision = 1;
  assert.equal(await sync.flush(), false, 'flush says the Host does not hold the latest bytes');
  assert.deepEqual([errors, staged, sync.changed()], [['offline'], [], true]);
  await timers.advance(1000);
  assert.deepEqual(errors, ['offline', 'offline'], 'retried after 1 s');
  fail = false;
  await timers.advance(2000);
  assert.deepEqual([staged, sync.changed()], [[1], false], 'retried after 2 s and recovered');
  sync.dispose();
});

test('an AI edit that staged the whole deck clears a pending hand edit', async () => {
  const { createEditSync } = await import('../src/browser/presentation-agent.mjs');
  const timers = fakeTimers();
  let revision = 0, fail = true;
  const sync = createEditSync({ revision: () => revision, stage: async () => { if (fail) throw new Error('offline'); }, timers, now: timers.now });
  revision = 1;
  await sync.exclusive(async () => { revision = 2; return 'ok'; });
  assert.equal(sync.changed(), false);
  sync.dispose();
});

test('a failed AI edit does not fail the next one', async () => {
  const { createEditSync } = await import('../src/browser/presentation-agent.mjs');
  const timers = fakeTimers();
  const sync = createEditSync({ revision: () => 0, stage: async () => {}, timers });
  await assert.rejects(sync.exclusive(async () => { throw new Error('invalid plan'); }), /invalid plan/);
  assert.equal(await sync.exclusive(async () => 'next'), 'next');
  sync.dispose();
});

test('while the AI applies, a flush neither counts its writes as hand edits nor stages before it ends', async () => {
  const { createEditSync } = await import('../src/browser/presentation-agent.mjs');
  const timers = fakeTimers();
  let revision = 0, busy = false, manual = 0;
  const staged = [];
  const sync = createEditSync({ revision: () => revision, stage: async () => { staged.push(revision); }, onManual: () => manual++, paused: () => busy, timers, now: timers.now });
  busy = true;
  let finish;
  const applying = sync.exclusive(() => new Promise(resolve => { finish = resolve; }));
  await new Promise(resolve => setImmediate(resolve));
  revision = 7; // the AI types
  const flushing = sync.flush(); // blur, or the panel's Save
  await timers.advance(1000);
  assert.deepEqual([staged, manual], [[], 0], 'nothing staged or counted while the AI runs');
  finish('ok'); busy = false;
  await applying;
  assert.equal(await flushing, true, 'the flush resolves after the AI edit, which staged the whole deck');
  await timers.advance(1000);
  assert.deepEqual(staged, [], 'no stray hand-edit stage after the AI edit');
  sync.dispose();
});
