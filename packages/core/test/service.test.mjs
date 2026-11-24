import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, readdir, symlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { ERROR_STATUS, SuiteError, normalizeLimits } from '../src/index.mjs';
import { fakeConverter, fixture, peerOf, scope, signal, taskDirs, zip } from './helpers.mjs';

const code = expected => error => { assert.equal(error.code, expected, error.message); return true; };

test('every error code has a status and unknown codes are refused', () => {
  for (const [name, status] of Object.entries(ERROR_STATUS)) assert.equal(new SuiteError(name, 'x').status, status);
  assert.throws(() => new SuiteError('nope', 'x'), TypeError);
});

test('limits are positive safe integers, nested per operation and format', () => {
  const limits = normalizeLimits({ timeoutMs: { import: { pptx: 5 } } });
  assert.equal(limits.timeoutMs.import.pptx, 5);
  assert.equal(limits.timeoutMs.import.docx, 60_000);
  assert.equal(limits.timeoutMs.export.pptx, 180_000);
  for (const bad of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1, '2']) assert.throws(() => normalizeLimits({ maxLeases: bad }), TypeError);
  assert.throws(() => normalizeLimits({ unknown: 1 }), TypeError);
  assert.throws(() => normalizeLimits({ timeoutMs: { export: { xlsx: 0 } } }), TypeError);
});

for (const format of ['docx', 'xlsx', 'pptx']) {
  test(`${format}: import, edit, save and export round trip on a fake converter`, async t => {
    const f = await fixture(t);
    const peer = peerOf();
    const handle = await f.service.open(scope, `a.${format}`, 1, signal(), peer, { mode: 'edit' });
    assert.match(handle.id, /^[a-f0-9]{48}$/);
    assert.equal(handle.entry, `/shimo/doc/${handle.id}/index.html`);
    assert.deepEqual(handle.capabilities, { import: true, preview: true, edit: true, save: true, export: true });
    const meta = JSON.parse((await f.service.resource(handle.id, 'meta', peer, signal())).body);
    assert.equal(meta.v, 1); assert.equal(meta.routePrefix, '/shimo'); assert.equal(meta.stageUrl, `/shimo/doc/${handle.id}/stage/`);
    const image = format === 'xlsx' ? 'media/a.png' : format === 'docx' ? 'images/a.png' : 'media/a.png';
    assert.equal((await f.service.resource(handle.id, `resources/${image}`, peer, signal())).body.toString(), 'png');
    const [part, body] = format === 'docx' ? ['package', zip({ 'content.json': '{"doc":2,"nodeId":"n"}', 'image/b.png': 'b' })]
      : format === 'xlsx' ? ['content', Buffer.from('modoc-2')] : ['snapshot', Buffer.from('{"slides":[1]}')];
    assert.deepEqual(await f.service.stage(handle.id, part, body, { baseEditRevision: 0 }, peer, signal()), { editRevision: 1 });
    const content = (await f.service.resource(handle.id, 'content', peer, signal())).body.toString();
    assert.match(content, format === 'xlsx' ? /modoc-2/ : format === 'docx' ? /"doc":2/ : /\[1\]/);
    const saved = await f.service.save(scope, handle.id, { editRevision: 1 }, signal(), peer);
    assert.deepEqual(saved, { version: 'v2', editRevision: 1, path: `/authorized/a.${format}` });
    assert.equal(f.state.writes.length, 1);
    if (format === 'docx') assert.equal(f.converter.calls.exports.at(-1), '{"doc":2}', 'Writer runtime nodeId removed before export');
    const download = await f.service.exportTo(scope, handle.id, { editRevision: 1, target: 'download' }, signal(), peer);
    assert.match(download.url, new RegExp(`^/shimo/doc/${handle.id}/export/[a-f0-9]{48}$`));
    const route = download.url.slice(`/shimo/doc/${handle.id}/`.length);
    const file = await f.service.resource(handle.id, route, peer, signal());
    assert.match(file.headers['Content-Disposition'], new RegExp(`a\\.${format}`));
    await assert.rejects(f.service.resource(handle.id, route, peer, signal()), code('not-found'), 'download tokens are one-shot');
    const copy = await f.service.exportTo(scope, handle.id, { editRevision: 1, target: { path: `copy.${format}` } }, signal(), peer);
    assert.equal(copy.absolutePath, `/authorized/copy.${format}`);
    await assert.rejects(f.service.exportTo(scope, handle.id, { editRevision: 1, target: { path: `copy.${format}` } }, signal(), peer), code('exists'));
    await f.service.release(scope, handle.id, signal(), peer);
    assert.deepEqual(await taskDirs(f.tempRoot), []);
  });
}

test('rejects invalid requests before conversion', async t => {
  const f = await fixture(t);
  const peer = peerOf();
  await assert.rejects(f.service.open(scope, 'a.doc', 1, signal(), peer), code('unsupported-format'));
  await assert.rejects(f.service.open(scope, 'a.docx', -1, signal(), peer), code('invalid-revision'));
  await assert.rejects(f.service.open(scope, 'a.docx', 1.5, signal(), peer), code('invalid-revision'));
  await assert.rejects(f.service.open(scope, 'a.docx', 1, signal(), undefined), code('forbidden'));
  await assert.rejects(f.service.open({ sessionId: 'other', workspaceRoot: '/x' }, 'a.docx', 1, signal(), peer), code('forbidden'));
  await assert.rejects(f.service.open(scope, 'a.docx', 1, signal(), peer, { mode: 'write' }), code('unsupported-operation'));
  f.state.allowed = false;
  await assert.rejects(f.service.open(scope, 'a.docx', 1, signal(), peer), code('forbidden'));
  assert.equal(f.converter.calls.import, 0);
});

test('edit mode requires edit authorization and edit capability', async t => {
  const f = await fixture(t, { noWriter: true });
  await assert.rejects(f.service.open(scope, 'a.docx', 1, signal(), peerOf(), { mode: 'edit' }), code('unsupported-operation'));
  const view = await f.service.open(scope, 'a.docx', 1, signal(), peerOf());
  assert.equal(view.capabilities.save, false);
  assert.equal(view.capabilities.export, true);
  const g = await fixture(t);
  g.state.denyActions.add('edit');
  await assert.rejects(g.service.open(scope, 'a.docx', 1, signal(), peerOf(), { mode: 'edit' }), code('forbidden'));
});

test('cancellation at every open step leaves no task directory', async t => {
  const steps = ['acquire', 'authorize', 'stat', 'read', 'convert', 'restat'];
  for (const step of steps) {
    const f = await fixture(t);
    const controller = new AbortController();
    const at = name => { if (name === step) controller.abort(); };
    const authorize = f.service.authorizer.authorize;
    let authorizations = 0;
    f.service.authorizer = { authorize: async (...args) => { at(++authorizations === 1 ? 'authorize' : 'restat'); return authorize(...args); } };
    const stat = f.source.stat.bind(f.source), read = f.source.readBytes.bind(f.source);
    f.source.stat = async (...args) => { if (authorizations === 1) at('stat'); return stat(...args); };
    f.source.readBytes = async (...args) => { at('read'); return read(...args); };
    const imported = f.converter.import;
    f.converter.import = async (...args) => { at('convert'); return imported(...args); };
    if (step === 'acquire') controller.abort();
    await assert.rejects(f.service.open(scope, 'a.docx', 1, controller.signal, peerOf()), code('cancelled'), step);
    assert.deepEqual(await taskDirs(f.tempRoot), [], step);
    assert.equal(f.service.sessions.size, 0, step);
  }
});

test('timeouts, source changes and invalid converter output are rejected', async t => {
  const slow = fakeConverter({ import: (task, sgn) => new Promise((_, reject) => sgn.addEventListener('abort', () => reject(sgn.reason))) });
  const f = await fixture(t, { converter: slow, limits: { timeoutMs: { import: { docx: 20 } } } });
  await assert.rejects(f.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('timeout'));
  const g = await fixture(t);
  const imported = g.converter.import;
  g.converter.import = async (...args) => { const result = await imported(...args); g.state.version++; return result; };
  await assert.rejects(g.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('source-changed'));
  for (const result of [{ format: 'xlsx', content: 'content.json', assets: [] }, { format: 'docx', content: '../x', assets: [] }, { format: 'docx', content: 'missing', assets: [] }]) {
    const h = await fixture(t, { converter: fakeConverter({ import: async () => result }) });
    await assert.rejects(h.service.open(scope, 'a.docx', 1, signal(), peerOf()), error => ['invalid-output', 'invalid-path'].includes(error.code) || error.code === 'ENOENT');
    assert.deepEqual(await taskDirs(h.tempRoot), []);
  }
  const linked = await fixture(t, { converter: fakeConverter({ import: async ({ format, output }) => {
    await writeFile(join(output, 'content.json'), '{}'); await symlink('/etc/hosts', join(output, 'link'));
    return { format, content: 'content.json', assets: [] };
  } }) });
  await assert.rejects(linked.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('invalid-output'));
});

test('scratch may hold special files, output may not; both share the byte budget', async t => {
  const f = await fixture(t, { converter: fakeConverter({ import: async ({ format, output }) => {
    const scratch = join(output, '..', 'scratch');
    await symlink('/etc/hosts', join(scratch, 'socket-like'));
    await writeFile(join(output, 'content.json'), '{}');
    return { format, content: 'content.json', assets: [] };
  } }) });
  await f.service.open(scope, 'a.docx', 1, signal(), peerOf());
  const big = await fixture(t, { limits: { maxOutputBytes: 64 }, converter: fakeConverter({ import: async ({ format, output }, sgn) => {
    await writeFile(join(output, '..', 'scratch', 'cache'), Buffer.alloc(100));
    await writeFile(join(output, 'content.json'), '{}');
    await new Promise((resolve, reject) => { const timer = setTimeout(resolve, 300); sgn.addEventListener('abort', () => { clearTimeout(timer); reject(sgn.reason); }); });
    return { format, content: 'content.json', assets: [] };
  } }) });
  await assert.rejects(big.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('output-too-large'));
});

test('input, session, queue and retained-byte limits', async t => {
  const f = await fixture(t, { limits: { maxInputBytes: 4 } });
  await assert.rejects(f.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('input-too-large'));
  const g = await fixture(t, { limits: { maxLeases: 1 } });
  await g.service.open(scope, 'a.docx', 1, signal(), peerOf());
  await assert.rejects(g.service.open(scope, 'b.docx', 1, signal(), peerOf()), code('busy'));
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  const h = await fixture(t, { limits: { maxConcurrent: 1, maxQueued: 1 }, converter: fakeConverter({ import: async ({ format, output }) => { await gate; await writeFile(join(output, 'c.json'), '{}'); return { format, content: 'c.json', assets: [] }; } }) });
  const first = h.service.open(scope, 'a.docx', 1, signal(), peerOf());
  const second = h.service.open(scope, 'a.docx', 1, signal(), peerOf());
  await new Promise(resolve => setImmediate(resolve));
  await assert.rejects(h.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('busy'));
  release(); await first; await second;
  const k = await fixture(t, { limits: { maxRetainedBytes: 10, maxInputBytes: 8, maxOutputBytes: 8 } });
  await assert.rejects(k.service.open(scope, 'a.docx', 1, signal(), peerOf()), code('busy'));
});

test('sessions are bound to peer and scope, recheck the source and expire', async t => {
  let now = 1_000;
  const f = await fixture(t, { now: () => now });
  const peer = peerOf();
  const handle = await f.service.open(scope, 'a.docx', 1, signal(), peer);
  await assert.rejects(f.service.resource(handle.id, 'content', peerOf('other'), signal()), code('lease-unavailable'));
  await assert.rejects(f.service.release({ ...scope, sessionId: 'x' }, handle.id, signal(), peer), code('lease-unavailable'));
  assert.deepEqual(await f.service.renew(scope, handle.id, signal(), peer), { expiresInMs: 15 * 60_000 });
  f.state.version++;
  await assert.rejects(f.service.resource(handle.id, 'content', peer, signal()), code('source-changed'));
  await assert.rejects(f.service.renew(scope, handle.id, signal(), peer), code('source-changed'));
  f.state.version--;
  f.state.allowed = false;
  await assert.rejects(f.service.resource(handle.id, 'content', peer, signal()), code('forbidden'));
  f.state.allowed = true;
  now += 15 * 60_000;
  await assert.rejects(f.service.resource(handle.id, 'content', peer, signal()), code('lease-unavailable'));
  assert.deepEqual(await taskDirs(f.tempRoot), []);
});

test('peer disconnect closes its sessions', async t => {
  const f = await fixture(t);
  const controller = new AbortController();
  const peer = { id: 'p', principal: 'owner', signal: controller.signal };
  const handle = await f.service.open(scope, 'a.docx', 1, signal(), peer);
  controller.abort();
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(f.service.sessions.has(handle.id), false);
  assert.deepEqual(await taskDirs(f.tempRoot), []);
});

test('stage enforces edit mode, revisions, part whitelists and the stage limit', async t => {
  const f = await fixture(t, { limits: { maxStageBytes: 1024 } });
  const peer = peerOf();
  const view = await f.service.open(scope, 'a.docx', 1, signal(), peer);
  await assert.rejects(f.service.stage(view.id, 'package', Buffer.from('x'), { baseEditRevision: 0 }, peer, signal()), code('read-only'));
  const edit = await f.service.open(scope, 'a.docx', 1, signal(), peer, { mode: 'edit' });
  await assert.rejects(f.service.stage(edit.id, 'package', Buffer.alloc(2048), { baseEditRevision: 0 }, peer, signal()), code('input-too-large'));
  await assert.rejects(f.service.stage(edit.id, 'content', Buffer.from('x'), { baseEditRevision: 0 }, peer, signal()), code('not-found'));
  await assert.rejects(f.service.stage(edit.id, 'package', zip({ '../evil': 'x', 'content.json': '{}' }), { baseEditRevision: 0 }, peer, signal()), code('invalid-body'));
  await assert.rejects(f.service.stage(edit.id, 'package', zip({ 'other/x': 'x', 'content.json': '{}' }), { baseEditRevision: 0 }, peer, signal()), code('invalid-body'));
  await assert.rejects(f.service.stage(edit.id, 'package', zip({ 'image/a.png': 'x' }), { baseEditRevision: 0 }, peer, signal()), code('invalid-body'));
  await f.service.stage(edit.id, 'package', zip({ 'content.json': '{}' }), { baseEditRevision: 0 }, peer, signal());
  await assert.rejects(f.service.stage(edit.id, 'package', zip({ 'content.json': '{}' }), { baseEditRevision: 0 }, peer, signal()), code('edit-conflict'));
  const deck = await f.service.open(scope, 'a.pptx', 1, signal(), peer, { mode: 'edit' });
  await assert.rejects(f.service.stage(deck.id, 'asset/media/abc.png', Buffer.from('png'), { baseEditRevision: 0 }, peer, signal()), code('invalid-body'));
  await assert.rejects(f.service.stage(deck.id, 'asset/other/x.png', Buffer.from('png'), { baseEditRevision: 0 }, peer, signal()), code('invalid-path'));
  const { createHash } = await import('node:crypto');
  const digest = createHash('sha256').update('png').digest('hex');
  await f.service.stage(deck.id, `asset/media/${digest}.png`, Buffer.from('png'), { baseEditRevision: 0 }, peer, signal());
  const inline = { slides: [{ fill: { resourceUri: 'data:image/png;base64,' + Buffer.from('fill').toString('base64') } }] };
  await f.service.stage(deck.id, 'snapshot', Buffer.from(JSON.stringify(inline)), { baseEditRevision: 1 }, peer, signal());
  const staged = JSON.parse((await f.service.resource(deck.id, 'content', peer, signal())).body);
  assert.match(staged.slides[0].fill.resourceUri, /^mop-asset:\/media\/[a-f0-9]{64}\.png$/);
  const name = staged.slides[0].fill.resourceUri.split('/').at(-1);
  assert.equal((await f.service.resource(deck.id, `resources/media/${name}`, peer, signal())).body.toString(), 'fill');
});

test('save rejects stale edit revisions, oversized output and incomplete packages', async t => {
  const f = await fixture(t, { limits: { maxInputBytes: 400 } });
  const peer = peerOf();
  const edit = await f.service.open(scope, 'a.xlsx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(edit.id, 'content', Buffer.from('m'), { baseEditRevision: 0 }, peer, signal());
  await assert.rejects(f.service.save(scope, edit.id, { editRevision: 0 }, signal(), peer), code('edit-conflict'));
  await f.service.stage(edit.id, 'content', Buffer.from('m'.repeat(600)), { baseEditRevision: 1 }, peer, signal());
  await assert.rejects(f.service.save(scope, edit.id, { editRevision: 2 }, signal(), peer), code('output-too-large'));
  assert.deepEqual(await f.service.listDrafts(scope, signal(), peer), [], 'an over-size save keeps the session instead of a draft');
  const g = await fixture(t, { converter: fakeConverter({ export: async ({ output }) => writeFile(output, zip({ '[Content_Types].xml': 'x' })) }) });
  const doc = await g.service.open(scope, 'a.docx', 1, signal(), peer, { mode: 'edit' });
  await assert.rejects(g.service.save(scope, doc.id, { editRevision: 0 }, signal(), peer), code('invalid-output'));
  assert.equal(g.state.writes.length, 0);
});

test('save conflicts never overwrite and keep an Office draft any peer of the principal can recover', async t => {
  const f = await fixture(t);
  const peer = peerOf('tab-1', 'alice');
  const edit = await f.service.open(scope, 'a.docx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(edit.id, 'package', zip({ 'content.json': '{"mine":true}' }), { baseEditRevision: 0 }, peer, signal());
  f.state.version = 7;
  await assert.rejects(f.service.save(scope, edit.id, { editRevision: 1 }, signal(), peer), error => error.code === 'source-changed' && /^[a-f0-9]{32}$/.test(error.details.draftId));
  assert.equal(f.state.writes.length, 0);
  const reconnected = peerOf('tab-2', 'alice');
  const drafts = await f.service.listDrafts(scope, signal(), reconnected);
  assert.equal(drafts.length, 1);
  assert.equal(drafts[0].kind, 'office');
  assert.deepEqual(await f.service.listDrafts(scope, signal(), peerOf('tab-3', 'mallory')), []);
  const saved = await f.service.exportDraft(scope, drafts[0].id, { target: { path: 'recovered.docx' } }, signal(), reconnected);
  assert.equal(saved.absolutePath, '/authorized/recovered.docx');
  const download = await f.service.exportDraft(scope, drafts[0].id, { target: 'download' }, signal(), reconnected);
  assert.match(download.url, new RegExp(`^/shimo/draft/${drafts[0].id}/export/`));
  const body = await f.service.draftResource(drafts[0].id, download.url.split(`/draft/${drafts[0].id}/`)[1], reconnected, signal());
  assert.ok(body.body.length > 0);
  await f.service.discardDraft(scope, drafts[0].id, signal(), reconnected);
  assert.deepEqual(await f.service.listDrafts(scope, signal(), reconnected), []);
});

test('a save outside the writable roots keeps the edits in the session until the user agrees', async t => {
  const f = await fixture(t);
  const peer = peerOf('tab', 'alice');
  f.state.outside.add('b.pptx');
  const deck = await f.service.open(scope, 'b.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(deck.id, 'snapshot', Buffer.from('{"slides":[2]}'), { baseEditRevision: 0 }, peer, signal());
  await assert.rejects(f.service.save(scope, deck.id, { editRevision: 1 }, signal(), peer), error => error.code === 'sandbox-denied' && error.details.outside === true && !error.details.draftId);
  assert.equal(f.state.writes.length, 0);
  assert.deepEqual(await f.service.listDrafts(scope, signal(), peer), [], 'nothing to recover: the session still holds the edits');
  const saved = await f.service.save(scope, deck.id, { editRevision: 1, approved: true }, signal(), peer);
  assert.equal(saved.version, 'v2');
  assert.equal(f.state.writes.at(-1).path, 'b.pptx');
});

test('a released tab saves its edits where it may; otherwise they stay parked for the next open', async t => {
  const f = await fixture(t);
  const peer = peerOf('tab', 'owner');
  const mine = await f.service.open(scope, 'a.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(mine.id, 'snapshot', Buffer.from('{"slides":[9]}'), { baseEditRevision: 0 }, peer, signal());
  await f.service.release(scope, mine.id, signal(), peer);
  assert.deepEqual(f.state.writes.map(w => w.path), ['a.pptx'], 'a writable file is saved when its tab goes away');
  assert.equal(f.service.sessions.has(mine.id), false);

  f.state.version = 1;
  f.state.outside.add('b.pptx');
  const outside = await f.service.open(scope, 'b.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(outside.id, 'snapshot', Buffer.from('{"slides":[7]}'), { baseEditRevision: 0 }, peer, signal());
  await f.service.release(scope, outside.id, signal(), peer);
  assert.equal(f.state.writes.length, 1, 'outside the writable roots nothing is written without consent');
  await assert.rejects(f.service.resource(outside.id, 'content', peer, signal()), code('lease-unavailable'), 'a parked session answers no old lease');
  const imports = f.converter.calls.import;
  const back = await f.service.open(scope, 'b.pptx', 2, signal(), peerOf('tab-2', 'owner'), { mode: 'view' });
  assert.equal(back.id, outside.id, 'the same file in the same Session resumes the parked session');
  assert.deepEqual([back.mode, back.unsaved, back.editRevision, back.revision, back.saveNeedsApproval], ['edit', true, 1, 2, true], 'the panel marks it as waiting for a confirmed write-back');
  assert.equal(f.converter.calls.import, imports, 'no second conversion');
  assert.equal((await f.service.resource(back.id, 'content', peerOf('tab-2', 'owner'), signal())).body.toString(), '{"slides":[7]}');

  const stranger = await f.service.open(scope, 'b.pptx', 2, signal(), peerOf('x', 'mallory'), { mode: 'edit' });
  assert.notEqual(stranger.id, outside.id, 'another principal never picks up these edits');
  await f.service.release(scope, stranger.id, signal(), peerOf('x', 'mallory'));

  await f.service.release(scope, back.id, signal(), peerOf('tab-2', 'owner'));
  f.state.version = 5;
  const changed = await f.service.open(scope, 'b.pptx', 3, signal(), peerOf('tab-3', 'owner'), { mode: 'edit' });
  assert.notEqual(changed.id, outside.id, 'a file changed meanwhile opens fresh');
  assert.equal(changed.unsaved, false);
  const drafts = await f.service.listDrafts(scope, signal(), peerOf('tab-3', 'owner'));
  assert.deepEqual(drafts.map(d => [d.kind, d.format]), [['snapshot', 'pptx']], 'the parked edits became a draft');
});

test('shutdown drafts every unsaved session first, then saves the ones it may', async t => {
  const f = await fixture(t);
  const peer = peerOf('tab', 'owner');
  const mine = await f.service.open(scope, 'a.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(mine.id, 'snapshot', Buffer.from('{"slides":[9]}'), { baseEditRevision: 0 }, peer, signal());
  f.state.outside.add('c.xlsx');
  const outside = await f.service.open(scope, 'c.xlsx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(outside.id, 'content', Buffer.from('dirty'), { baseEditRevision: 0 }, peer, signal());
  await f.service.open(scope, 'd.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.dispose();
  assert.deepEqual(f.state.writes.map(w => w.path), ['a.pptx']);
  const { SuiteService } = await import('../src/index.mjs');
  const restarted = new SuiteService({ source: f.source, writer: f.writer, authorizer: f.service.authorizer, converter: f.converter, runtimeRoot: f.service.runtimeRoot, tempRoot: f.tempRoot, dataRoot: f.dataRoot });
  t.after(() => restarted.dispose());
  const drafts = await restarted.listDrafts(scope, signal(), peerOf('new-tab', 'owner'));
  assert.deepEqual(drafts.map(d => [d.kind, d.format]), [['snapshot', 'xlsx']], 'the saved one dropped its draft, the clean one never had one');
  await restarted.exportDraft(scope, drafts[0].id, { target: { path: 'copy.xlsx' } }, signal(), peerOf('new-tab', 'owner'));
  assert.equal(f.converter.calls.exports.at(-1), 'dirty', 'snapshot drafts convert lazily on export');
});

test('only the owner autosaves: a principal denied save keeps its edits parked, then as a draft', async t => {
  let now = 0;
  const f = await fixture(t, { now: () => now, limits: { parkTtlMs: 5 } });
  const guest = peerOf('tab', 'guest');
  f.state.denyActions.add('save');
  const deck = await f.service.open(scope, 'a.pptx', 1, signal(), guest, { mode: 'edit' });
  await f.service.stage(deck.id, 'snapshot', Buffer.from('{"slides":[3]}'), { baseEditRevision: 0 }, guest, signal());
  await f.service.release(scope, deck.id, signal(), guest);
  assert.equal(f.state.writes.length, 0, 'closing the tab writes nothing for a non-owner');
  now = 10;
  await f.service.expire();
  assert.equal(f.state.writes.length, 0);
  assert.deepEqual((await f.service.listDrafts(scope, signal(), guest)).map(d => d.kind), ['snapshot']);

  const other = await f.service.open(scope, 'b.pptx', 1, signal(), guest, { mode: 'edit' });
  await f.service.stage(other.id, 'snapshot', Buffer.from('{"slides":[4]}'), { baseEditRevision: 0 }, guest, signal());
  await f.service.dispose();
  assert.equal(f.state.writes.length, 0, 'shutdown does not write for a non-owner either');
});

test('a draft evicted after a failed save is kept again when the session closes', async t => {
  const f = await fixture(t);
  const peer = peerOf('tab');
  const deck = await f.service.open(scope, 'a.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(deck.id, 'snapshot', Buffer.from('{"slides":[6]}'), { baseEditRevision: 0 }, peer, signal());
  f.state.version = 9;
  const draftId = await f.service.save(scope, deck.id, { editRevision: 1 }, signal(), peer).then(() => undefined, error => error.details?.draftId);
  assert.ok(draftId, 'the failed save kept a draft');
  await f.service.drafts.remove(draftId); // evicted by count, bytes or age
  await f.service.closeSession(f.service.sessions.get(deck.id), { keepDraft: true });
  assert.equal((await f.service.listDrafts(scope, signal(), peer)).length, 1, 'the unsaved edits are not lost');
});

test('a failed shutdown save reuses the draft it kept first', async t => {
  const f = await fixture(t);
  const peer = peerOf('tab');
  const deck = await f.service.open(scope, 'a.pptx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(deck.id, 'snapshot', Buffer.from('{"slides":[5]}'), { baseEditRevision: 0 }, peer, signal());
  f.state.version = 9; // the file moved on: the write is refused as a version conflict
  await f.service.dispose();
  assert.equal(f.state.writes.length, 0);
  const { SuiteService } = await import('../src/index.mjs');
  const restarted = new SuiteService({ source: f.source, writer: f.writer, authorizer: f.service.authorizer, converter: f.converter, runtimeRoot: f.service.runtimeRoot, tempRoot: f.tempRoot, dataRoot: f.dataRoot });
  t.after(() => restarted.dispose());
  assert.equal((await restarted.listDrafts(scope, signal(), peerOf('new-tab'))).length, 1, 'one draft, not one per failed step');
});
test('drafts are evicted by count, bytes and age', async t => {
  let now = 0;
  const f = await fixture(t, { now: () => now, limits: { maxDrafts: 2, maxDraftBytes: 1_000_000, draftTtlMs: 1000, parkTtlMs: 5 } });
  const peer = peerOf('p', 'alice');
  // Parked edits that nobody reopens become drafts when their keep time runs out.
  const draftOf = async (service, state, name, body) => {
    state.outside.add(`${name}.xlsx`);
    const edit = await service.open(scope, `${name}.xlsx`, 1, signal(), peer, { mode: 'edit' });
    await service.stage(edit.id, 'content', Buffer.from(body), { baseEditRevision: 0 }, peer, signal());
    await service.release(scope, edit.id, signal(), peer);
  };
  for (const name of ['a', 'b', 'c']) {
    now += 10;
    await draftOf(f.service, f.state, name, name);
    now += 10;
    await f.service.expire();
  }
  assert.deepEqual((await f.service.listDrafts(scope, signal(), peer)).map(d => d.path), ['/authorized/c.xlsx', '/authorized/b.xlsx']);
  now += 2000;
  assert.deepEqual(await f.service.listDrafts(scope, signal(), peer), []);
  let later = 0;
  const g = await fixture(t, { now: () => later, limits: { maxDraftBytes: 30, parkTtlMs: 5 } });
  for (const name of ['a', 'b']) {
    later += 10;
    await draftOf(g.service, g.state, name, name.repeat(20));
    later += 10;
    await g.service.expire();
  }
  assert.equal((await g.service.listDrafts(scope, signal(), peer)).length, 1);
});

test('rebind keeps a saved session across the revision bump and diverges a dirty one', async t => {
  const f = await fixture(t);
  const peer = peerOf();
  const edit = await f.service.open(scope, 'a.xlsx', 1, signal(), peer, { mode: 'edit' });
  await f.service.stage(edit.id, 'content', Buffer.from('x'), { baseEditRevision: 0 }, peer, signal());
  await f.service.save(scope, edit.id, { editRevision: 1 }, signal(), peer);
  assert.deepEqual(await f.service.rebind(scope, edit.id, 2, {}, signal(), peer), { revision: 2, state: 'active', version: 'v2' });
  assert.equal(JSON.parse((await f.service.resource(edit.id, 'meta', peer, signal())).body).revision, 2);
  await assert.rejects(f.service.rebind(scope, edit.id, 3, { diverged: true }, signal(), peer), code('unsupported-operation'), 'clean sessions reopen instead');
  await f.service.stage(edit.id, 'content', Buffer.from('y'), { baseEditRevision: 1 }, peer, signal());
  f.state.version = 9;
  await assert.rejects(f.service.rebind(scope, edit.id, 3, {}, signal(), peer), code('source-changed'));
  assert.equal((await f.service.rebind(scope, edit.id, 3, { diverged: true }, signal(), peer)).state, 'diverged');
  assert.ok(await f.service.renew(scope, edit.id, signal(), peer));
  assert.equal((await f.service.resource(edit.id, 'content', peer, signal())).body.toString(), 'y');
  assert.equal((await f.service.resource(edit.id, 'resources/media/a.png', peer, signal())).body.toString(), 'png');
  await assert.rejects(f.service.save(scope, edit.id, { editRevision: 2 }, signal(), peer), error => error.code === 'source-changed' && !!error.details.draftId);
  assert.equal(f.state.writes.length, 1, 'diverged sessions never overwrite');
  assert.ok((await f.service.exportTo(scope, edit.id, { editRevision: 2, target: 'download' }, signal(), peer)).url);
});

test('start removes orphaned task directories from a crashed host', async t => {
  const f = await fixture(t);
  await mkdir(f.tempRoot, { recursive: true });
  await mkdir(join(f.tempRoot, 'officesdk-orphan'));
  await mkdir(join(f.tempRoot, 'unrelated'));
  await f.service.open(scope, 'a.docx', 1, signal(), peerOf());
  const names = await readdir(f.tempRoot);
  assert.ok(!names.includes('officesdk-orphan'));
  assert.ok(names.includes('unrelated'));
});

test('dispose cancels running work and leaves no task directories', async t => {
  const f = await fixture(t, { converter: fakeConverter({ import: (task, sgn) => new Promise((_, reject) => sgn.addEventListener('abort', () => reject(sgn.reason))) }) });
  const pending = f.service.open(scope, 'a.docx', 1, signal(), peerOf());
  await new Promise(resolve => setTimeout(resolve, 20));
  await f.service.dispose();
  await assert.rejects(pending, code('cancelled'));
  assert.deepEqual(await taskDirs(f.tempRoot), []);
});

test('binary presentation snapshots are stored byte for byte', async t => {
  const f = await fixture(t, { converter: fakeConverter({ import: async ({ format, output }) => {
    await writeFile(join(output, 'deck.bin'), Buffer.from([0x6d, 0x6f, 0x70, 0x30, 0xff]));
    return { format, content: 'deck.bin', assets: [] };
  } }) });
  const peer = peerOf();
  const edit = await f.service.open(scope, 'a.pptx', 1, signal(), peer, { mode: 'edit' });
  const body = Buffer.from([0x6d, 0x6f, 0x70, 0x30, 0x80, 0xfe, 0x00, 0x64, 0x61, 0x74, 0x61, 0x3a]);
  await f.service.stage(edit.id, 'snapshot', body, { baseEditRevision: 0 }, peer, signal());
  const stored = await f.service.resource(edit.id, 'content', peer, signal());
  assert.deepEqual(stored.body, body);
  assert.equal(stored.type, 'application/octet-stream');
});

test('resource reads over the reader limit wait for a slot instead of failing', async t => {
  const f = await fixture(t, { limits: { maxReaders: 2 } });
  const peer = peerOf();
  const deck = await f.service.open(scope, 'a.pptx', 1, signal(), peer);
  const bodies = await Promise.all(Array.from({ length: 20 }, () => f.service.resource(deck.id, 'resources/media/a.png', peer, signal())));
  assert.deepEqual(new Set(bodies.map(({ body }) => body.toString())), new Set(['png']));
  const stopped = new AbortController();
  const held = Array.from({ length: 3 }, () => f.service.resource(deck.id, 'resources/media/a.png', peer, stopped.signal));
  stopped.abort();
  for (const read of held) await read.catch(() => {});
  assert.equal(f.service.readers.active, 0, 'aborted waiters release their place');
});
