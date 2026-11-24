import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdir, mkdtemp, readFile, realpath, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { homedir, tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { Context } from '@deepseek-ai/cordis';
import { withFileLock, writeFileAtomic } from '@deepseek-ai/dsh-atomic-write';
import { writableRoots } from '@deepseek-ai/dsh-sandbox';
import { SandboxedFileSystem } from '@deepseek-ai/dsh-fs-sandbox';
import ShimoPreview from '../src/host/index.mjs';
import { createWriter } from '../src/host/writer.mjs';
import { KitProcess } from '../../kit-host/src/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';

const FAKE_KIT = fileURLToPath(new URL('../../kit-host/test/fake-kit.mjs', import.meta.url));
const signal = () => new AbortController().signal;
const remoteCode = expected => error => { assert.equal(error.code, expected, error.message); return true; };

/** A Session over real files: `ws` is the workspace, `outside` is not; the policy is workspace-write unless `mode` says otherwise. */
async function load(t, { mode = 'workspace-write', kitProcess, environment: overrides = {} } = {}) {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-kit-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const ws = join(root, 'ws'), outside = join(root, 'outside'), data = join(root, 'data');
  for (const dir of [ws, outside, data, join(root, 'runtime')]) await mkdir(dir, { recursive: true });
  await writeFile(join(root, 'runtime', 'preview.html'), '');
  for (const name of ['deck.pptx', 'changed.pptx']) await writeFile(join(ws, name), 'pptx');
  await writeFile(join(ws, 'a.docx'), 'docx');
  await writeFile(join(ws, 'book.xlsx'), 'xlsx');
  await writeFile(join(ws, 'mismatch.docx'), 'docx');
  const versions = new Map();
  const workspaceFiles = {
    async stat(scope, path) { const absolutePath = path.startsWith('/') ? path : join(ws, path); return { absolutePath, version: versions.get(absolutePath) ?? 'v1', bytes: (await readFile(absolutePath)).length }; },
    async readBytes(scope, path) { const absolutePath = path.startsWith('/') ? path : join(ws, path); const data = await readFile(absolutePath); return { absolutePath, version: 'v1', data, eof: true }; },
  };
  const policyMode = { mode };
  const policy = { resolve: ({ session } = {}) => ({ mode: session ? policyMode.mode : 'danger-full-access', workspaceRoot: ws }) };
  const services = { sandboxPolicy: policy, sessions: { get: id => ({ id, header: { cwd: ws } }) }, fs: Object.create(SandboxedFileSystem.prototype) };
  const ctxLike = { get: name => services[name] };
  const writer = createWriter({ ctx: ctxLike, workspaceFiles, dataRoot: data, writableRoots, atomic: { withFileLock, writeFileAtomic } });
  // `kitProcess: null` leaves the Host to find Kit itself (environment.kitRoot / host).
  const kit = kitProcess === null ? undefined : kitProcess ?? new KitProcess({ binary: process.execPath, work: join(data, 'kit'), restartMs: [50], killGraceMs: 300, spawn: (bin, args, options) => spawn(process.execPath, [FAKE_KIT, ...args], options) });
  const app = new Context();
  const provided = { workspaceFiles, webServer: { register: () => () => {} }, connection: { admit: () => ({ rejection: 403 }) },
    typert: { lookups: { get: () => ({ resolve: async id => ({ sessionId: id, workspaceRoot: ws }) }) } } };
  for (const [name, value] of Object.entries(provided)) { app.provide(name); app.set(name, value); }
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: data, writer, kitProcess: kit, ...overrides };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  await app.plugin(Plugin, {});
  await new Promise(resolve => setTimeout(resolve, 20));
  const service = app.get('shimoPreview');
  t.after(() => service.kitDocs?.dispose());
  const peer = { id: 'tab', principal: { kind: 'owner' }, signal: signal() };
  const remote = service.peerOf({ id: peer.id, principal: peer.principal, signal: peer.signal });
  service.remotePeer = () => remote;
  const scope = { sessionId: 's1', workspaceRoot: ws };
  const commands = async () => (await readFile(join(data, 'kit', 'commands.log'), 'utf8').catch(() => '')).trim().split('\n').filter(Boolean).map(line => JSON.parse(line));
  return { service, scope, ws, outside, data, kit, commands, writer, versions, policyMode };
}

test('writer.accessFor answers for a real host path the way the Session policy would', async t => {
  const { writer, scope, ws } = await load(t);
  assert.equal(await writer.accessFor(scope, join(ws, 'deck.pptx')), 'allowed');
  // The temp directory is a writable root of workspace-write Sessions; the home directory is not.
  assert.equal(await writer.accessFor(scope, join(homedir(), `officesdk-absent-${process.pid}`, 'far.pptx')), 'outside');
  assert.equal(await writer.accessFor(scope, 'deck.pptx'), 'denied', 'only absolute host paths');
  const readOnly = await load(t, { mode: 'read-only' });
  assert.equal(await readOnly.writer.accessFor(readOnly.scope, join(readOnly.ws, 'deck.pptx')), 'denied');
});

test('access follows the Session\'s permission while a Kit tab is open (DOCX, XLSX, PPTX); other documents and views say nothing', async t => {
  const { service, scope, policyMode } = await load(t);
  const views = await Promise.all(['a.docx', 'book.xlsx', 'deck.pptx'].map(name => service.open(scope, name, 1, 'edit', signal())));
  for (const view of views) assert.deepEqual([view.access, await service.access(scope, view.id, signal())], ['allowed', { access: 'allowed' }], view.format);
  policyMode.mode = 'read-only';
  for (const view of views) {
    assert.deepEqual(await service.access(scope, view.id, signal()), { access: 'denied' }, `${view.format}: read only now`);
    await assert.rejects(service.save(scope, view.id, 1, false, false, signal()), remoteCode('officesdk/sandbox-denied'), `${view.format}: a save is refused`);
  }
  policyMode.mode = 'workspace-write';
  for (const view of views) assert.deepEqual(await service.access(scope, view.id, signal()), { access: 'allowed' }, `${view.format}: writable again`);
  const read = await service.open(scope, 'a.docx', 1, 'view', signal());
  assert.deepEqual(await service.access(scope, read.id, signal()), {}, 'a view-mode tab has nothing to write');
  assert.deepEqual(await service.access(scope, 'not-a-kit-view', signal()), {});
});

test('DOCX, XLSX and PPTX open through Kit as leases with the editor page; `preview` asks for Core\'s read-only preview', async t => {
  const { service, scope, ws, commands } = await load(t);
  const deck = await service.open(scope, 'deck.pptx', 3, 'edit', signal());
  assert.deepEqual([deck.backend, deck.format, deck.mode, deck.revision, deck.capabilities.save, deck.platform], ['kit', 'pptx', 'edit', 3, true, process.platform]);
  assert.equal(deck.entry, `/shimo/kit/static/page/index.html?view=${deck.id}`);
  assert.deepEqual((await commands()).map(command => [command.op, command.path]), [['open', join(ws, 'deck.pptx')]]);
  const doc = await service.open(scope, 'a.docx', 1, 'edit', signal());
  const book = await service.open(scope, 'book.xlsx', 1, 'edit', signal());
  assert.deepEqual([doc.backend, doc.format, doc.mode, book.backend, book.format, book.mode], ['kit', 'docx', 'edit', 'kit', 'xlsx', 'edit']);
  const preview = await service.open(scope, 'a.docx', 1, 'preview', signal());
  assert.deepEqual([preview.backend, preview.mode], [undefined, 'view']);
  assert.equal((await commands()).filter(command => command.op === 'open').length, 3, 'the preview never reaches Kit');
  // No converter ships with the plugin: a file Kit reads as another suite is Kit's own error, not a Core preview.
  await assert.rejects(service.open(scope, 'mismatch.docx', 1, 'edit', signal()), error => error.code === 'officesdk/unsupported-format');
  assert.equal((await commands()).filter(command => command.op === 'discard').length, 0);
  assert.deepEqual(await service.renew(scope, deck.id, signal()), { expiresInMs: service.kitDocs.viewTtlMs });
  assert.deepEqual(await service.rebind(scope, deck.id, 4, false, signal()), { revision: 4, state: 'active', version: 'v1' }, 'the file is what the view opened');
});

test('Kit is found by this platform\'s executable name; without a Kit, opening says 本平台暂不支持', async t => {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-kit-platform-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const kitRoot = async (name, files) => { const dir = join(root, name); await mkdir(dir, { recursive: true }); for (const file of files) await writeFile(join(dir, file), ''); return dir; };
  const windowsKit = await kitRoot('win', ['offline-kit-runtime.exe', 'compose-worker.exe', 'modoc.js']);
  const macKit = await kitRoot('mac', ['offline-kit-runtime', 'compose-worker', 'modoc.js']);
  const win32 = { platform: 'win32', arch: 'x64' };

  // A Windows package: its .exe Kit is launched.
  const windows = await load(t, { kitProcess: null, environment: { host: win32, kitRoot: windowsKit, converter: undefined } });
  assert.equal(windows.service.kit.binary, join(windowsKit, 'offline-kit-runtime.exe'));
  assert.equal(windows.service.platform, 'win32', 'leases carry it for the Client\'s path rules');
  // A macOS arm64 package: the extension-less Kit.
  const mac = await load(t, { kitProcess: null, environment: { host: { platform: 'darwin', arch: 'arm64' }, kitRoot: macKit } });
  assert.equal(mac.service.kit.binary, join(macKit, 'offline-kit-runtime'));

  const unsupported = (reason, text) => error => {
    assert.equal(error.code, 'officesdk/unsupported-operation', error.message);
    assert.match(error.message, /^本平台暂不支持/);
    assert.match(error.message, text);
    return true;
  };
  // The macOS package installed on Windows: its Kit is not launched and the message says which package is needed.
  const wrong = await load(t, { kitProcess: null, environment: { host: win32, kitRoot: macKit, converter: undefined } });
  assert.equal(wrong.service.kitDocs, undefined);
  for (const mode of ['edit', 'view', 'preview']) await assert.rejects(wrong.service.open(wrong.scope, 'a.docx', 0, mode, signal()), unsupported('mismatch', /其他平台.*win32-x64/));
  // No Kit at all, and a platform no Kit is built for (an Intel Mac): the same plain answer, never a generic unavailable.
  const missing = await load(t, { kitProcess: null, environment: { host: win32, kitRoot: join(root, 'none'), converter: undefined } });
  await assert.rejects(missing.service.open(missing.scope, 'book.xlsx', 0, 'edit', signal()), unsupported('missing', /offline-kit-runtime\.exe/));
  const intel = await load(t, { kitProcess: null, environment: { host: { platform: 'darwin', arch: 'x64' }, kitRoot: macKit, converter: undefined } });
  await assert.rejects(intel.service.open(intel.scope, 'deck.pptx', 0, 'edit', signal()), unsupported('platform', /darwin-x64/));
});

test('save, overwrite, download and close go to Kit with this view\'s real path; failures keep shared codes', async t => {
  const { service, scope, ws, commands } = await load(t);
  const view = await service.open(scope, 'changed.pptx', 1, 'edit', signal());
  await service.kitDocs.change(view.id, service.remotePeer(), { seq: 0, format: 'mop-v1', body: 'edit' });
  await assert.rejects(service.save(scope, view.id, 0, false, false, signal()), remoteCode('officesdk/source-changed'));
  const saved = await service.save(scope, view.id, 0, false, true, signal());
  assert.deepEqual([saved.saved, saved.path], [true, join(ws, 'changed.pptx')]);
  const saves = (await commands()).filter(command => command.op === 'save');
  assert.deepEqual(saves.map(command => [command.target, command.overwrite ?? false]), [[join(ws, 'changed.pptx'), false], [join(ws, 'changed.pptx'), true]]);
  const download = await service.exportTo(scope, view.id, 0, 'download', signal());
  assert.match(download.url, new RegExp(`^/shimo/kit/v/${view.id}/download/[0-9a-f-]{36}$`));
  assert.deepEqual(await service.resolveDiskChange(scope, view.id, false, signal()), { kept: true });
  assert.deepEqual(await service.release(scope, view.id, true, signal()), { closed: true, saved: false });
  await assert.rejects(service.renew(scope, view.id, signal()), remoteCode('officesdk/lease-unavailable'));
});

test('a Kit restart surfaces as edit-conflict with rebuild; a Kit that cannot start says so (no Core fallback)', async t => {
  const { service, scope, kit } = await load(t);
  const view = await service.open(scope, 'deck.pptx', 1, 'edit', signal());
  const before = kit.generation;
  (await kit.ready()).request('crash');
  for (let i = 0; i < 100 && kit.generation === before; i++) await new Promise(resolve => setTimeout(resolve, 20));
  await assert.rejects(service.renew(scope, view.id, signal()), error => error.code === 'officesdk/edit-conflict' && error.details?.rebuild === true);

  const broken = Object.assign(new KitProcess({ binary: '/nonexistent/offline-kit-runtime', work: join(tmpdir(), `kit-broken-${process.pid}`) }), {});
  broken.ready = async () => { throw new Error('spawn ENOENT'); };
  const failing = await load(t, { kitProcess: broken });
  for (const name of ['deck.pptx', 'a.docx', 'book.xlsx']) {
    await assert.rejects(failing.service.open(failing.scope, name, 1, 'edit', signal()), error => error.code === 'officesdk/unavailable', `${name}: Kit's own failure is reported`);
  }
});

/** Plays the Client: takes the next AI command for the Session and answers it. */
const client = (service, answer) => (async () => { const command = await service.agents.next('s1'); service.agents.reply('s1', command.requestId, await answer(command)); return command; })();

test('AI on a Kit view: saving waits while an edit runs; an AI save takes the page\'s edits, then writes the real path', async t => {
  const { service, scope, ws, commands } = await load(t);
  const view = await service.open(scope, 'deck.pptx', 1, 'edit', signal());
  const editing = service.aiEdit('s1', { path: 'deck.pptx', ops: [{ op: 'addSlide' }], summary: 'add a slide' }, signal());
  const command = await service.agents.next('s1');
  assert.equal(command.kind, 'edit');
  await assert.rejects(service.save(scope, view.id, 0, false, false, signal()), remoteCode('officesdk/busy'), 'no save while the AI writes');
  await service.kitDocs.change(view.id, service.remotePeer(), { seq: 0, format: 'mop-v1', body: 'ai change' });
  service.agents.reply('s1', command.requestId, { ok: true, result: { applied: 1, editRevision: 1, undoable: command.editId } });
  assert.equal((await editing).unsaved, true);
  assert.equal(service.kitDocs.views.get(view.id).aiActive, 0, 'the AI is done writing');

  const saving = service.aiSave('s1', 'deck.pptx', { signal: signal() });
  const flush = await client(service, () => ({ ok: true, result: { editRevision: 1 } }));
  assert.deepEqual([flush.kind, flush.path], ['flush', 'deck.pptx']);
  assert.deepEqual(await saving, { saved: true, path: join(ws, 'deck.pptx') });
  assert.equal((await commands()).filter(entry => entry.op === 'save').at(-1).target, join(ws, 'deck.pptx'));
  const notice = await service.agents.next('s1');
  assert.deepEqual([notice.kind, notice.path, notice.editRevision], ['saved', 'deck.pptx', 1]);

  const unchanged = service.aiSave('s1', 'deck.pptx', { signal: signal() });
  await client(service, () => ({ ok: true, result: { editRevision: 1 } }));
  assert.deepEqual(await unchanged, { saved: false, unchanged: true, path: join(ws, 'deck.pptx') });
});

test('a tab falling back to the preview releases its Kit view without saving its unsaved edits', async t => {
  const { service, scope, commands } = await load(t);
  const view = await service.open(scope, 'deck.pptx', 1, 'edit', signal());
  await service.kitDocs.change(view.id, service.remotePeer(), { seq: 0, format: 'mop-v1', body: 'edit' });
  assert.deepEqual(await service.release(scope, view.id, false, signal()), { closed: true, saved: false });
  assert.equal((await commands()).filter(command => command.op === 'save').length, 0);
});

test('rebinding a Kit view checks the file: another program\'s write is source-changed, the view\'s own save is not', async t => {
  const { service, scope, ws, versions } = await load(t);
  const view = await service.open(scope, 'deck.pptx', 1, 'edit', signal());
  await service.kitDocs.change(view.id, service.remotePeer(), { seq: 0, format: 'mop-v1', body: 'edit' });
  versions.set(join(ws, 'deck.pptx'), 'v2');
  const saved = await service.save(scope, view.id, 1, false, false, signal());
  assert.equal(saved.version, 'v2');
  assert.deepEqual(await service.rebind(scope, view.id, 2, false, signal()), { revision: 2, state: 'active', version: 'v2' });
  versions.set(join(ws, 'deck.pptx'), 'v3');
  await assert.rejects(service.rebind(scope, view.id, 3, false, signal()), remoteCode('officesdk/source-changed'));
});
