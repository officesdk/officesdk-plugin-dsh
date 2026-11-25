import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, readdir, realpath, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { RemoteError } from '@deepseek-ai/dsh-typert-protocol';
import { withFileLock, writeFileAtomic } from '@deepseek-ai/dsh-atomic-write';
import { writableRoots } from '@deepseek-ai/dsh-sandbox';
import { SandboxedFileSystem } from '@deepseek-ai/dsh-fs-sandbox';
import { LocalFileSystem } from '@deepseek-ai/dsh-fs-local';
import { SuiteError } from '../../core/src/index.mjs';
import { createAuthorizer } from '../src/host/authorizer.mjs';
import { fromWorkspaceFile, principalKey, toRemote } from '../src/host/errors.mjs';
import { canonicalTarget, createWriter, isLocalFileSystem } from '../src/host/writer.mjs';
import ShimoPreview, { sameOrigin } from '../src/host/index.mjs';

const signal = () => new AbortController().signal;
const code = expected => error => { assert.equal(error.code, expected, error.message); return true; };

async function workspace(t) {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-adapter-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const ws = join(root, 'ws'), outside = join(root, 'outside'), data = join(root, 'data');
  await mkdir(ws); await mkdir(outside); await mkdir(data);
  await writeFile(join(ws, 'a.docx'), 'old'); await writeFile(join(outside, 'b.docx'), 'old');
  await symlink(outside, join(ws, 'escape'));
  return { root, ws, outside, data };
}

/** WorkspaceFiles-shaped stat over the real disk, with fs-local's dev:ino:size:mtime:ctime version. */
const workspaceFiles = ws => ({
  async stat(scope, path) {
    const absolute = path.startsWith('/') ? path : join(ws, path);
    const info = await stat(absolute, { bigint: true }).catch(() => { throw Object.assign(new Error('missing'), { code: 'workspace-file/not-found' }); });
    return { absolutePath: absolute, version: `${info.dev}:${info.ino}:${info.size}:${info.mtimeNs}:${info.ctimeNs}`, bytes: Number(info.size) };
  },
});

function ctxFor({ ws, mode, live = true, fs = Object.create(SandboxedFileSystem.prototype), services = {} }) {
  const policy = { resolve: ({ session } = {}) => ({ mode: session ? mode : 'danger-full-access', workspaceRoot: session?.header.cwd ?? '/' }) };
  const sessions = { get: id => live ? { id, header: { cwd: ws } } : undefined };
  const all = { sandboxPolicy: policy, sessions, fs, ...services };
  return { get: name => all[name] };
}
const writerFor = (f, options) => createWriter({ ctx: ctxFor({ ws: f.ws, ...options }), workspaceFiles: workspaceFiles(f.ws), dataRoot: f.data, writableRoots, atomic: { withFileLock, writeFileAtomic } });
const scope = f => ({ sessionId: 's', workspaceRoot: f.ws });

test('write-back matches dsh-fs-sandbox checkedTarget for every mode and path', async t => {
  const f = await workspace(t);
  const oracle = Object.create(SandboxedFileSystem.prototype);
  oracle.resolve = async displayPath => ({ targetKey: await canonicalTarget(displayPath), displayPath });
  const cases = [join(f.ws, 'new.docx'), join(f.ws, 'nested', 'deep', 'new.docx'), join(f.outside, 'new.docx'), join(f.ws, 'escape', 'new.docx'), join(tmpdir(), 'suite-tmp-new.docx'), '/tmp/suite-tmp-new.docx'];
  for (const mode of ['read-only', 'workspace-write', 'danger-full-access']) {
    for (const path of cases) {
      const expected = await oracle.checkedTarget({ displayPath: path }, { mode, workspaceRoot: f.ws }).then(() => 'allow', error => error.code === 'FS_SANDBOX_DENIED' ? 'deny' : error);
      const writer = writerFor(f, { mode });
      const actual = await writer.create(scope(f), path, Buffer.from('x'), signal()).then(() => 'allow', error => error.code === 'sandbox-denied' ? 'deny' : error);
      assert.equal(actual, expected, `${mode} ${path}`);
      if (actual === 'allow') await rm(path, { force: true });
    }
  }
});

test('replace checks the live Session policy, the version and leaves no lock in the workspace', async t => {
  const f = await workspace(t);
  const files = workspaceFiles(f.ws);
  const before = await files.stat(scope(f), 'a.docx');
  await assert.rejects(writerFor(f, { mode: 'read-only' }).replace(scope(f), 'a.docx', Buffer.from('new'), { expectedVersion: before.version }, signal()), code('sandbox-denied'), 'deployment workspace-write, Session switched to read-only');
  await assert.rejects(writerFor(f, { mode: 'workspace-write', live: false }).replace(scope(f), 'a.docx', Buffer.from('new'), { expectedVersion: before.version }, signal()), code('sandbox-denied'), 'cold Session');
  const writer = writerFor(f, { mode: 'workspace-write' });
  await assert.rejects(writer.replace(scope(f), 'a.docx', Buffer.from('new'), { expectedVersion: 'stale' }, signal()), code('source-changed'));
  const mode = (await stat(join(f.ws, 'a.docx'))).mode & 0o777;
  const written = await writer.replace(scope(f), 'a.docx', Buffer.from('new'), { expectedVersion: before.version }, signal());
  assert.notEqual(written.version, before.version);
  assert.equal(await readFile(join(f.ws, 'a.docx'), 'utf8'), 'new');
  assert.equal((await stat(join(f.ws, 'a.docx'))).mode & 0o777, mode, 'permission bits kept');
  assert.deepEqual((await readdir(f.ws)).filter(n => n.endsWith('.lock')), []);
  assert.ok((await readdir(join(f.data, 'locks'))).length >= 0);
  await assert.rejects(writer.create(scope(f), 'a.docx', Buffer.from('x'), signal()), code('exists'));
  assert.equal(await readFile(join(f.ws, 'a.docx'), 'utf8'), 'new', 'save-as never overwrites');
});

test('only local filesystems and complete compositions can write', async t => {
  const f = await workspace(t);
  assert.equal(isLocalFileSystem(Object.create(SandboxedFileSystem.prototype)), true);
  assert.equal(isLocalFileSystem(Object.create(LocalFileSystem.prototype)), true);
  class RemoteFileSystem {}
  assert.equal(isLocalFileSystem(new RemoteFileSystem()), false);
  assert.equal(isLocalFileSystem(undefined), false);
  const remote = writerFor(f, { mode: 'workspace-write', fs: new RemoteFileSystem() });
  assert.equal(remote.canWrite(), false);
  await assert.rejects(remote.create(scope(f), 'n.docx', Buffer.from('x'), signal()), code('unsupported-operation'));
  const partial = createWriter({ ctx: { get: name => name === 'fs' ? Object.create(SandboxedFileSystem.prototype) : undefined }, workspaceFiles: workspaceFiles(f.ws), dataRoot: f.data, writableRoots, atomic: { withFileLock, writeFileAtomic } });
  assert.equal(partial.canWrite(), false);
  await assert.rejects(partial.create(scope(f), 'n.docx', Buffer.from('x'), signal()), code('unsupported-operation'));
});

test('suite errors map to officesdk codes with status; nothing else leaks', () => {
  for (const name of ['busy', 'timeout', 'source-changed', 'sandbox-denied', 'password-protected']) {
    const remote = toRemote(new SuiteError(name, 'm', { draftId: 'd', path: '/secret' }), signal(), RemoteError);
    assert.equal(remote.code, `officesdk/${name}`);
    assert.equal(remote.details.status, new SuiteError(name, 'm').status);
    assert.equal(remote.details.draftId, 'd');
    assert.equal(remote.details.path, undefined, 'paths never cross the wire');
  }
  const aborted = new AbortController(); aborted.abort();
  assert.equal(toRemote(new SuiteError('busy', 'm'), aborted.signal, RemoteError).code, 'gateway/cancelled');
  // Kit's Windows Save As refusal keeps its own code and a conflict status on the way to the Client.
  const busy = toRemote(Object.assign(new Error('Another program has that file open; close it there, then save again.'), { name: 'KitError', code: 'target-busy' }), signal(), RemoteError);
  assert.deepEqual([busy.code, busy.message.startsWith('Another program'), busy.details?.status], ['officesdk/target-busy', true, 409]);
  const opaque = toRemote(new Error('/Users/x/secret failed'), signal(), RemoteError);
  assert.equal(opaque.code, 'gateway/internal');
  assert.doesNotMatch(opaque.message, /secret/);
  const passthrough = new RemoteError('gateway/bad-request', 'x', {});
  assert.equal(toRemote(passthrough, signal(), RemoteError), passthrough);
});

test('a SuiteError from another copy of Core maps like our own; workspace-file mapping is unchanged', async () => {
  // The release bundles Core once per entry; an error made by another bundle's copy is a different class.
  const other = await import('../../core/src/errors.mjs?other-copy');
  const foreign = new other.SuiteError('busy', 'm', { draftId: 'd' });
  assert.equal(foreign instanceof SuiteError, false);
  const remote = toRemote(foreign, signal(), RemoteError);
  assert.equal(remote.code, 'officesdk/busy');
  assert.equal(remote.details.status, 429);
  assert.equal(remote.details.draftId, 'd');
  assert.equal(toRemote(Object.assign(new Error('m'), { name: 'SuiteError', code: 'toString' }), signal(), RemoteError).code, 'gateway/internal');
  assert.equal(fromWorkspaceFile(foreign), foreign);
  const named = Object.assign(new Error('x'), { name: 'SuiteError', code: 'workspace-file/not-found' });
  assert.equal(fromWorkspaceFile(named).code, 'not-found', 'an unknown code is still mapped, not passed through');
});

test('WorkspaceFiles failures map to suite codes', () => {
  const table = { 'workspace-file/not-found': 'not-found', 'workspace-file/too-large': 'input-too-large', 'workspace-file/not-regular-file': 'invalid-path',
    'workspace-file/unknown-workspace': 'forbidden', 'workspace-file/unsupported-address': 'forbidden', 'workspace-file/something-new': 'unavailable' };
  for (const [source, expected] of Object.entries(table)) assert.equal(fromWorkspaceFile(Object.assign(new Error('x'), { code: source })).code, expected, source);
  const plain = new Error('x');
  assert.equal(fromWorkspaceFile(plain), plain);
});

test('authorizer checks the Session lookup and asks remoteAccess for non-owners per action', async () => {
  const decisions = [];
  let allow = true;
  const ctx = {
    typert: { lookups: { get: () => ({ resolve: async id => id === 's' ? { sessionId: 's', workspaceRoot: '/w' } : undefined }) } },
    get: name => name === 'remoteAccess' ? { decide: request => { decisions.push(request.action); return { kind: allow ? 'allow' : 'deny' }; } } : undefined,
  };
  const authorizer = createAuthorizer(ctx);
  const owner = { remote: { principal: { kind: 'owner' } } };
  const guest = { remote: { principal: { kind: 'share', id: 'g' } } };
  await authorizer.authorize(owner, { sessionId: 's', workspaceRoot: '/w' }, signal(), { action: 'save', path: 'a', revision: 1 });
  assert.deepEqual(decisions, []);
  for (const action of ['open', 'edit', 'save', 'export']) await authorizer.authorize(guest, { sessionId: 's', workspaceRoot: '/w' }, signal(), { action, path: 'a', revision: 1 });
  assert.deepEqual(decisions, ['shimoPreview/open', 'shimoPreview/edit', 'shimoPreview/save', 'shimoPreview/export']);
  allow = false;
  await assert.rejects(authorizer.authorize(guest, { sessionId: 's', workspaceRoot: '/w' }, signal(), { action: 'open' }), code('forbidden'));
  await assert.rejects(authorizer.authorize(owner, { sessionId: 's', workspaceRoot: '/other' }, signal(), { action: 'open' }), code('forbidden'));
  await assert.rejects(authorizer.authorize(owner, { sessionId: 'gone', workspaceRoot: '/w' }, signal(), { action: 'open' }), code('forbidden'));
  assert.equal(principalKey({ kind: 'owner' }), 'owner');
  assert.equal(principalKey({ id: 'g', kind: 'share' }), principalKey({ kind: 'share', id: 'g' }));
  assert.notEqual(principalKey({ kind: 'share', id: 'g' }), principalKey({ kind: 'share', id: 'h' }));
});

test('Remote parameter names stay the Typert lookup contract', () => {
  const params = name => /\(([^)]*)\)/.exec(ShimoPreview.prototype[name].toString())[1].split(',').map(s => s.trim());
  assert.deepEqual(params('open'), ['workspaceFileScope', 'path', 'revision', 'mode', 'signal']);
  assert.deepEqual(params('exportTo'), ['workspaceFileScope', 'id', 'editRevision', 'target', 'signal']);
  assert.deepEqual(params('rebind'), ['workspaceFileScope', 'id', 'revision', 'diverged', 'signal']);
  assert.deepEqual(params('release'), ['workspaceFileScope', 'id', 'autosave', 'signal']);
  assert.deepEqual(params('renew'), ['workspaceFileScope', 'id', 'signal']);
  assert.deepEqual(params('access'), ['workspaceFileScope', 'id', 'signal']);
  assert.deepEqual(params('agentNext'), ['workspaceFileScope', 'signal']);
  assert.deepEqual(params('save'), ['workspaceFileScope', 'id', 'editRevision', 'approved', 'overwrite', 'signal']);
  assert.deepEqual(params('resolveDiskChange'), ['workspaceFileScope', 'id', 'useDisk', 'signal']);
  assert.deepEqual(params('agentAlive'), ['workspaceFileScope', 'requestId', 'signal']);
  assert.deepEqual(params('agentReply'), ['workspaceFileScope', 'requestId', 'ok', 'result', 'code', 'message', 'signal']);
  assert.deepEqual(params('updateStatus'), ['workspaceFileScope', 'refresh', 'signal']);
  assert.deepEqual(params('probe'), ['workspaceFileScope', 'path', 'kind', 'signal']);
  assert.deepEqual(params('browse'), ['workspaceFileScope', 'path', 'signal']);
  assert.deepEqual(params('listStyles'), ['workspaceFileScope', 'signal']);
  assert.deepEqual(params('rememberStyle'), ['workspaceFileScope', 'style', 'signal']);
  assert.deepEqual(params('selectTemplate'), ['workspaceFileScope', 'kind', 'id', 'signal']);
  assert.deepEqual(params('templatePick'), ['workspaceFileScope', 'signal']);
  assert.deepEqual(params('styleSuggestion'), ['workspaceFileScope', 'signal']);
  assert.deepEqual(params('dismissStyleSuggestion'), ['workspaceFileScope', 'id', 'signal']);
  assert.deepEqual(params('whatsNew'), ['workspaceFileScope', 'signal']);
  assert.deepEqual(params('whatsNewSeen'), ['workspaceFileScope', 'version', 'signal']);
});

test('update status says only the owner may install the new release', async () => {
  const status = principal => ShimoPreview.prototype.updateStatus.call({ remote: ShimoPreview.prototype.remote, remotePeer: () => ({ principal }),
    updater: { status: async () => ({ current: '0.0.28', latest: '0.0.29', available: true }) } }, {}, false, signal());
  assert.deepEqual(await status('owner'), { current: '0.0.28', latest: '0.0.29', available: true, canUpdate: true });
  assert.equal((await status('share:g')).canUpdate, false);
});

test('the update notes are the owner\'s: others get none and cannot mark them seen', async () => {
  const marked = [];
  const service = principal => ({ remote: ShimoPreview.prototype.remote, remotePeer: () => ({ principal }),
    whatsNewNotes: { pending: async () => ({ current: '0.0.47', entries: [{ version: '0.0.47', highlights: ['u1'] }] }), markSeen: async version => { if (version === 'bad') throw new Error('no'); marked.push(version); } } });
  assert.equal((await ShimoPreview.prototype.whatsNew.call(service('owner'), {}, signal())).current, '0.0.47');
  assert.equal(await ShimoPreview.prototype.whatsNew.call(service('share:g'), {}, signal()), null);
  assert.deepEqual(await ShimoPreview.prototype.whatsNewSeen.call(service('owner'), {}, '0.0.47', signal()), { seen: true });
  await assert.rejects(ShimoPreview.prototype.whatsNewSeen.call(service('share:g'), {}, '0.0.47', signal()), error => error.code === 'officesdk/forbidden' || /forbidden|owner/i.test(String(error?.code ?? error?.message)));
  await assert.rejects(ShimoPreview.prototype.whatsNewSeen.call(service('owner'), {}, 'bad', signal()));
  assert.deepEqual(marked, ['0.0.47']);
});

test('updating counts every Kit view (other windows too) and stops Kit on Windows; refusals keep their codes', async () => {
  const update = (views, { platform = 'win32', principal = 'owner', fail, version = '0.0.40' } = {}) => {
    const events = [];
    const self = { remote: ShimoPreview.prototype.remote, remotePeer: () => ({ principal }), platform, kitDocs: { views: new Map(views.map((view, i) => [`v${i}`, view])) },
      kit: { suspend: async () => { events.push('suspend'); }, resume: () => { events.push('resume'); } },
      updater: { checked: { latest: '0.0.40' }, update: async target => { assert.equal(target, '0.0.40'); events.push('update'); if (fail) throw fail; return { installed: '0.0.40', reloading: true }; } } };
    return { events, result: ShimoPreview.prototype.updatePlugin.call(self, {}, version, signal()) };
  };
  const stale = update([], { version: '0.0.39' });
  await assert.rejects(stale.result, { code: 'officesdk/unavailable' });
  assert.deepEqual(stale.events, []);
  const refused = update([{ unsaved: false }, { unsaved: true, principal: 'share:g' }]);
  await assert.rejects(refused.result, { code: 'officesdk/unsaved-changes' });
  assert.deepEqual(refused.events, []);
  const done = update([{ unsaved: false }]);
  assert.deepEqual(await done.result, { installed: '0.0.40', reloading: true });
  assert.deepEqual(done.events, ['suspend', 'update']);
  const failed = update([], { fail: Object.assign(new Error('https://user:token@registry.example.invalid/x failed'), { stderr: 'npm error https://user:token@registry.example.invalid/x' }) });
  await assert.rejects(failed.result, error => error.code === 'officesdk/unavailable' && !/token/.test(error.message));
  assert.deepEqual(failed.events, ['suspend', 'update', 'resume']);
  await assert.rejects(update([], { principal: 'share:g' }).result, { code: 'officesdk/forbidden' });
});

test('POST origin must be this host or the desktop app scheme', () => {
  const req = { headers: { host: '127.0.0.1:7788' } };
  assert.equal(sameOrigin('http://127.0.0.1:7788', req), true);
  assert.equal(sameOrigin('dsh-app://app', req), true);
  assert.equal(sameOrigin('http://evil.example', req), false);
  assert.equal(sameOrigin('http://127.0.0.1:9999', req), false);
  assert.equal(sameOrigin('null', req), false);
  assert.equal(sameOrigin('file://127.0.0.1:7788', req), false);
});
