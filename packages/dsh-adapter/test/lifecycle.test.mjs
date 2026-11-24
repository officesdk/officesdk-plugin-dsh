import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { Context } from '@deepseek-ai/cordis';
import { SandboxedFileSystem } from '@deepseek-ai/dsh-fs-sandbox';
import ShimoPreview from '../src/host/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';

async function load(t, services = {}) {
  const root = await mkdtemp(join(tmpdir(), 'suite-plugin-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  await mkdir(join(root, 'runtime')); await writeFile(join(root, 'runtime', 'preview.html'), '');
  const app = new Context();
  const routes = [];
  const base = {
    workspaceFiles: { stat: async () => ({ absolutePath: '/w/a.docx', version: 'v1', bytes: 3 }), readBytes: async (s, p, o) => ({ absolutePath: '/w/a.docx', version: 'v1', data: Buffer.from('abc'), eof: true }) },
    webServer: { register: route => { routes.push(route.path); return () => routes.push(`-${route.path}`); } },
    connection: { admit: () => ({ rejection: 403 }) },
    typert: { lookups: { get: () => ({ resolve: async id => ({ sessionId: id, workspaceRoot: '/w' }) }) } },
    ...services,
  };
  for (const [name, value] of Object.entries(base)) { app.provide(name); app.set(name, value); }
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: join(root, 'data') };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  const fiber = app.plugin(Plugin, {});
  await fiber;
  await new Promise(resolve => setTimeout(resolve, 20));
  return { app, fiber, routes, root, service: app.get('shimoPreview') };
}

test('loads without the optional write services and previews read-only', async t => {
  const plugin = await load(t);
  assert.deepEqual(plugin.routes, ['/shimo']);
  assert.ok(plugin.service, 'service registered');
  assert.deepEqual(plugin.service.suite.capabilities('docx'), { import: true, preview: true, edit: false, save: false, export: true });
  const peer = { id: 'p', principal: 'owner', signal: new AbortController().signal, remote: { principal: { kind: 'owner' } } };
  const handle = await plugin.service.suite.open({ sessionId: 's', workspaceRoot: '/w' }, 'a.docx', 1, new AbortController().signal, peer);
  assert.equal(handle.capabilities.save, false);
  await assert.rejects(plugin.service.suite.open({ sessionId: 's', workspaceRoot: '/w' }, 'a.docx', 1, new AbortController().signal, peer, { mode: 'edit' }), { code: 'unsupported-operation' });
});

test('with a local fs, sandbox policy and sessions the same plugin can edit and save', async t => {
  const plugin = await load(t, { fs: Object.create(SandboxedFileSystem.prototype), sandboxPolicy: { resolve: () => ({ mode: 'workspace-write', workspaceRoot: '/w' }) }, sessions: { get: id => ({ id, header: { cwd: '/w' } }) } });
  assert.equal(plugin.service.suite.capabilities('xlsx').save, true);
});

test('unloading unregisters the route and leaves no task directories', async t => {
  const plugin = await load(t);
  const peer = { id: 'p', principal: 'owner', signal: new AbortController().signal, remote: { principal: { kind: 'owner' } } };
  await plugin.service.suite.open({ sessionId: 's', workspaceRoot: '/w' }, 'a.docx', 1, new AbortController().signal, peer);
  assert.equal((await readdir(join(plugin.root, 'tmp'))).filter(n => n.startsWith('officesdk-')).length, 1);
  await plugin.fiber.dispose();
  await new Promise(resolve => setTimeout(resolve, 50));
  assert.deepEqual(plugin.routes, ['/shimo', '-/shimo']);
  assert.deepEqual((await readdir(join(plugin.root, 'tmp'))).filter(n => n.startsWith('officesdk-')), []);
});
