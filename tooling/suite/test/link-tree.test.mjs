import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { linkTree, replaceFile } from '../link-tree.mjs';

const sha = async file => createHash('sha256').update(await fs.readFile(file)).digest('hex');

async function fixture(t) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'link-tree-'));
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const src = path.join(dir, 'src');
  await fs.mkdir(path.join(src, 'bin'), { recursive: true });
  await fs.writeFile(path.join(src, 'package.json'), JSON.stringify({ name: 'link-tree-fixture', version: '1.0.0', files: ['bin', 'data.txt', 'catalog.json'] }));
  await fs.writeFile(path.join(src, 'data.txt'), 'data');
  await fs.writeFile(path.join(src, 'catalog.json'), '{"v":1}');
  await fs.writeFile(path.join(src, 'bin', 'tool'), '#!/bin/sh\n');
  await fs.chmod(path.join(src, 'bin', 'tool'), 0o755);
  await fs.chmod(path.join(src, 'bin'), 0o750);
  await fs.writeFile(path.join(src, '.complete'), 'marker');
  await fs.symlink('data.txt', path.join(src, 'alias'));
  return { dir, src };
}

test('linkTree: hard links with modes, links and the filter kept', async t => {
  const { dir, src } = await fixture(t);
  const to = path.join(dir, 'nested', 'to');
  const counts = await linkTree(src, to, { filter: from => path.basename(from) !== '.complete' });
  assert.deepEqual(counts, { linked: 4, copied: 0 });
  assert.equal((await fs.stat(path.join(to, 'data.txt'))).ino, (await fs.stat(path.join(src, 'data.txt'))).ino);
  assert.equal((await fs.stat(path.join(to, 'bin', 'tool'))).mode & 0o777, 0o755);
  assert.equal((await fs.stat(path.join(to, 'bin'))).mode & 0o777, 0o750);
  assert.equal(await fs.readlink(path.join(to, 'alias')), 'data.txt');
  await assert.rejects(fs.access(path.join(to, '.complete')));
});

test('replaceFile on a linked tree leaves the source as it was and keeps the mode', async t => {
  const { dir, src } = await fixture(t);
  const to = path.join(dir, 'to');
  await linkTree(src, to);
  await replaceFile(path.join(to, 'bin', 'tool'), '#!/bin/sh\necho changed\n');
  await replaceFile(path.join(to, 'new.json'), '{}');
  assert.equal(await fs.readFile(path.join(src, 'bin', 'tool'), 'utf8'), '#!/bin/sh\n');
  assert.equal(await fs.readFile(path.join(to, 'bin', 'tool'), 'utf8'), '#!/bin/sh\necho changed\n');
  assert.equal((await fs.stat(path.join(to, 'bin', 'tool'))).mode & 0o777, 0o755);
  assert.equal(await fs.readFile(path.join(to, 'new.json'), 'utf8'), '{}');
  assert.deepEqual((await fs.readdir(path.join(to, 'bin'))).filter(name => name.endsWith('.tmp')), []);
  await fs.rm(path.join(to, 'data.txt'));
  assert.equal(await fs.readFile(path.join(src, 'data.txt'), 'utf8'), 'data');
});

test('linkTree copies an inode it meets twice, so one tree never has two names for it', async t => {
  const { dir, src } = await fixture(t);
  await fs.link(path.join(src, 'data.txt'), path.join(src, 'data-again.txt'));
  const to = path.join(dir, 'to');
  const counts = await linkTree(src, to);
  assert.deepEqual(counts, { linked: 5, copied: 1 });
  assert.notEqual((await fs.stat(path.join(to, 'data.txt'))).ino, (await fs.stat(path.join(to, 'data-again.txt'))).ino);
});

test('linkTree copies when links are refused (another mount) and keeps modes', async t => {
  const { dir, src } = await fixture(t);
  t.mock.method(fs, 'link', async () => { throw Object.assign(new Error('cross-device link'), { code: 'EXDEV' }); });
  const to = path.join(dir, 'to');
  const counts = await linkTree(src, to);
  assert.deepEqual(counts, { linked: 0, copied: 5 });
  assert.notEqual((await fs.stat(path.join(to, 'data.txt'))).ino, (await fs.stat(path.join(src, 'data.txt'))).ino);
  assert.equal((await fs.stat(path.join(to, 'bin', 'tool'))).mode & 0o777, 0o755);
});

test('npm pack of a linked tree is byte for byte the pack of a copied one', async t => {
  const { dir, src } = await fixture(t);
  await fs.rm(path.join(src, 'alias'));
  const linked = path.join(dir, 'linked');
  const copied = path.join(dir, 'copied');
  await linkTree(src, linked);
  await fs.cp(src, copied, { recursive: true });
  const pack = async from => {
    const result = spawnSync('npm', ['pack', '--silent', '--pack-destination', from, '--cache', path.join(dir, 'npm-cache')], { cwd: from, encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr);
    return sha(path.join(from, result.stdout.trim().split(/\r?\n/).at(-1)));
  };
  assert.equal(await pack(linked), await pack(copied));
});
