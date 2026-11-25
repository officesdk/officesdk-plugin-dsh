import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { availableVersions, catalogOf, catalogText, compareBase, entryProblems, skillEntry, skillFields } from '../src/skills/catalog.mjs';
import { SkillStore } from '../src/skills/store.mjs';
import { createStoreProvider } from '../src/skills/index.mjs';
import { Downloader } from '../src/host/fetch/downloader.mjs';
import { makeTar } from './fixtures/tar-writer.mjs';

const temp = () => mkdtempSync(join(tmpdir(), 'skills-'));
const skillMd = (version, extra = '', description = `xlsx ${version}`) => `---\nname: shimo-xlsx\ndescription: ${description}\nwhenToUse: 表格\nmetadata:\n  version: ${version}\n${extra}---\nbody ${version}\n`;
const files = (version, extra = {}) => ({ 'SKILL.md': Buffer.from(skillMd(version)), 'templates/blank.xlsx': Buffer.from(`blank ${version}`), ...extra });

test('catalog entries: fields from SKILL.md, file hashes, tag, templates; equal content gives equal entries', () => {
  assert.deepEqual(skillFields(skillMd('0.0.3', '  minPluginVersion: 0.1.2\n')), { name: 'shimo-xlsx', description: 'xlsx 0.0.3', whenToUse: '表格', version: '0.0.3', minPluginVersion: '0.1.2' });
  const entry = skillEntry('shimo-xlsx', files('0.0.3'), { commit: 'c1' });
  assert.equal(entry.minPluginVersion, '0.0.0', 'default: the universal package starts at 0.0.0-alpha.1');
  assert.deepEqual(entry.source, { commit: 'c1', path: 'packages/dsh-adapter/skills/shimo-xlsx', tag: 'skill/shimo-xlsx/v0.0.3' });
  assert.deepEqual(entry.templates, ['blank']);
  assert.deepEqual(Object.keys(entry.files), ['SKILL.md', 'templates/blank.xlsx']);
  assert.deepEqual(entryProblems('shimo-xlsx', entry, skillEntry('shimo-xlsx', files('0.0.3'), { commit: 'c2' })), []);
  assert.deepEqual(entryProblems('shimo-xlsx', entry, skillEntry('shimo-xlsx', files('0.0.3', { 'templates/blank.xlsx': Buffer.from('changed') }))), ['shimo-xlsx 0.0.3: templates/blank.xlsx differs']);
  assert.throws(() => skillEntry('shimo-pptx', files('0.0.3')), /names shimo-xlsx/);
  assert.throws(() => skillEntry('shimo-xlsx', { 'SKILL.md': Buffer.from(skillMd('beta')) }), /X\.Y\.Z/);
});

test('version choice: channel, minPluginVersion on X.Y.Z only, yanked skipped, packaged always available', () => {
  const e = (version, extra = {}) => ({ version, channel: 'stable', minPluginVersion: '0.1.0', files: {}, ...extra });
  const packaged = { skills: { 'shimo-xlsx': [e('0.0.3')] } };
  const remote = { skills: { 'shimo-xlsx': [e('0.0.3'), e('0.0.4', { channel: 'test' }), e('0.0.5', { minPluginVersion: '0.2.0' }), e('0.0.6', { yanked: true }), e('0.0.2')] } };
  const versions = (pluginVersion, r = remote) => availableVersions('shimo-xlsx', { packaged, remote: r, pluginVersion }).map(entry => `${entry.version}${entry.packaged ? '*' : ''}`);
  assert.deepEqual(versions('0.1.0'), ['0.0.3*', '0.0.2'], 'a stable plugin sees stable only');
  assert.deepEqual(versions('0.1.0-test.7'), ['0.0.4', '0.0.3*', '0.0.2'], 'a test plugin also sees test; its own X.Y.Z counts for minPluginVersion');
  assert.deepEqual(versions('0.2.0-test.1'), ['0.0.5', '0.0.4', '0.0.3*', '0.0.2'], 'minPluginVersion 0.2.0 admits 0.2.0-test.1 (G4)');
  assert.deepEqual(versions('0.1.0', { skills: { 'shimo-xlsx': [e('0.0.3', { yanked: true })] } }), [], 'a yanked packaged version is dropped');
  assert.ok(compareBase('0.1.0-test.3', '0.1.0') === 0 && compareBase('0.1.10', '0.1.9') > 0);
  // Preview stages (0.0.0-alpha.1 on) see test skills like a test build; minPluginVersion compares their X.Y.Z.
  const preview = { skills: { 'shimo-xlsx': [e('0.0.7', { minPluginVersion: '0.0.0', channel: 'test' }), e('0.0.8', { minPluginVersion: '0.0.0' }), e('0.0.9')] } };
  const pick = pluginVersion => availableVersions('shimo-xlsx', { packaged: { skills: {} }, remote: preview, pluginVersion }).map(entry => entry.version);
  assert.deepEqual(pick('0.0.0-alpha.1'), ['0.0.8', '0.0.7'], 'an alpha sees test and stable at minPluginVersion 0.0.0, not 0.1.0');
  assert.deepEqual(pick('0.0.0-rc.2'), ['0.0.8', '0.0.7']);
  assert.deepEqual(pick('0.0.0'), ['0.0.8'], 'a release sees stable only');
});

const BASE = 'https://mirror.example.invalid/skills';
/** A skill mirror (`public` source) serving `<base>/<name>-<version>/<pack>` and `<base>/catalog/skills-catalog.json`. */
function mirror(packs, catalog) {
  const requests = [];
  const fetchImpl = async url => {
    requests.push(String(url));
    if (mirror.down) return new Response('', { status: 503 });
    const body = String(url) === `${BASE}/catalog/skills-catalog.json` ? Buffer.from(catalog) : packs[String(url)];
    return body ? new Response(body, { status: 200, headers: { 'content-length': String(body.length) } }) : new Response('', { status: 404 });
  };
  return { fetchImpl, requests };
}

/** A skill mirror with packed skill versions, and a package that ships `version` as the light layer. */
function setup({ packagedVersion = '0.0.3', layer = 'light', remoteVersions = ['0.0.3'], pluginVersion = '0.1.0-test.1' } = {}) {
  const entries = {}; const packs = {};
  for (const version of [...new Set([packagedVersion, ...remoteVersions])].sort()) {
    const tgz = makeTar(Object.entries(files(version)).map(([name, data]) => ({ name, data })), { gzip: true });
    const file = `shimo-xlsx-${version}.tgz`;
    packs[`${BASE}/shimo-xlsx-${version}/${file}`] = tgz;
    entries[version] = { ...skillEntry('shimo-xlsx', files(version), { commit: 'x', channel: 'test' }), pack: { file, size: tgz.length, integrity: `sha512-${createHash('sha512').update(tgz).digest('base64')}` } };
  }
  const served = mirror(packs, catalogText(catalogOf({ 'shimo-xlsx': remoteVersions.map(version => entries[version]) }, {})));
  const packageRoot = temp();
  mkdirSync(join(packageRoot, 'skills/shimo-xlsx/templates'), { recursive: true });
  writeFileSync(join(packageRoot, 'skills/shimo-xlsx/SKILL.md'), skillMd(packagedVersion));
  if (layer === 'full') writeFileSync(join(packageRoot, 'skills/shimo-xlsx/templates/blank.xlsx'), `blank ${packagedVersion}`);
  writeFileSync(join(packageRoot, 'skills-catalog.json'), catalogText(catalogOf({ 'shimo-xlsx': [entries[packagedVersion]] }, { layer, remoteCatalog: layer !== 'full' || undefined })));
  const dataRoot = temp();
  const downloader = new Downloader({ root: join(dataRoot, 'skills'), delaysMs: [1], fetchImpl: served.fetchImpl });
  const store = new SkillStore({ packageRoot, dataRoot, names: ['shimo-xlsx'], pluginVersion, sources: { sources: ['public'], catalog: ['public'], publicBase: BASE }, fetchImpl: served.fetchImpl, downloader });
  return { packageRoot, store, served };
}

test('light layer: the packaged version lists without resources, downloads from the mirror on first use, then is effective', async () => {
  const { store } = setup();
  const before = store.effective('shimo-xlsx');
  assert.deepEqual([before.version, before.dir, before.ready], ['0.0.3', undefined, false]);
  assert.match(store.skillText('shimo-xlsx'), /body 0\.0\.3/);
  const dir = await store.resolve('shimo-xlsx', 'session-1');
  assert.equal(readFileSync(join(dir, 'templates/blank.xlsx'), 'utf8'), 'blank 0.0.3');
  assert.equal(store.effective('shimo-xlsx').dir, dir);
  assert.equal(store.resolveNow('shimo-xlsx', 'session-1'), dir);
});

test('a newer version from the remote catalog: selected, downloaded, then effective; a session keeps its pinned version', async () => {
  const { store } = setup({ remoteVersions: ['0.0.3', '0.0.4'] });
  const pinned = await store.resolve('shimo-xlsx', 'old-session');
  assert.match(pinned, /0\.0\.3$/);
  assert.equal(await store.refresh(), true, 'the selected version changed');
  assert.equal(store.effective('shimo-xlsx').selected, '0.0.4');
  assert.equal(store.effective('shimo-xlsx').version, '0.0.3', 'not effective until its files are here');
  await store.ensure('shimo-xlsx');
  assert.equal(store.effective('shimo-xlsx').version, '0.0.4');
  assert.equal(await store.resolve('shimo-xlsx', 'old-session'), pinned, 'the session that started on 0.0.3 stays there');
  assert.match(await store.resolve('shimo-xlsx', 'new-session'), /0\.0\.4$/);
  assert.match(store.repin('shimo-xlsx', 'old-session'), /0\.0\.4$/, 'repin moves it to the effective version (J1)');
  assert.equal(await store.refresh(), false);
});

test('offline and bridge packages: a full layer is used in place; an offline one never asks for the remote catalog', async () => {
  const { store, packageRoot } = setup({ layer: 'full' });
  assert.equal(store.effective('shimo-xlsx').dir, join(packageRoot, 'skills', 'shimo-xlsx'));
  assert.equal(await store.resolve('shimo-xlsx', 's'), join(packageRoot, 'skills', 'shimo-xlsx'));
  const catalog = JSON.parse(readFileSync(join(packageRoot, 'skills-catalog.json'), 'utf8'));
  writeFileSync(join(packageRoot, 'skills-catalog.json'), JSON.stringify({ ...catalog, remoteCatalog: false }));
  const offline = new SkillStore({ packageRoot, dataRoot: temp(), names: ['shimo-xlsx'], pluginVersion: '0.1.0' });
  assert.equal(await offline.refresh(), false);
});

test('the public build: a skill its catalog does not list is not one of the store\'s, bundled skills need no source (261113)', async () => {
  const { store: _unused, packageRoot } = setup({ layer: 'full' });
  const catalog = JSON.parse(readFileSync(join(packageRoot, 'skills-catalog.json'), 'utf8'));
  writeFileSync(join(packageRoot, 'skills-catalog.json'), JSON.stringify({ ...catalog, remoteCatalog: false }));
  const store = new SkillStore({ packageRoot, dataRoot: temp(), names: ['shimo-pptx', 'shimo-docx', 'shimo-xlsx'], pluginVersion: '0.1.0', sources: { sources: ['bundled'], catalog: ['bundled'] } });
  assert.deepEqual(store.names, ['shimo-xlsx'], 'only what the package lists');
  assert.deepEqual(store.status().map(row => [row.name, row.state]), [['shimo-xlsx', 'ready']]);
  assert.deepEqual(store.item('shimo-xlsx', store.versions('shimo-xlsx')[0]).sources, [], 'bundled: no download source');
  assert.equal(await store.resolve('shimo-xlsx', 's'), join(packageRoot, 'skills', 'shimo-xlsx'));
  assert.equal(await store.refresh(), false);
  // The repository (no packaged catalog) keeps every name.
  assert.deepEqual(new SkillStore({ packageRoot: temp(), dataRoot: temp(), names: ['shimo-pptx', 'shimo-xlsx'], pluginVersion: '0.1.0' }).names, ['shimo-pptx', 'shimo-xlsx']);
});

test('settings: disabled skills leave the list (host tools still resolve), clear cache keeps leased versions and fetches again', async () => {
  const { store } = setup();
  const invalidations = [];
  const provider = createStoreProvider(store, { invalidate: () => invalidations.push(1) });
  assert.deepEqual((await provider.list()).map(item => item.name), ['shimo-xlsx']);
  await store.setEnabled('shimo-xlsx', false);
  assert.deepEqual(await provider.list(), []);
  assert.ok(await store.resolve('shimo-xlsx', 's1'), 'the host still gets the templates');
  await store.setEnabled('shimo-xlsx', true);
  const dir = store.effective('shimo-xlsx').dir;
  assert.deepEqual(await store.clear('shimo-xlsx'), [], 'pinned by this process: kept');
  store.pins.clear();
  assert.deepEqual(await store.clear('shimo-xlsx'), [dir]);
  await store.downloader.inflight.get('skill/shimo-xlsx/0.0.3');
  assert.ok(store.effective('shimo-xlsx').ready, 'downloaded again');
  assert.deepEqual(store.status(), [{ name: 'shimo-xlsx', enabled: true, version: '0.0.3', latest: '0.0.3', state: 'ready' }]);
});

test('provider: list and get serve one version; get downloads the listed version, or answers without resources within its budget', async () => {
  const { store } = setup();
  const invalidations = [];
  const provider = createStoreProvider(store, { invalidate: () => invalidations.push(1) });
  const [candidate] = await provider.list();
  assert.equal(candidate.version, '0.0.3');
  assert.equal(candidate.resourceBase, undefined, 'no resourceBase before the files are here');
  const full = await provider.get(candidate);
  assert.match(full.content, /^body 0\.0\.3/);
  assert.match(full.resourceBase.path, /shimo-xlsx\/0\.0\.3$/);
  assert.ok(invalidations.length >= 1);
  const [after] = await provider.list();
  assert.ok(after.resourceBase, 'listed with its resources once downloaded');
  // A source that cannot deliver: the body comes back with a note, never an error.
  const broken = setup();
  broken.store.fetchImpl = broken.store.downloader.fetchImpl = async () => new Response('', { status: 503 });
  const slow = createStoreProvider(broken.store, { budgetMs: 2000 });
  const [item] = await slow.list();
  const answer = await slow.get(item);
  assert.match(answer.content, /注意：shimo-xlsx 的样式包、模板和脚本/);
  assert.match(answer.content, /body 0\.0\.3/);
  assert.equal(answer.resourceBase, undefined);
});

test('the repository itself (no packaged catalog) keeps every skill in place, without network', async () => {
  const packageRoot = temp();
  mkdirSync(join(packageRoot, 'skills/shimo-xlsx'), { recursive: true });
  writeFileSync(join(packageRoot, 'skills/shimo-xlsx/SKILL.md'), skillMd('0.0.9'));
  const store = new SkillStore({ packageRoot, dataRoot: temp(), names: ['shimo-xlsx'], pluginVersion: '0.1.0' });
  assert.equal(store.dev, true);
  assert.equal(await store.resolve('shimo-xlsx', 's'), join(packageRoot, 'skills/shimo-xlsx'));
  assert.equal(await store.refresh(), false);
  assert.deepEqual(store.status(), [{ name: 'shimo-xlsx', enabled: true, version: '0.0.9', latest: '0.0.9', state: 'bundled' }]);
});

test('the mirror serves the catalog and the packs; a mirror that is down fails with what it said', async () => {
  const { store, served } = setup({ remoteVersions: ['0.0.3', '0.0.4'] });
  assert.equal(await store.refresh(), true, 'the catalog came from the mirror');
  const dir = await store.resolve('shimo-xlsx', 's');
  assert.match(dir, /0\.0\.4$/);
  assert.equal(readFileSync(join(dir, 'templates/blank.xlsx'), 'utf8'), 'blank 0.0.4');
  assert.ok(served.requests.includes(`${BASE}/shimo-xlsx-0.0.4/shimo-xlsx-0.0.4.tgz`));
  const fresh = setup().store;
  fresh.downloader.fetchImpl = async () => new Response('', { status: 503 });
  await assert.rejects(fresh.ensure('shimo-xlsx'), /HTTP 503/);
});
