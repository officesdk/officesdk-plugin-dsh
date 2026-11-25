import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DEFAULT_SOURCES, NPM_DEFAULTS, readDownloadSources, selectProfile, sourcesFor, sourcesProblems } from '../src/host/fetch/download-sources.mjs';
import { officeKitItems } from '../src/host/fetch/office-kit.mjs';
import { SkillStore } from '../src/skills/store.mjs';

const temp = () => mkdtempSync(join(tmpdir(), 'sources-'));
const mac = { platform: 'darwin', arch: 'arm64' };
const manifest = () => ({ officeKit: { version: '0.1.3', protocol: '5' }, sources: [{ id: 'release', kind: 'url', base: 'https://release.example/0.1.3/' }],
  kit: { 'darwin-arm64': { file: 'k.tgz', size: 1, integrity: 'sha512-k', binary: 'offline-kit-runtime', status: 'stable', files: {} } },
  editor: { file: 'e.tgz', size: 2, integrity: 'sha512-e', root: 'package/dist/resources', files: {}, npm: { registry: 'https://npm.example.invalid/', package: '@example/offline-kit-editor', version: '0.1.3' } } });

test('the public profile is the only one and the default: the Kit from office-kit\'s release, the editor from npm', () => {
  assert.deepEqual(sourcesProblems(DEFAULT_SOURCES), []);
  assert.deepEqual(Object.keys(DEFAULT_SOURCES.profiles), ['public']);
  assert.deepEqual(DEFAULT_SOURCES.profiles.public.kit, ['mirror', 'release']);
  assert.deepEqual(DEFAULT_SOURCES.profiles.public.editor, ['mirror', 'npm']);
  const items = officeKitItems(manifest(), { root: '/d', host: mac });
  assert.deepEqual(items.editor.sources.map(s => s.id), ['npm']);
  assert.deepEqual(items.kit.sources.map(s => [s.id, s.base]), [['release', 'https://release.example/0.1.3/']]);
  const shipped = JSON.parse(readFileSync(new URL('../download-sources.json', import.meta.url), 'utf8'));
  assert.deepEqual(shipped, JSON.parse(JSON.stringify(DEFAULT_SOURCES)), 'the committed file agrees');
  const mirrored = officeKitItems(manifest(), { root: '/d', host: mac, mirror: 'https://m/' });
  assert.deepEqual(mirrored.editor.sources.map(s => s.id), ['mirror', 'npm'], 'a mirror stays first');
  // npm is the editor's last source: the long limits.
  const npm = items.editor.sources[0];
  assert.deepEqual([npm.registry, npm.package, npm.version, npm.connectTimeoutMs, npm.idleMs, npm.attempts], ['https://npm.example.invalid/', '@example/offline-kit-editor', '0.1.3', 15_000, 30_000, 3]);
  assert.deepEqual(NPM_DEFAULTS, { connectTimeoutMs: 15_000, idleMs: 30_000, attempts: 3 });
  assert.throws(() => sourcesFor('internal'), /unknown sources profile/);
  assert.deepEqual(sourcesProblems({ ...sourcesFor('public'), profile: 'internal' }), ['download-sources.json: profile must be one of public']);
});

test('the release carries no editor, so it is not an editor source; npm is not a Kit source', () => {
  const doc = sourcesFor('public');
  doc.profiles.public.editor = ['mirror', 'release'];
  doc.profiles.public.kit = ['mirror', 'npm'];
  assert.deepEqual(sourcesProblems(doc), ['profiles.public.kit: unknown source "npm"', 'profiles.public.editor: unknown source "release"']);
  // The id of the old storage-bucket source (before 0.3.0-rc.1) is refused like any unknown id; no alias.
  const legacy = ['o', 'b', 's'].join('');
  assert.deepEqual(sourcesProblems({ ...sourcesFor('public'), profiles: { public: { ...sourcesFor('public').profiles.public, kit: [legacy] } } }), [`profiles.public.kit: unknown source "${legacy}"`]);
  const profile = selectProfile(sourcesFor('public'), {});
  assert.deepEqual(officeKitItems(manifest(), { root: '/d', host: mac, profile, mirror: 'https://m/' }).kit.sources.map(s => s.id), ['mirror', 'release'], 'an explicit mirror still works');
});

test('the file decides the order and limits; a broken or missing file falls back to the defaults', () => {
  const root = temp();
  assert.equal(readDownloadSources(root), DEFAULT_SOURCES, 'missing');
  const doc = sourcesFor('public');
  doc.profiles.public.editor = [{ id: 'npm', connectTimeoutMs: 1500, idleMs: 4000 }, 'mirror'];
  writeFileSync(join(root, 'download-sources.json'), JSON.stringify(doc));
  const items = officeKitItems(manifest(), { root: '/d', host: mac, mirror: 'https://m/', profile: selectProfile(readDownloadSources(root), {}) });
  assert.deepEqual(items.editor.sources.map(s => s.id), ['npm', 'mirror']);
  assert.deepEqual([items.editor.sources[0].connectTimeoutMs, items.editor.sources[0].idleMs, items.editor.sources[0].attempts], [1500, 4000, 3]);
  const logs = [];
  writeFileSync(join(root, 'download-sources.json'), JSON.stringify({ ...doc, profiles: { ...doc.profiles, public: { ...doc.profiles.public, kit: ['ftp'] } } }));
  assert.equal(readDownloadSources(root, (level, text) => logs.push(text)), DEFAULT_SOURCES);
  assert.match(logs[0], /unknown source/);
  writeFileSync(join(root, 'download-sources.json'), '{');
  assert.equal(readDownloadSources(root, () => {}), DEFAULT_SOURCES);
  assert.deepEqual(sourcesProblems({ ...sourcesFor('public', { publicBase: 'http://example.com/x' }) }), ['profiles.public.skills.publicBase must be https (or a loopback http address for drills)']);
  assert.deepEqual(sourcesProblems(sourcesFor('public', { publicBase: 'http://127.0.0.1:8080/skills/' })), []);
});

test('OFFICESDK_SKILLS_BASE replaces publicBase', () => {
  assert.equal(selectProfile(sourcesFor('public', { publicBase: 'https://a/skills/' }), {}).skills.publicBase, 'https://a/skills');
  assert.equal(selectProfile(sourcesFor('public'), { OFFICESDK_SKILLS_BASE: 'https://b/s/' }).skills.publicBase, 'https://b/s');
});

const entry = { version: '0.0.4', files: { 'SKILL.md': { sha512: 'sha512-x' } }, source: { tag: 'skill/shimo-xlsx/v0.0.4', path: 'packages/dsh-adapter/skills/shimo-xlsx' }, pack: { file: 'shimo-xlsx-0.0.4.tgz', size: 9, integrity: 'sha512-p' } };
const store = options => new SkillStore({ packageRoot: temp(), dataRoot: temp(), names: ['shimo-xlsx'], pluginVersion: '0.1.0', ...options });

test('skill sources follow the profile: a mirror only with a base; bundled by default', () => {
  assert.deepEqual(store({}).item('shimo-xlsx', entry).sources, [], 'bundled default: nothing to download');
  const pub = selectProfile(sourcesFor('public', { publicBase: 'https://cdn.example/officesdk-skills/' }), {}).skills;
  const sources = store({ sources: pub }).item('shimo-xlsx', entry).sources;
  assert.deepEqual(sources.map(s => s.id), ['public']);
  assert.equal(sources[0].kind, 'url');
  assert.equal(`${sources[0].base}${entry.pack.file}`, 'https://cdn.example/officesdk-skills/shimo-xlsx-0.0.4/shimo-xlsx-0.0.4.tgz');
  assert.deepEqual(store({ sources: selectProfile(sourcesFor('public'), {}).skills }).item('shimo-xlsx', entry).sources, [], 'no base: nothing to try (not a timeout)');
  assert.deepEqual(store({ sources: pub }).item('shimo-xlsx', { ...entry, pack: undefined }).sources, [], 'without a pinned pack the public copy cannot be checked as a whole');
});

test('the skill catalog comes from the public mirror in the public profile', async () => {
  const asked = [];
  const fetchImpl = async url => { asked.push(url); return new Response('{"v":1,"skills":{}}', { status: 200 }); };
  const pub = selectProfile(sourcesFor('public', { publicBase: 'https://cdn.example/officesdk-skills' }), {}).skills;
  assert.equal(await store({ sources: pub, fetchImpl }).fetchCatalog(1000), '{"v":1,"skills":{}}');
  assert.deepEqual(asked, ['https://cdn.example/officesdk-skills/catalog/skills-catalog.json']);
  await assert.rejects(store({ sources: selectProfile(sourcesFor('public'), {}).skills, fetchImpl }).fetchCatalog(1000), /no skill catalog source/);
  const failing = async () => new Response('', { status: 503 });
  await assert.rejects(store({ sources: pub, fetchImpl: failing }).fetchCatalog(1000), /public catalog: HTTP 503/);
});

test('the skills component reads the same download profile as the Host (smoke 261111)', async () => {
  const { skillStoreFor } = await import('../src/skills/index.mjs');
  const packageRoot = temp(); const dataRoot = temp();
  writeFileSync(join(packageRoot, 'package.json'), JSON.stringify({ version: '0.1.1' }));
  writeFileSync(join(packageRoot, 'download-sources.json'), JSON.stringify(sourcesFor('public', { publicBase: 'https://cdn.example/s' })));
  assert.deepEqual(skillStoreFor({ packageRoot, dataRoot }).catalogIds, ['public']);
  writeFileSync(join(packageRoot, 'download-sources.json'), JSON.stringify(sourcesFor('public')));
  assert.deepEqual(skillStoreFor({ packageRoot, dataRoot }).catalogIds, ['bundled']);
});

test('a public-profile download before the first catalog fetches the catalog instead of failing with no source (smoke 261111)', async () => {
  const packageRoot = temp();
  writeFileSync(join(packageRoot, 'skills-catalog.json'), JSON.stringify({ v: 1, layer: 'light', remoteCatalog: true, skills: { 'shimo-xlsx': [{ ...entry, pack: undefined, channel: 'stable', minPluginVersion: '0.1.0' }] } }));
  const asked = []; const items = [];
  const pub = selectProfile(sourcesFor('public', { publicBase: 'https://cdn.example/s' }), {}).skills;
  const fetchImpl = async url => { asked.push(url); return new Response(JSON.stringify({ v: 1, skills: { 'shimo-xlsx': [{ ...entry, channel: 'stable', minPluginVersion: '0.1.0' }] } }), { status: 200 }); };
  const downloader = { ensure: async item => { items.push(item); return { dir: '/x', source: 'public' }; }, isComplete: () => false };
  const s = new SkillStore({ packageRoot, dataRoot: temp(), names: ['shimo-xlsx'], pluginVersion: '0.1.0', sources: pub, fetchImpl, downloader });
  const got = await s.ensure('shimo-xlsx');
  assert.equal(got.source, 'public');
  assert.deepEqual(asked, ['https://cdn.example/s/catalog/skills-catalog.json']);
  assert.deepEqual(items[0].sources.map(source => source.id), ['public']);
  assert.equal(items[0].archive.integrity, 'sha512-p');
});

test('the public profile carries its skills (bundled, no mirror); a publicBase drill still names the public mirror (261113)', () => {
  assert.deepEqual(DEFAULT_SOURCES.profiles.public.skills, { sources: ['bundled'], catalog: ['bundled'] });
  assert.deepEqual(selectProfile(sourcesFor('public'), {}).skills, { sources: ['bundled'], catalog: ['bundled'], publicBase: '' });
  assert.deepEqual(sourcesProblems(sourcesFor('public')), []);
  assert.deepEqual(sourcesFor('public', { publicBase: 'https://cdn.example/s/' }).profiles.public.skills, { sources: ['public'], catalog: ['public'], publicBase: 'https://cdn.example/s' });
  const doc = sourcesFor('public'); doc.profiles.public.skills = { sources: ['elsewhere'], catalog: [] };
  assert.deepEqual(sourcesProblems(doc), ['profiles.public.skills.sources: unknown source elsewhere']);
});

test('a profile never mixes bundled skills with download sources (impl review R1-7)', () => {
  const doc = sourcesFor('public'); doc.profiles.public.skills = { sources: ['bundled', 'public'], catalog: ['bundled'], publicBase: 'https://cdn.example/s' };
  assert.deepEqual(sourcesProblems(doc), ['profiles.public.skills: bundled cannot be mixed with download sources']);
});
