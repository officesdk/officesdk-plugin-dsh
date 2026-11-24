import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { REQUIRED, isExecutableImage, releaseVersion, renameClient, renamePlugin, universalManifest, universalProblems } from '../universal-package.mjs';

const sri = data => `sha256-${createHash('sha256').update(data).digest('base64')}`;
const source = { name: 'officesdk', version: '0.0.50', private: true, os: ['darwin'], cpu: ['arm64'], officesdkSource: {}, files: ['lib/host', 'resources/kit', 'resources/office-kit', 'resources/runtime', 'resources/editor', 'skills', '.gitattributes'] };

test('the public (npm) manifest: npmjs public access, no maintainers or scripts, repository only when given (261111 §2)', () => {
  const bare = universalManifest({ ...source, scripts: { test: 'x' }, maintainers: [{ name: 'someone' }], license: 'UNLICENSED', engines: { node: '>=22.19.0' } }, '0.1.1');
  assert.deepEqual(bare.publishConfig, { registry: 'https://registry.npmjs.org/', access: 'public' });
  for (const key of ['maintainers', 'scripts', 'repository', 'homepage', 'bugs', 'private']) assert.equal(key in bare, false, key);
  assert.equal(bare.license, 'UNLICENSED');
  const withRepo = universalManifest(source, '0.1.1', { repository: 'https://github.com/officesdk/officesdk-demo' });
  assert.deepEqual(withRepo.repository, { type: 'git', url: 'git+https://github.com/officesdk/officesdk-demo.git' });
  assert.equal(withRepo.homepage, 'https://github.com/officesdk/officesdk-demo#readme');
  assert.deepEqual(withRepo.bugs, { url: 'https://github.com/officesdk/officesdk-demo/issues' });
});

test('an npm build for another registry publishes there', () => {
  const other = universalManifest({ ...source, scripts: { test: 'x' } }, '0.1.0', { registry: 'https://registry.example.invalid/' });
  assert.deepEqual(other.publishConfig, { registry: 'https://registry.example.invalid/', access: 'public' });
  for (const key of ['maintainers', 'scripts', 'private']) assert.equal(key in other, false, key);
});

test('the universal manifest: published name and version, no os/cpu, no platform directories, kit-manifest.json listed', () => {
  const manifest = universalManifest(source, '0.1.0-test.4');
  assert.equal(manifest.name, '@officesdk/officesdk-plugin-dsh');
  assert.equal(manifest.version, '0.1.0-test.4');
  for (const key of ['os', 'cpu', 'private', 'officesdkSource', 'maintainers']) assert.equal(key in manifest, false, key);
  // The Kit, the editor resources (they download on first use, 261113-bundle-public-skills §6) and runtime never ship.
  assert.deepEqual(manifest.files, ['lib/host', 'resources/editor', 'skills', 'kit-manifest.json', '.gitattributes', 'download-sources.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md']);
  for (const bad of ['0.1.0-win32-test.1', '0.1', 'v0.1.0', '0.1.0-test']) assert.throws(() => universalManifest(source, bad), /invalid release version/);
  assert.ok(releaseVersion('1.2.3') && releaseVersion('1.2.3-test.45'));
  assert.equal(renamePlugin('- insert:\n    - id: officesdk\n      name: officesdk\n    - id: s\n      name: officesdk/skills\n'), '- insert:\n    - id: officesdk\n      name: "@officesdk/officesdk-plugin-dsh"\n    - id: s\n      name: "@officesdk/officesdk-plugin-dsh/skills"\n');
  assert.match(renameClient("window.__ModuleLoader__.load({id:'officesdk',factory(){}})"), /\{id:"@officesdk\/officesdk-plugin-dsh",/);
  assert.throws(() => renameClient('other'), /unexpected client module header/);
});

async function tree(files) {
  const root = await mkdtemp(join(tmpdir(), 'universal-'));
  for (const [name, data] of Object.entries(files)) { await mkdir(dirname(join(root, name)), { recursive: true }); await writeFile(join(root, name), data); }
  return root;
}
const X = `sha512-${createHash('sha512').update('x').digest('base64')}`;
const CATALOG = JSON.stringify({ v: 1, layer: 'full', remoteCatalog: false, skills: {
  'shimo-pptx': [{ version: '0.0.9', files: { 'SKILL.md': { sha512: X } } }],
  'shimo-docx': [{ version: '0.0.5', files: { 'SKILL.md': { sha512: X } } }],
} });
const SOURCES = JSON.stringify({ v: 1, profile: 'public', profiles: { public: { kit: ['release'], editor: ['release'], skills: { sources: ['bundled'], catalog: ['bundled'] } } } });
const M = JSON.stringify({ editor: { files: { 'catalog.json': { sha256: sri('{}') } } } });
const base = () => ({
  ...Object.fromEntries(REQUIRED.map(name => [name, name === 'package.json' ? JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version: '0.1.0', files: ['lib'] }) : name === 'kit-manifest.json' ? M : name === 'skills-catalog.json' ? CATALOG : name === 'download-sources.json' ? SOURCES : 'x'])),
  'skills/shimo-pptx/SKILL.md': 'x', 'skills/shimo-docx/SKILL.md': 'x',
});
const FULL = { kitManifest: M, version: '0.1.0', layer: 'full' };

test('universalProblems: no package carries the editor resources; they download on first use (261113 §6)', async () => {
  assert.deepEqual(await universalProblems(await tree(base()), FULL), []);
  const bundled = await universalProblems(await tree({ ...base(), 'resources/office-kit/catalog.json': '{}' }), FULL);
  assert.deepEqual(bundled, ['resources/office-kit must not ship in the universal package']);
});

test('universalProblems: what may not ship, what must, and the manifest byte for byte', async () => {
  // A file the catalog does not list, or a layer other than the one asked for, is wrong.
  const heavy = await universalProblems(await tree({ ...base(), 'skills/shimo-docx/templates/extra.docx': 't' }), FULL);
  assert.ok(heavy.some(problem => /skills\/shimo-docx ships 2 files, the full layer has 1/.test(problem)), heavy.join(' | '));
  const light = await universalProblems(await tree(base()), { ...FULL, layer: 'light' });
  assert.ok(light.some(problem => /layer is full, expected light/.test(problem)), light.join(' | '));
  const bad = base();
  bad['package.json'] = JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version: '0.1.0', os: ['darwin'] });
  bad['resources/kit/offline-kit-runtime'] = Buffer.from([0xcf, 0xfa, 0xed, 0xfe, 0, 0]);
  bad['skills/tool.exe'] = Buffer.from('MZ\x90\x00');
  delete bad['lib/client.js'];
  const problems = await universalProblems(await tree(bad), { ...FULL, kitManifest: JSON.stringify({ editor: { files: {} } }) });
  for (const pattern of [/must not set os or cpu/, /lib\/client\.js is missing/, /resources\/kit must not ship/, /offline-kit-runtime is an executable image/, /tool\.exe is an executable image/, /kit-manifest\.json differs/]) assert.ok(problems.some(problem => pattern.test(problem)), `${pattern}: ${problems.join(' | ')}`);
  assert.ok(isExecutableImage(Buffer.from([0x7f, 0x45, 0x4c, 0x46])) && !isExecutableImage(Buffer.from('#!/u')));
});

test('universalProblems: the public build carries exactly its two skills, complete, from the bundled public profile (261113)', async () => {
  const extra = await universalProblems(await tree({ ...base(), 'skills/shimo-xlsx/SKILL.md': 'x' }), FULL);
  assert.ok(extra.includes('skills/shimo-xlsx must not ship in the public build'), extra.join(' | '));
  const lost = base(); delete lost['skills/shimo-docx/SKILL.md'];
  const lostProblems = await universalProblems(await tree(lost), FULL);
  assert.ok(lostProblems.includes('skills/shimo-docx/SKILL.md is missing'), lostProblems.join(' | '));
  const remote = await universalProblems(await tree({ ...base(), 'skills-catalog.json': JSON.stringify({ ...JSON.parse(CATALOG), remoteCatalog: true }) }), FULL);
  assert.ok(remote.includes('the public build carries its skills: skills-catalog.json must be layer full with remoteCatalog false'), remote.join(' | '));
  const other = await universalProblems(await tree({ ...base(), 'download-sources.json': JSON.stringify({ v: 1, profile: 'other' }) }), FULL);
  assert.ok(other.includes('download-sources.json must name the public profile with bundled skills'), other.join(' | '));
});
