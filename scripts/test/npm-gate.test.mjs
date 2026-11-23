import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { REQUIRED } from '../universal-package.mjs';
import { gate, gateChecks, sri512 } from '../npm-gate.mjs';
import { sourcesFor } from '../../packages/dsh-adapter/src/host/fetch/download-sources.mjs';
import { obfuscateText } from '../../tooling/suite/obfuscate.mjs';

const repo = join(dirname(fileURLToPath(import.meta.url)), '../..');
const FILES = ['lib/host', 'lib/tools', 'lib/skills', 'lib/client.js', 'plugin.yml', 'locale', 'resources/kit-page', 'resources/editor', 'resources/host-fonts', 'skills', 'skills-catalog.json', 'capabilities.json', 'whats-new.json', 'CHANGELOG.md', 'kit-manifest.json', '.gitattributes', 'download-sources.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md'];
const manifest = (extra = {}) => ({ name: '@officesdk/officesdk-plugin-dsh', version: '0.1.1-test.1', license: 'UNLICENSED', files: FILES, engines: { node: '>=22.19.0', dsh: '>=0.2.1-rc.3 <0.3.0-0' },
  publishConfig: { registry: 'https://registry.npmjs.org/', access: 'public' }, repository: { type: 'git', url: 'git+https://github.com/officesdk/officesdk-demo.git' }, ...extra });

const KIT_MANIFEST = JSON.stringify({ sources: [{ id: 'release', kind: 'url', base: 'https://bucket.example/offline-kit/releases/0.1.3/' }], editor: { files: { 'catalog.json': { sha256: 'sha256-x' } } } });
/** The publisher's own domains, as the gate is configured with them (never committed: NPM_GATE_INTERNAL_HOSTS). */
const HOSTS = ['corp.example'];
const SKILL_MD = '---\nname: x\n---\n';
const skillCatalog = (names, extra) => JSON.stringify({ v: 1, ...extra, skills: Object.fromEntries(names.map(name => [name, [{ version: '0.0.1', files: { 'SKILL.md': { sha512: sri512(Buffer.from(SKILL_MD)) } } }]])) });

/** An unpacked public package that passes the gate (its two skills in the package, spec 261113-bundle-public-skills); `edit(root)` breaks it. */
async function fixture({ pkg = {}, publicBase, edit } = {}) {
  const root = join(await mkdtemp(join(tmpdir(), 'gate-')), 'package');
  for (const file of REQUIRED) { await mkdir(dirname(join(root, file)), { recursive: true }); await writeFile(join(root, file), ''); }
  await writeFile(join(root, 'package.json'), JSON.stringify(manifest(pkg)));
  const names = !publicBase ? ['shimo-pptx', 'shimo-docx'] : ['shimo-pptx', 'shimo-docx', 'shimo-xlsx'];
  for (const name of names) { await mkdir(join(root, 'skills', name), { recursive: true }); await writeFile(join(root, 'skills', name, 'SKILL.md'), SKILL_MD); }
  await writeFile(join(root, 'skills-catalog.json'), skillCatalog(names, names.length === 2 ? { layer: 'full', remoteCatalog: false } : { layer: 'light' }));
  await writeFile(join(root, 'kit-manifest.json'), KIT_MANIFEST);
  await writeFile(join(root, 'download-sources.json'), JSON.stringify(sourcesFor('public', { publicBase })));
  await writeFile(join(root, 'THIRD-PARTY-NOTICES.md'), '# Third-party notices\n\n## Bundled open-source packages\n\nNone.\n\n## office-kit\n\n- editor\n');
  await writeFile(join(root, 'LICENSE'), 'proprietary\n');
  await writeFile(join(root, 'lib/host/index.mjs'), 'export const x = 1;\n');
  await edit?.(root);
  return root;
}
const checks = async (options = {}, gateOptions = {}) => gateChecks(await fixture(options), { allow: [], kitManifest: KIT_MANIFEST, internalHosts: HOSTS, ...gateOptions });

test('a clean public package passes, strictly too', async () => {
  const strict = await checks({}, { target: 'npmjs' });
  assert.deepEqual(strict.errors, []);
  assert.deepEqual(strict.warnings, []);
});

test('files: only what package.json lists, never .npmrc, maps, tests or sources', async () => {
  const result = await checks({ edit: async root => { await writeFile(join(root, 'extra.js'), ''); await writeFile(join(root, '.npmrc'), ''); await mkdir(join(root, 'lib/host/test'), { recursive: true }); await writeFile(join(root, 'lib/host/test/a.test.mjs'), ''); await writeFile(join(root, 'lib/client.js.map'), ''); } });
  assert.deepEqual(result.errors.filter(e => e.startsWith('files:')).sort(), ['files: .npmrc must not ship', 'files: extra.js is not in package.json files', 'files: lib/client.js.map must not ship', 'files: lib/host/test/a.test.mjs must not ship']);
});

test('the editor resources never ship: they download on first use (261113-bundle-public-skills §6); their source maps are refused like any other', async () => {
  const bundled = await checks({ pkg: { files: [...FILES, 'resources/office-kit'] }, edit: async root => {
    await mkdir(join(root, 'resources/office-kit/pdf'), { recursive: true });
    await writeFile(join(root, 'resources/office-kit/catalog.json'), '{}');
    await writeFile(join(root, 'resources/office-kit/pdf/pdf-sdk.js.map'), '{"version":3}');
  } });
  assert.ok(bundled.errors.includes('package: resources/office-kit must not ship in the universal package'), bundled.errors.join('\n'));
  assert.ok(bundled.errors.includes('files: resources/office-kit/pdf/pdf-sdk.js.map must not ship'), bundled.errors.join('\n'));
});

test('a rollback republishes an older release as it was: its own kit-manifest.json stands; a normal publish of it does not', async () => {
  const older = { edit: root => writeFile(join(root, 'kit-manifest.json'), JSON.stringify({ officeKit: { version: '0.1.2' }, sources: [] })) };
  assert.deepEqual((await checks(older, { rollback: true })).errors, []);
  assert.ok((await checks(older)).errors.includes('package: kit-manifest.json differs from components/kit-manifest.json'));
});

test('the allow file names no editor resources any more, and every rule has a reason', async () => {
  const rules = JSON.parse(await readFile(join(repo, 'scripts/npm-gate-allow.json'), 'utf8'));
  assert.deepEqual(rules.filter(rule => rule.file.startsWith('resources/office-kit/')), []);
  for (const rule of rules) assert.ok(rule.reason.length > 20, rule.file);
});

test('public skills in the package: no mirror needed, but the catalog must carry them complete with no remote catalog (261113)', async () => {
  const strict = await checks({}, { target: 'npmjs' });
  assert.deepEqual(strict.errors, []);
  assert.ok(strict.notes.includes('sources: skills bundled in the package (shimo-pptx, shimo-docx)'), strict.notes.join('\n'));
  const light = await checks({ edit: root => writeFile(join(root, 'skills-catalog.json'), skillCatalog(['shimo-pptx', 'shimo-docx'], { layer: 'light' })) }, { target: 'npmjs' });
  assert.ok(light.errors.includes('sources: skills are bundled but skills-catalog.json is not layer full with remoteCatalog false'), light.errors.join('\n'));
  const xlsx = await checks({ edit: async root => { await mkdir(join(root, 'skills/shimo-xlsx'), { recursive: true }); await writeFile(join(root, 'skills/shimo-xlsx/SKILL.md'), SKILL_MD); } }, { target: 'npmjs' });
  assert.ok(xlsx.errors.includes('package: skills/shimo-xlsx must not ship in the public build'), xlsx.errors.join('\n'));
});

test('secrets: tokens, private keys and credentials in URLs fail', async () => {
  const result = await checks({ edit: root => writeFile(join(root, 'lib/host/index.mjs'), 'const a = "glpat-abcdefghijklmnopqrstuv"; const b = "https://deploy:s3cr3t@example.com/x"; //_authToken=abc\n-----BEGIN PRIVATE KEY-----\n') });
  for (const name of ['gitlab-token', 'url-credentials', 'npm-auth', 'private-key']) assert.ok(result.errors.some(e => e.includes(name)), name);
});

test('company addresses fail unless listed with a reason; listed ones pass', async () => {
  const edit = root => writeFile(join(root, 'lib/host/index.mjs'), 'const a = "https://git.corp.example/api/v4"; const b = "10.0.0.1";\n');
  const unlisted = await checks({ edit });
  assert.deepEqual(unlisted.errors.sort(), ['internal: lib/host/index.mjs names 10.0.0.1 (private-ip); remove it or list it with its reason in scripts/npm-gate-allow.json', 'internal: lib/host/index.mjs names git.corp.example (company-host); remove it or list it with its reason in scripts/npm-gate-allow.json']);
  const listed = await gateChecks(await fixture({ edit }), { allow: [{ file: 'lib/host/index.mjs', host: 'git.corp.example', reason: 'r' }, { file: 'lib/host/index.mjs', host: '10.0.0.1', reason: 'r' }], kitManifest: KIT_MANIFEST, internalHosts: HOSTS });
  assert.deepEqual(listed.errors, []);
  // office-kit's release bucket: its release directory is public, any other path of it is not.
  const bucket = await checks({ edit: root => writeFile(join(root, 'lib/host/index.mjs'), 'const a = "https://bucket.example/private/thing"; const b = "https://bucket.example/offline-kit/releases/0.2.0/x.tgz";\n') });
  assert.deepEqual(bucket.errors, ['internal: lib/host/index.mjs names bucket.example/private/thing (manifest-host); remove it or list it with its reason in scripts/npm-gate-allow.json']);
  // The allow file's `*` covers, in kit-manifest.json, only the hosts office-kit's manifest names as its sources and editor
  // registries: a path of its bucket outside the release directory, but never a configured company host.
  const star = { allow: [{ file: 'kit-manifest.json', host: '*', reason: 'r' }] };
  const bucketPath = { edit: root => { const kit = JSON.parse(KIT_MANIFEST); kit.notes = 'https://bucket.example/offline-kit/README'; return writeFile(join(root, 'kit-manifest.json'), JSON.stringify(kit)); } };
  assert.ok((await checks(bucketPath)).errors.some(error => error.startsWith('internal: kit-manifest.json names bucket.example/offline-kit/readme (manifest-host)')));
  assert.ok(!(await checks(bucketPath, star)).errors.some(error => error.startsWith('internal:')));
  const kitHosts = { edit: root => { const kit = JSON.parse(KIT_MANIFEST); kit.editor.npm = { registry: 'https://npm.corp.example/', package: 'e', version: '1' }; return writeFile(join(root, 'kit-manifest.json'), JSON.stringify(kit)); } };
  assert.ok((await checks(kitHosts)).errors.some(error => error.startsWith('internal: kit-manifest.json names npm.corp.example')));
  const vouched = (await checks(kitHosts, star)).errors;
  assert.ok(vouched.some(error => error.startsWith('internal: kit-manifest.json names npm.corp.example (company-host)')), `* never covers a company host:\n${vouched.join('\n')}`);
  const elsewhere = { edit: root => { const kit = JSON.parse(KIT_MANIFEST); kit.mirror = 'https://other.corp.example/x'; kit.lan = 'http://10.1.2.3/'; return writeFile(join(root, 'kit-manifest.json'), JSON.stringify(kit)); } };
  const notCovered = (await checks(elsewhere, star)).errors;
  assert.ok(notCovered.some(error => error.startsWith('internal: kit-manifest.json names other.corp.example')) && notCovered.some(error => error.startsWith('internal: kit-manifest.json names 10.1.2.3')), notCovered.join('\n'));
});

test('a build for another registry may name that registry in package.json and capabilities.json only, and never for npmjs', async () => {
  const LAN = 'http://192.168.7.20:4873/';
  const edit = root => writeFile(join(root, 'capabilities.json'), JSON.stringify({ registry: LAN, npm: { registry: LAN } }));
  const pkg = { publishConfig: { registry: LAN, access: 'public' } };
  const local = await checks({ pkg, edit }, { target: 'local' });
  assert.deepEqual(local.errors, []);
  assert.ok(local.notes.some(note => note.includes('own registry 192.168.7.20')));
  const elsewhere = await checks({ pkg, edit: async root => { await edit(root); await writeFile(join(root, 'lib/host/index.mjs'), `const r = "${LAN}";\n`); } }, { target: 'local' });
  assert.deepEqual(elsewhere.errors, ['internal: lib/host/index.mjs names 192.168.7.20 (private-ip); remove it or list it with its reason in scripts/npm-gate-allow.json'], 'the registry anywhere else is still an internal address');
  const other = await checks({ pkg, edit: root => writeFile(join(root, 'capabilities.json'), JSON.stringify({ registry: 'http://192.168.7.21:4873/' })) }, { target: 'local' });
  assert.equal(other.errors.length, 1, 'another LAN host is not the package\'s registry');
  const npmjs = await checks({ pkg, edit }, { target: 'npmjs' });
  assert.ok(npmjs.errors.some(error => error.startsWith('internal: capabilities.json names 192.168.7.20')), 'npmjs: no exception');
});

test('obfuscation does not hide a company address or a token from the gate (R1-D4)', async () => {
  const obfuscator = createRequire(join(repo, 'tooling', 'package.json'))('javascript-obfuscator');
  const code = obfuscateText(obfuscator, 'export function f(){ const api = "https://git.corp.example/api/v4"; const t = "glpat-abcdefghijklmnopqrstuv"; return [api, t].join(","); }\n', { sourceType: 'module' });
  assert.ok(!code.includes('function f(){ const api'), 'really obfuscated');
  const result = await checks({ edit: root => writeFile(join(root, 'lib/host/index.mjs'), code) });
  assert.ok(result.errors.some(e => e.includes('git.corp.example')));
  assert.ok(result.errors.some(e => e.includes('gitlab-token')));
});

test('no configured host list: a warning in CI, an error for npmjs; the committed gate names no organisation', async () => {
  const ci = await checks({}, { target: 'ci', internalHosts: [] });
  assert.deepEqual(ci.errors, []);
  assert.ok(ci.warnings.some(w => w.startsWith('internal: no company host list is configured')), ci.warnings.join('\n'));
  const strict = await checks({}, { target: 'npmjs', internalHosts: [] });
  assert.ok(strict.errors.some(e => e.startsWith('internal: no company host list is configured')), strict.errors.join('\n'));
  const { normaliseHosts, readInternalHosts } = await import('../npm-gate.mjs');
  assert.deepEqual(normaliseHosts(['https://Corp.Example/x', '*.b.example', '.c.example', 'nodot', '']), ['corp.example', 'b.example', 'c.example']);
  assert.deepEqual(await readInternalHosts({ env: { NPM_GATE_INTERNAL_HOSTS: 'a.example, b.example' } }), ['a.example', 'b.example']);
  const file = join(await mkdtemp(join(tmpdir(), 'hosts-')), 'h.json');
  await writeFile(file, JSON.stringify({ hosts: ['corp.example'] }));
  assert.deepEqual(await readInternalHosts({ env: {}, file }), ['corp.example']);
  assert.deepEqual(await readInternalHosts({ env: {}, file: '/nonexistent' }), []);
  assert.deepEqual(await readInternalHosts({ env: { NPM_GATE_INTERNAL_HOSTS: 'a.example' }, file }), ['corp.example'], 'an explicit file wins over the environment');
});

test('default sources: a profile other than public, or sources beyond the public defaults, are warnings in CI, errors for npmjs (R1-D6)', async () => {
  const other = { edit: async root => { const doc = sourcesFor('public'); doc.profile = 'other'; doc.profiles.other = { kit: ['release', 'vendor-copy'], editor: ['vendor-editor'], skills: { sources: ['elsewhere'], catalog: [] } }; await writeFile(join(root, 'download-sources.json'), JSON.stringify(doc)); } };
  const ci = await checks(other, { target: 'ci' });
  assert.ok(ci.warnings.some(w => w.includes('other profile is not for a public package')), ci.warnings.join('\n'));
  for (const text of ['tries vendor-copy for the kit', 'tries vendor-editor for the editor', 'skills come from elsewhere']) assert.ok(ci.warnings.some(w => w.includes(text)), text);
  assert.ok((await checks(other, { target: 'npmjs' })).errors.some(e => e.includes('tries vendor-editor for the editor')));
  assert.ok(!(await checks({}, { target: 'npmjs' })).errors.some(e => e.startsWith('sources:')), 'the public defaults (release for the Kit, npm for the editor) pass');
  // A public build still on the old mirror source (a drill) with no mirror configured.
  const strict = await checks({ publicBase: 'https://cdn.example/s', edit: async root => { const doc = sourcesFor('public', { publicBase: 'https://cdn.example/s' }); doc.profiles.public.skills.publicBase = ''; await writeFile(join(root, 'download-sources.json'), JSON.stringify(doc)); } }, { target: 'npmjs' });
  // (The package layer also refuses it: a public build carries its two skills complete.)
  assert.deepEqual(strict.errors.filter(error => error.startsWith('sources:')), ['sources: no public skill mirror (publicBase): skills cannot download for public users']);
  const company = await checks({ publicBase: 'https://git.corp.example/skills' }, { target: 'ci' });
  assert.match(company.errors.join('\n'), /publicBase https:\/\/git\.corp\.example\/skills is a company address/);
  const kit = await checks({ edit: root => writeFile(join(root, 'kit-manifest.json'), JSON.stringify({ sources: [{ id: 'm', base: 'https://npm.corp.example/x/' }] })) }, { allow: [{ file: 'kit-manifest.json', host: '*', reason: 'r' }] });
  assert.match(kit.errors.join('\n'), /kit-manifest.json source m is a company address/);
});

test('license: license field, LICENSE, notices, and bundled packages under allowed licenses only', async () => {
  const result = await checks({ pkg: { license: undefined }, edit: async root => { await writeFile(join(root, 'THIRD-PARTY-NOTICES.md'), '## Bundled open-source packages\n\n- left-pad 1.0.0 — MIT\n- copyleft 2.0.0 — GPL-3.0\n'); } });
  assert.deepEqual(result.errors, ['license: package.json has no license', 'license: copyleft 2.0.0 — GPL-3.0 is not under an allowed license (MIT, ISC, BSD-2-Clause, BSD-3-Clause, Apache-2.0, 0BSD, CC0-1.0, Unlicense, BlueOak-1.0.0)']);
});

test('metadata: install scripts, os/cpu, company email, non-public access and a company repository fail', async () => {
  const result = await checks({ pkg: { scripts: { postinstall: 'x' }, os: ['darwin'], maintainers: [{ name: 'a', email: 'a@corp.example' }], publishConfig: { access: 'restricted' }, repository: 'git+ssh://git@git.corp.example/officesdk/officesdk.git' } }, { target: 'npmjs', allow: [{ file: 'package.json', host: 'git.corp.example', reason: 'test' }, { file: 'package.json', host: 'corp.example', reason: 'test' }] });
  for (const message of ['scripts.postinstall', 'must not set os', 'company email', 'access must be public', 'repository git+ssh://git@git.corp.example/officesdk/officesdk.git is a company address']) assert.ok(result.errors.some(e => e.includes(message)), message);
});

test('version: tags match the version kind; same bytes are already published, other bytes never; next never comes first', async () => {
  assert.ok((await checks({}, { tag: 'latest' })).errors.some(e => e.includes('latest takes X.Y.Z')));
  assert.ok((await checks({ pkg: { version: '0.1.1' } }, { tag: 'next' })).errors.some(e => e.includes('next takes X.Y.Z-test.N')));
  const packument = integrity => ({ 'dist-tags': { latest: '0.1.0' }, versions: { '0.1.1-test.1': { dist: { integrity } } } });
  const same = await checks({}, { tag: 'next', registryInfo: { packument: packument('sha512-same') }, tgzIntegrity: 'sha512-same' });
  assert.equal(same.alreadyPublished, true);
  assert.deepEqual(same.errors, []);
  const other = await checks({}, { tag: 'next', registryInfo: { packument: packument('sha512-other') }, tgzIntegrity: 'sha512-mine' });
  assert.match(other.errors.join('\n'), /already published with other bytes/);
  const first = await checks({}, { tag: 'next', registryInfo: { packument: undefined } });
  assert.equal(first.firstPublish, true);
  assert.match(first.errors.join('\n'), /first publish must be a release/);
  assert.deepEqual((await checks({ pkg: { version: '0.1.1' } }, { tag: 'latest', registryInfo: { packument: undefined } })).errors, []);
});

test('version: preview stages publish under their own dist-tag and may come first (0.0.0-alpha.1)', async () => {
  const alpha = { pkg: { version: '0.0.0-alpha.1' } };
  assert.deepEqual((await checks(alpha, { tag: 'alpha', registryInfo: { packument: undefined } })).errors, [], 'the first publish of the alpha stage');
  assert.ok((await checks(alpha, { tag: 'latest' })).errors.some(e => e.includes('latest takes X.Y.Z, not 0.0.0-alpha.1')));
  assert.ok((await checks(alpha, { tag: 'beta' })).errors.some(e => e.includes('beta takes X.Y.Z-beta.N, not 0.0.0-alpha.1')));
  assert.ok((await checks(alpha, { tag: 'next' })).errors.some(e => e.includes('next takes X.Y.Z-test.N')));
  assert.ok((await checks(alpha, { tag: 'canary' })).errors.some(e => e.includes('canary is not a dist-tag')));
  assert.ok((await checks({ pkg: { version: '0.0.0-alpha' } }, { tag: 'alpha' })).errors.some(e => e.includes('is not X.Y.Z, X.Y.Z-alpha|beta|rc.N or X.Y.Z-test.N')));
  assert.deepEqual((await checks({ pkg: { version: '0.0.0-rc.2' } }, { tag: 'rc' })).errors, []);
});

test('the whole gate on a tarball: integrity, report, and the registry asked', async () => {
  const root = await fixture();
  const tgz = join(dirname(root), 'p.tgz');
  assert.equal(spawnSync('tar', ['-czf', tgz, '-C', dirname(root), 'package']).status, 0);
  const asked = [];
  const kitManifestFile = join(dirname(root), 'kit-manifest.json');
  await writeFile(kitManifestFile, KIT_MANIFEST);
  const report = await gate(tgz, { target: 'ci', registry: 'http://127.0.0.1:4873', tag: 'next', allowFile: '/nonexistent', kitManifestFile, fetchImpl: async url => { asked.push(url); return new Response(JSON.stringify({ 'dist-tags': { latest: '0.1.0' }, versions: {} }), { status: 200 }); } });
  assert.equal(report.ok, true, report.errors.join('\n'));
  assert.equal(report.integrity, sri512(await readFile(tgz)));
  assert.deepEqual(asked, ['http://127.0.0.1:4873/@officesdk%2Fofficesdk-plugin-dsh']);
});

test('every text file is scanned whatever its name, and every GitLab token kind is caught (R1-F7)', async () => {
  const result = await checks({ pkg: { files: [...FILES, 'bin'] }, edit: async root => {
    await mkdir(join(root, 'bin'), { recursive: true });
    await writeFile(join(root, 'bin/run.sh'), 'TOKEN=gldt-abcdefghijklmnopqrstuv\n');
    await writeFile(join(root, 'bin/config'), 'runner glrt-abcdefghijklmnopqrstuv\n');
    await writeFile(join(root, 'bin/blob.bin'), Buffer.concat([Buffer.from([0, 1, 2]), Buffer.from('glpat-abcdefghijklmnopqrstuv')]));
  } });
  assert.deepEqual(result.errors.filter(e => e.startsWith('secrets:')).map(e => e.split(' (')[0]).sort(), ['secrets: bin/config has gitlab-token', 'secrets: bin/run.sh has gitlab-token']);
});

/** office-kit 0.3's layout: the Kit from a GitHub release, the editor from npm (registry.npmjs.org). */
const GITHUB_KIT = JSON.stringify({ sources: [{ id: 'github', kind: 'url', base: 'https://github.com/officesdk/officesdk-core/releases/download/v0.3.0-rc.1/' }],
  editor: { files: { 'catalog.json': { sha256: 'sha256-x' } }, npm: { registry: 'https://registry.npmjs.org/', package: '@officesdk/officesdk-core-editor', version: '0.3.0-rc.1' },
    sources: [{ id: 'npm', kind: 'npm', registry: 'https://registry.npmjs.org/', package: '@officesdk/officesdk-core-editor', version: '0.3.0-rc.1' }] } });
const withKit = (text, more) => async root => { await writeFile(join(root, 'kit-manifest.json'), text); await more?.(root); };

test('a GitHub release source is not a private bucket: the package may name its own GitHub repository; a bucket still is', async () => {
  const pkg = { repository: { type: 'git', url: 'git+https://github.com/officesdk/officesdk-plugin-dsh.git' }, homepage: 'https://github.com/officesdk/officesdk-plugin-dsh#readme', bugs: { url: 'https://github.com/officesdk/officesdk-plugin-dsh/issues' } };
  const github = await checks({ pkg, edit: withKit(GITHUB_KIT, root => writeFile(join(root, 'lib/host/index.mjs'), 'const a = "https://github.com/officesdk/officesdk-core/issues";\n')) }, { target: 'npmjs', kitManifest: GITHUB_KIT });
  assert.deepEqual(github.errors, []);
  assert.deepEqual(github.warnings, []);
  // The same package against a storage bucket source: any other path of that host is still refused (regression).
  const bucket = await checks({ pkg, edit: root => writeFile(join(root, 'lib/host/index.mjs'), 'const a = "https://bucket.example/private/x";\n') }, { target: 'npmjs' });
  assert.deepEqual(bucket.errors, ['internal: lib/host/index.mjs names bucket.example/private/x (manifest-host); remove it or list it with its reason in scripts/npm-gate-allow.json']);
});

test('the publisher\'s internal registry, git host and storage bucket are refused wherever they appear, the allow file\'s * included', async () => {
  // Synthetic stand-ins for the three kinds (the real names are configuration, never committed).
  const hosts = ['corp.example', 'bucket.corp-cloud.example'];
  const star = [{ file: 'kit-manifest.json', host: '*', reason: 'r' }];
  const kit = JSON.parse(GITHUB_KIT);
  kit.editor.npm.registry = 'https://npm-registry.corp.example/';
  kit.editor.sources[0].registry = 'https://npm-registry.corp.example/';
  kit.sources.push({ id: 'bucket', kind: 'url', base: 'https://bucket.corp-cloud.example/offline-kit/releases/0.3.0/' });
  const text = JSON.stringify(kit);
  const result = await checks({ edit: withKit(text, root => writeFile(join(root, 'lib/host/index.mjs'), 'const g = "https://git.corp.example/group/repo.git";\n')) }, { target: 'npmjs', kitManifest: text, internalHosts: hosts, allow: star });
  for (const expected of [
    'sources: kit-manifest.json editor.npm.registry https://npm-registry.corp.example/ is a company address (company-host)',
    'sources: kit-manifest.json editor.sources[0].registry https://npm-registry.corp.example/ is a company address (company-host)',
    'sources: kit-manifest.json source bucket is a company address (company-host)',
    'internal: kit-manifest.json names npm-registry.corp.example (company-host); remove it or list it with its reason in scripts/npm-gate-allow.json',
    'internal: kit-manifest.json names bucket.corp-cloud.example (company-host); remove it or list it with its reason in scripts/npm-gate-allow.json',
    'internal: lib/host/index.mjs names git.corp.example (company-host); remove it or list it with its reason in scripts/npm-gate-allow.json',
  ]) assert.ok(result.errors.includes(expected), `${expected}\n--- got:\n${result.errors.join('\n')}`);
  // A private address as the editor registry is refused too.
  const lan = JSON.parse(GITHUB_KIT); lan.editor.npm.registry = 'http://10.1.2.3:4873/';
  const lanText = JSON.stringify(lan);
  assert.ok((await checks({ edit: withKit(lanText) }, { target: 'npmjs', kitManifest: lanText, allow: star })).errors.includes('sources: kit-manifest.json editor.npm.registry http://10.1.2.3:4873/ is a company address (private-ip)'));
});

