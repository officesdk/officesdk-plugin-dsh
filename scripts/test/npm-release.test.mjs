import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, realpath, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { REQUIRED } from '../universal-package.mjs';
import { sri512 } from '../npm-gate.mjs';
import { NO_AUTH_DIST_TAG_WARNING, NO_TOKEN_WARNING, SKIPPED, buildFor, distTag, distTagOidcCli, distTagOidcHint, distTagProblem, isNpmjs, npmPublish, oidcAvailable, oidcOutcome, publish, publishPackage, releaseNotes, rollback, rollbackVersion, sourceOf } from '../npm-release.mjs';
import { sourcesFor } from '../../packages/dsh-adapter/src/host/fetch/download-sources.mjs';

const REGISTRY = 'http://127.0.0.1:4873/';
const FILES = ['lib/host', 'lib/tools', 'lib/skills', 'lib/client.js', 'plugin.yml', 'locale', 'resources/kit-page', 'resources/editor', 'resources/host-fonts', 'skills', 'skills-catalog.json', 'capabilities.json', 'whats-new.json', 'CHANGELOG.md', 'kit-manifest.json', '.gitattributes', 'download-sources.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md'];

const DEFAULT_KIT = { version: '0.1.3', indexVersion: 3, readsIndexVersions: [1, 2, 3], downgradeFloor: '0.1.0' };
// The gate's repository manifest (the packaged kit-manifest.json must equal it outside a rollback).
const kitManifestText = kit => JSON.stringify({ officeKit: kit, sources: [{ id: 'release', kind: 'url', base: 'https://bucket.example/offline-kit/releases/0.1.3/' }],
  editor: { files: { 'catalog.json': { sha256: `sha256-${createHash('sha256').update('{}').digest('base64')}` } } } });
const kitManifestFile = join(await mkdtemp(join(tmpdir(), 'npm-release-kit-')), 'kit-manifest.json');
await writeFile(kitManifestFile, kitManifestText(DEFAULT_KIT));

/** A packed public package at `version` (office-kit `kit`), passing the gate. */
async function packed(version, { kit = DEFAULT_KIT, legacy = false } = {}) {
  const base = await mkdtemp(join(tmpdir(), 'npm-release-'));
  const root = join(base, 'package');
  for (const file of REQUIRED) { await mkdir(dirname(join(root, file)), { recursive: true }); await writeFile(join(root, file), ''); }
  const preview = { version, highlights: ['版本说明'], highlightsEn: ['Version notes'] };
  await writeFile(join(root, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version, ...!legacy && { officeSdkReleaseNotes: preview }, license: 'UNLICENSED', files: FILES, engines: { node: '>=22', dsh: '>=0.2.1-rc.3' }, publishConfig: { access: 'public' }, repository: { type: 'git', url: 'git+https://github.com/officesdk/officesdk-demo.git' } }));
  // The public build's two skills, in the package (spec 261113-bundle-public-skills).
  const skillMd = '---\nname: x\n---\n';
  for (const name of ['shimo-pptx', 'shimo-docx']) { await mkdir(join(root, 'skills', name), { recursive: true }); await writeFile(join(root, 'skills', name, 'SKILL.md'), skillMd); }
  const entry = [{ version: '0.0.1', files: { 'SKILL.md': { sha512: sri512(Buffer.from(skillMd)) } } }];
  await writeFile(join(root, 'skills-catalog.json'), JSON.stringify({ v: 1, layer: 'full', remoteCatalog: false, skills: { 'shimo-pptx': entry, 'shimo-docx': entry } }));
  await writeFile(join(root, 'kit-manifest.json'), kitManifestText(kit));
  await writeFile(join(root, 'download-sources.json'), JSON.stringify(sourcesFor('public')));
  await writeFile(join(root, 'THIRD-PARTY-NOTICES.md'), '## Bundled open-source packages\n\nNone.\n');
  await writeFile(join(root, 'LICENSE'), 'x\n');
  await writeFile(join(root, 'CHANGELOG.md'), `# Changelog\n\n## ${version.replace(/-test\.\d+$/, '')} — 2026-11-11\n\n### 更新要点\n- 版本说明\n### Highlights\n- Version notes\n`);
  await writeFile(join(root, 'whats-new.json'), JSON.stringify({ v: 1, entries: [{ ...preview, version: version.replace(/-test\.\d+$/, '') }] }));
  await writeFile(join(root, 'capabilities.json'), JSON.stringify({ v: 2, version, channel: /-test/.test(version) ? 'test' : 'stable', plugin: {}, platforms: [], install: {}, tagPattern: {}, requires: {}, officeKit: {}, common: [], products: {}, skills: [], highlights: [] }));
  const tgz = join(base, `p-${version}.tgz`);
  assert.equal(spawnSync('tar', ['-czf', tgz, '-C', base, 'package']).status, 0);
  return tgz;
}
/** A registry answering with `tags` and the tarballs in `files` (`version → tgz`). */
const registry = (tags, files = {}) => async url => {
  if (url.endsWith('.tgz')) return new Response(await readFile(files[url.split('/').at(-1).replace(/\.tgz$/, '')]), { status: 200 });
  const versions = {};
  for (const [version, file] of Object.entries(files)) versions[version] = { dist: { integrity: sri512(await readFile(file)), tarball: `${REGISTRY}@officesdk/officesdk-plugin-dsh/-/${version}.tgz` } };
  return new Response(JSON.stringify({ 'dist-tags': tags, versions }), { status: 200 });
};
const quiet = { log: () => {}, allowFile: '/nonexistent', kitManifestFile, internalHosts: ['corp.example'] };

test('no NPM_TOKEN publishes nothing and is not a failure', async () => {
  const calls = [];
  const logs = [];
  // oidc: false, so the test means the same on a GitHub Actions runner.
  const result = await publish({ tgz: await packed('0.1.1-test.1'), tag: 'next', registry: REGISTRY, token: '', oidc: false, publishImpl: async () => calls.push(1), ...quiet, log: m => logs.push(m) });
  assert.deepEqual([result.result, result.reason, calls.length], [SKIPPED, 'no-token', 0]);
  assert.deepEqual(logs, [NO_TOKEN_WARNING]);
});

test('the warning without a token says where it goes on GitHub, not GitLab\'s CI/CD variables', () => {
  assert.match(NO_TOKEN_WARNING, /^WARNING: NPM_TOKEN is not set/);
  assert.match(NO_TOKEN_WARNING, /npm-publish workflow/);
  assert.match(NO_TOKEN_WARNING, /trusted publishing/);
  assert.match(NO_TOKEN_WARNING, /Settings → Secrets and variables → Actions/);
  assert.match(NO_TOKEN_WARNING, /NPM_TOKEN=<token> node scripts\/npm-release\.mjs publish/);
  assert.match(NO_TOKEN_WARNING, /docs\/NPM\.md/);
  assert.doesNotMatch(NO_TOKEN_WARNING, /CI\/CD|Variables,|GitLab/);
});

test('OIDC: only a GitHub Actions job given id-token: write has it', () => {
  const job = { GITHUB_ACTIONS: 'true', ACTIONS_ID_TOKEN_REQUEST_URL: 'https://token.actions.example/x', ACTIONS_ID_TOKEN_REQUEST_TOKEN: 'r' };
  assert.equal(oidcAvailable(job), true);
  assert.equal(oidcAvailable({ ...job, GITHUB_ACTIONS: undefined }), false, 'not GitHub Actions');
  assert.equal(oidcAvailable({ ...job, GITHUB_ACTIONS: 'false' }), false);
  assert.equal(oidcAvailable({ ...job, ACTIONS_ID_TOKEN_REQUEST_URL: '' }), false, 'no id-token: write');
  assert.equal(oidcAvailable({ ...job, ACTIONS_ID_TOKEN_REQUEST_TOKEN: undefined }), false);
  assert.equal(oidcAvailable({}), false);
});

test('no NPM_TOKEN but OIDC: npmjs gets the package through trusted publishing; a drill registry still needs the token', async () => {
  const tgz = await packed('0.1.2');
  const calls = [];
  const npmjs404 = async () => new Response('', { status: 404 });
  const result = await publish({ tgz, tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, fetchImpl: npmjs404, publishImpl: async args => calls.push(args), ...quiet });
  assert.deepEqual([result.result, result.auth, result.registry], ['published', 'oidc', NPMJS_URL]);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].token, undefined, 'no token is handed to npm publish');
  const withToken = await publish({ tgz, tag: 'latest', registry: NPMJS_URL, token: 't', oidc: true, fetchImpl: npmjs404, publishImpl: async args => calls.push(args), ...quiet });
  assert.deepEqual([withToken.result, withToken.auth, calls[1].token], ['published', 'token', 't']);
  const drill = await publish({ tgz, tag: 'latest', registry: REGISTRY, token: '', oidc: true, fetchImpl: registry({}), publishImpl: async () => assert.fail('published'), ...quiet });
  assert.deepEqual([drill.result, drill.reason], [SKIPPED, 'no-token']);
});

test('next before any latest is skipped with a warning (npm would tag the test build latest)', async () => {
  const logs = [];
  const result = await publish({ tgz: await packed('0.1.1-test.1'), tag: 'next', registry: REGISTRY, token: 't', fetchImpl: async () => new Response('', { status: 404 }), publishImpl: async () => assert.fail('published'), ...quiet, log: m => logs.push(m) });
  assert.deepEqual([result.result, result.reason], [SKIPPED, 'no-latest']);
  assert.match(logs[0], /first publish must be a release/);
});

test('a gated package is published; the same bytes again are already done; other bytes under the number fail', async () => {
  const tgz = await packed('0.1.1-test.1');
  const calls = [];
  const result = await publish({ tgz, tag: 'next', registry: REGISTRY, token: 't', fetchImpl: registry({ latest: '0.1.0' }), publishImpl: async args => calls.push(args), ...quiet });
  assert.equal(result.result, 'published');
  assert.deepEqual(calls.map(c => [c.tag, c.registry]), [['next', REGISTRY]]);
  const again = await publish({ tgz, tag: 'next', registry: REGISTRY, token: 't', fetchImpl: registry({ latest: '0.1.0' }, { '0.1.1-test.1': tgz }), publishImpl: async () => assert.fail('published twice'), ...quiet });
  assert.equal(again.result, 'already-published');
  const other = await packed('0.1.1-test.1');
  await writeFile(join(dirname(other), 'package', 'LICENSE'), 'changed\n');
  spawnSync('tar', ['-czf', other, '-C', dirname(other), 'package']);
  await assert.rejects(publish({ tgz: other, tag: 'next', registry: REGISTRY, token: 't', fetchImpl: registry({ latest: '0.1.0' }, { '0.1.1-test.1': tgz }), publishImpl: async () => assert.fail('published'), ...quiet }), /already published with other bytes/);
});

test('the gate refuses: a test build under latest is never published', async () => {
  await assert.rejects(publish({ tgz: await packed('0.1.1-test.1'), tag: 'latest', registry: REGISTRY, token: 't', fetchImpl: registry({ latest: '0.1.0' }), publishImpl: async () => assert.fail('published'), ...quiet }), /latest takes X\.Y\.Z/);
});

test('rollback numbers above npm and 121, so a later release cannot collide (R2-D13)', () => {
  assert.equal(rollbackVersion(['0.1.0', '0.1.1', '0.1.2-test.4'], ['v0.1.1']), '0.1.2');
  assert.equal(rollbackVersion(['0.1.0', '0.1.1'], ['v0.1.3', 'v0.1.3-test.2']), '0.1.4', '121 released further than npm');
  // After a rollback claims v0.1.2 on 121, the next rollback and release_test (refuses a released X.Y.Z) move past it.
  assert.equal(rollbackVersion(['0.1.0', '0.1.1', '0.1.2'], ['v0.1.1', 'v0.1.2']), '0.1.3');
});

test('a preview rollback takes the next number of its stage on its highest X.Y.Z', () => {
  const versions = ['0.0.0-alpha.1', '0.0.0-alpha.2', '0.0.0-alpha.10'];
  assert.equal(rollbackVersion(versions, [], 'alpha'), '0.0.0-alpha.11', 'numbers, not text');
  assert.equal(rollbackVersion([...versions, '0.0.0-beta.1'], [], 'beta'), '0.0.0-beta.2');
  assert.throws(() => rollbackVersion([...versions, '0.0.0-beta.1'], [], 'alpha'), /below 0\.0\.0-beta\.1, which alpha users already take/, 'an alpha rollback below a beta never reaches alpha users');
  assert.equal(rollbackVersion(['0.0.0-alpha.4', '0.0.1-alpha.1'], [], 'alpha'), '0.0.1-alpha.2');
  assert.throws(() => rollbackVersion(versions, [], 'rc'), /no rc version/);
  assert.equal(rollbackVersion(['0.0.0-alpha.3', '0.0.0-beta.2', '0.0.1-alpha.1'], [], 'beta'), '0.0.0-beta.3', 'a newer X.Y.Z in a less mature stage does not block');
});

test('rollback republishes the target\'s files as the next patch (into a folder it creates, R1-F1); an office-kit below the floor is refused', async () => {
  const v010 = await packed('0.1.0');
  const v011 = await packed('0.1.1');
  const out = join(dirname(v011), 'not-yet', 'npm', 'rollback.tgz');
  const result = await rollback({ to: '0.1.0', registry: REGISTRY, out, fetchImpl: registry({ latest: '0.1.1' }, { '0.1.0': v010, '0.1.1': v011 }) });
  assert.equal(result.version, '0.1.2');
  const work = await mkdtemp(join(tmpdir(), 'rb-'));
  spawnSync('tar', ['-xzf', out, '-C', work]);
  assert.equal(JSON.parse(await readFile(join(work, 'package/package.json'), 'utf8')).version, '0.1.2');
  const preview = JSON.parse(await readFile(join(work, 'package/package.json'), 'utf8')).officeSdkReleaseNotes;
  assert.deepEqual(preview, { version: '0.1.2', highlights: ['本次更新包含重要的安全修复和改进'], highlightsEn: ['This update includes important security fixes and improvements.'] });
  assert.deepEqual(JSON.parse(await readFile(join(work, 'package/whats-new.json'), 'utf8')).entries[0], preview);
  const old = await packed('0.1.0', { kit: { version: '0.1.0', indexVersion: 1, readsIndexVersions: [1], downgradeFloor: '0.1.0' } });
  const newer = await packed('0.1.1', { kit: { version: '0.1.4', indexVersion: 4, readsIndexVersions: [3, 4], downgradeFloor: '0.1.3' } });
  await assert.rejects(rollback({ to: '0.1.0', registry: REGISTRY, out, fetchImpl: registry({ latest: '0.1.1' }, { '0.1.0': old, '0.1.1': newer }) }), /older office-kit/);
  await assert.rejects(rollback({ to: '0.0.9', registry: REGISTRY, out, fetchImpl: registry({ latest: '0.1.1' }, { '0.1.1': v011 }) }), /not a release/);
});

test('publish --rollback lets an older release (its own kit-manifest.json) through the gate; without it the gate refuses', async () => {
  const base = await mkdtemp(join(tmpdir(), 'npm-release-old-'));
  const root = join(base, 'package');
  for (const file of REQUIRED) { await mkdir(dirname(join(root, file)), { recursive: true }); await writeFile(join(root, file), ''); }
  const preview = { version: '0.1.2', highlights: ['修订'], highlightsEn: ['Revision'] };
  await writeFile(join(root, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version: '0.1.2', officeSdkReleaseNotes: preview, license: 'UNLICENSED', files: FILES, engines: { node: '>=22', dsh: '>=0.2.1-rc.3' }, publishConfig: { access: 'public' }, repository: { type: 'git', url: 'git+https://github.com/officesdk/officesdk-demo.git' } }));
  // An older release: its two public skills in the package, and an office-kit other than the repository's.
  for (const name of ['shimo-pptx', 'shimo-docx']) { await mkdir(join(root, 'skills', name), { recursive: true }); await writeFile(join(root, 'skills', name, 'SKILL.md'), ''); }
  const entry = [{ version: '0.0.1', files: { 'SKILL.md': { sha512: sri512(Buffer.from('')) } } }];
  await writeFile(join(root, 'skills-catalog.json'), JSON.stringify({ v: 1, layer: 'full', remoteCatalog: false, skills: { 'shimo-pptx': entry, 'shimo-docx': entry } }));
  await writeFile(join(root, 'kit-manifest.json'), JSON.stringify({ officeKit: DEFAULT_KIT, sources: [{ id: 'release', kind: 'url', base: 'https://bucket.example/offline-kit/releases/0.1.3/' }] }));
  await writeFile(join(root, 'download-sources.json'), JSON.stringify(sourcesFor('public')));
  await writeFile(join(root, 'THIRD-PARTY-NOTICES.md'), '## Bundled open-source packages\n\nNone.\n');
  await writeFile(join(root, 'LICENSE'), 'x\n');
  await writeFile(join(root, 'CHANGELOG.md'), '# Changelog\n\n## 0.1.2 — 2026-11-12\n\n### 更新要点\n- 修订\n### Highlights\n- Revision\n');
  await writeFile(join(root, 'whats-new.json'), JSON.stringify({ v: 1, entries: [preview] }));
  await writeFile(join(root, 'capabilities.json'), JSON.stringify({ v: 2, version: '0.1.2', channel: 'stable', plugin: {}, platforms: [], install: {}, tagPattern: {}, requires: {}, officeKit: {}, common: [], products: {}, skills: [], highlights: [] }));
  const tgz = join(base, 'rollback.tgz');
  assert.equal(spawnSync('tar', ['-czf', tgz, '-C', base, 'package']).status, 0);
  const args = { tgz, tag: 'latest', registry: REGISTRY, token: 't', fetchImpl: registry({ latest: '0.1.1' }), ...quiet };
  await assert.rejects(publish({ ...args, publishImpl: async () => assert.fail('published') }), /kit-manifest\.json differs from components\/kit-manifest\.json/);
  const calls = [];
  assert.equal((await publish({ ...args, rollback: true, publishImpl: async call => calls.push(call) })).result, 'published');
  assert.equal(calls.length, 1);
});

// publish --version (spec 261113-fix-npmpublish-public-tgz): npmjs takes the public build, any other registry the build made for it.
const NPMJS_URL = 'https://registry.npmjs.org/';
const FILE = 'officesdk-officesdk-plugin-dsh-0.0.0-alpha.3.tgz';

test('npmjs (with or without the slash) takes the public build artifacts/npm/, another registry the build made for it', () => {
  for (const registry of [NPMJS_URL, 'https://registry.npmjs.org']) {
    assert.equal(isNpmjs(registry), true, registry);
    assert.equal(buildFor({ version: '0.0.0-alpha.3', registry }), join('artifacts', 'npm', FILE));
  }
  for (const registry of [REGISTRY, 'http://192.168.1.20:4873', 'https://npm.example.com/', 'https://registry.npmjs.org.example.com/']) {
    assert.equal(isNpmjs(registry), false, registry);
    assert.equal(buildFor({ version: '0.0.0-alpha.3', registry }), join('artifacts', FILE));
  }
  assert.equal(buildFor({ version: '1.2.3', registry: NPMJS_URL, artifacts: '/x' }), '/x/npm/officesdk-officesdk-plugin-dsh-1.2.3.tgz');
  assert.throws(() => buildFor({ registry: NPMJS_URL }), /version/);
});

test('publish takes --package as given or --version\'s build for the registry, which must exist; not both, not neither', () => {
  const all = () => true;
  assert.equal(publishPackage({ package: 'x.tgz' }, NPMJS_URL, all), 'x.tgz');
  assert.equal(publishPackage({ version: '0.0.0-alpha.3' }, NPMJS_URL, all), join('artifacts', 'npm', FILE));
  assert.equal(publishPackage({ version: '0.0.0-alpha.3' }, REGISTRY, all), join('artifacts', FILE));
  assert.throws(() => publishPackage({ package: 'x.tgz', version: '1.0.0' }, NPMJS_URL, all), /not both/);
  assert.throws(() => publishPackage({}, NPMJS_URL, all), /--package <tgz> or --version/);
  assert.throws(() => publishPackage({ version: '0.0.0-alpha.3' }, NPMJS_URL, () => false), /public npm build .* is missing/);
});

test('dry run: no token needed, the gate runs for the registry, nothing is published', async () => {
  const calls = [];
  const tgz = await packed('0.1.2');
  const result = await publish({ tgz, tag: 'latest', registry: REGISTRY, token: '', dryRun: true, fetchImpl: registry({}), publishImpl: async () => calls.push(1), ...quiet });
  assert.deepEqual([result.result, result.target, result.tgz, calls.length], ['dry-run', 'local', tgz, 0]);
  // The strict npmjs gate (a fake npmjs answering 404) reads the same package: the target follows the registry.
  const strict = await publish({ tgz, tag: 'latest', registry: NPMJS_URL, token: '', dryRun: true, fetchImpl: async () => new Response('', { status: 404 }), publishImpl: async () => calls.push(1), ...quiet }).catch(error => error);
  assert.equal(calls.length, 0);
  assert.ok(!(strict instanceof Error), String(strict.message));
  assert.deepEqual([strict.result, strict.target, strict.registry], ['dry-run', 'npmjs', NPMJS_URL]);
});

test('the CLI picks the build per registry and says so; a dry run writes no npm-publish.json', async () => {
  const script = fileURLToPath(new URL('../npm-release.mjs', import.meta.url));
  const cwd = await realpath(await mkdtemp(join(tmpdir(), 'npm-release-cli-')));
  await mkdir(join(cwd, 'artifacts', 'npm'), { recursive: true });
  const run = (registry, env = {}) => spawnSync(process.execPath, [script, 'publish', '--version', '0.0.0-alpha.3', '--tag', 'alpha', '--registry', registry, '--dry-run'], { cwd, encoding: 'utf8', env: { ...process.env, NPM_TOKEN: '', NPM_REGISTRY: '', ...env } });
  // Neither build exists: the error names the one the registry needs.
  let result = run(NPMJS_URL);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /artifacts\/npm\/officesdk-officesdk-plugin-dsh-0\.0\.0-alpha\.3\.tgz does not exist: the public npm build/);
  result = run('http://127.0.0.1:1/');
  assert.match(result.stderr, /artifacts\/officesdk-officesdk-plugin-dsh-0\.0\.0-alpha\.3\.tgz does not exist: the build for this registry/);
  // Both exist (not tarballs: the selection line comes before the gate reads them).
  await writeFile(join(cwd, 'artifacts', FILE), 'internal');
  await writeFile(join(cwd, 'artifacts', 'npm', FILE), 'public');
  result = run('http://127.0.0.1:1/');
  assert.ok(result.stderr.includes(`npm-release: http://127.0.0.1:1/ takes ${resolve(cwd, 'artifacts', FILE)} (gate local), dry run`), result.stderr);
  result = run('', { NPM_REGISTRY: '' });
  assert.ok(result.stderr.includes(`npm-release: ${NPMJS_URL} takes ${resolve(cwd, 'artifacts', 'npm', FILE)} (gate npmjs), dry run`), result.stderr);
  assert.equal(spawnSync('test', ['-e', join(cwd, 'artifacts', 'npm-publish.json')]).status, 1, 'a dry run writes no record');
});

test('npm publish: the throwaway userconfig holds the token, or without one nothing (the npm CLI does OIDC); it is removed', async () => {
  const seen = [];
  const spawn = (command, args, options) => {
    const userconfig = args[args.indexOf('--userconfig') + 1];
    seen.push({ userconfig, text: spawnSync('cat', [userconfig], { encoding: 'utf8' }).stdout, env: options.env.npm_config_userconfig, args });
    return { status: 0, stdout: '+ @officesdk/officesdk-plugin-dsh@0.1.2\n' };
  };
  assert.equal(await npmPublish({ tgz: 'x.tgz', tag: 'latest', registry: NPMJS_URL, token: 'secret-t', spawn }), '+ @officesdk/officesdk-plugin-dsh@0.1.2');
  await npmPublish({ tgz: 'x.tgz', tag: 'latest', registry: NPMJS_URL, token: undefined, spawn });
  assert.equal(seen[0].text, '//registry.npmjs.org/:_authToken=secret-t\n');
  assert.equal(seen[1].text, '', 'no _authToken without a token');
  for (const { userconfig, env, args } of seen) {
    assert.equal(env, userconfig, 'npm reads only this userconfig, not ~/.npmrc');
    assert.equal(args.includes('--provenance'), false, 'npm adds provenance itself (public repositories only)');
    assert.equal(spawnSync('test', ['-e', userconfig]).status, 1, 'the userconfig is removed');
  }
});

test('npm publish failing: the token never shows; without a token the error points at the trusted publisher', async () => {
  const failing = () => ({ status: 1, stderr: 'npm error code E404\nnpm error 404 Not Found - PUT https://registry.npmjs.org/x - secret-t\n' });
  await assert.rejects(npmPublish({ tgz: 'x.tgz', tag: 'rc', registry: NPMJS_URL, token: 'secret-t', spawn: failing }), error => !error.message.includes('secret-t') && /\*\*\*/.test(error.message) && !/trusted publisher/.test(error.message));
  await assert.rejects(npmPublish({ tgz: 'x.tgz', tag: 'rc', registry: NPMJS_URL, token: undefined, spawn: () => ({ status: 1, stderr: 'npm error code ENEEDAUTH\n' }) }), /ENEEDAUTH.*trusted publisher on npmjs\.com — workflow filename npm-publish\.yml, Environment, Allowed actions/);
});

test('the publish record names its source: GitHub Actions commit, tree and run; else this checkout; git failing leaves fields out', () => {
  const asked = [];
  const git = args => { asked.push(args.join(' ')); return { 'rev-parse HEAD': 'local-commit', 'rev-parse abc123^{tree}': 'tree-of-abc', 'rev-parse local-commit^{tree}': 'tree-of-local' }[args.join(' ')]; };
  assert.deepEqual(sourceOf({ GITHUB_ACTIONS: 'true', GITHUB_SHA: 'abc123', GITHUB_SERVER_URL: 'https://github.com', GITHUB_REPOSITORY: 'officesdk/officesdk-plugin-dsh', GITHUB_RUN_ID: '42' }, git),
    { commit: 'abc123', tree: 'tree-of-abc', run: 'https://github.com/officesdk/officesdk-plugin-dsh/actions/runs/42' });
  assert.deepEqual(asked, ['rev-parse abc123^{tree}'], 'on GitHub Actions the commit is GITHUB_SHA, not HEAD');
  assert.deepEqual(sourceOf({ CI_COMMIT_SHA: 'gitlab', CI_PIPELINE_IID: '7' }, git), { commit: 'local-commit', tree: 'tree-of-local', run: undefined }, 'GitLab variables are not read');
  assert.deepEqual(sourceOf({}, () => { throw new Error('git: not found'); }), { commit: undefined, tree: undefined, run: undefined });
  // This checkout (when the tests run from a git checkout): a commit and its tree.
  if (spawnSync('git', ['rev-parse', 'HEAD'], { cwd: dirname(fileURLToPath(import.meta.url)) }).status === 0) {
    const here = sourceOf({});
    assert.match(here.commit ?? '', /^[0-9a-f]{40}$/);
    assert.match(here.tree ?? '', /^[0-9a-f]{40}$/);
  }
});

test('npm publish gets the package as an absolute path (npm 11 reads artifacts/x.tgz as a GitHub repository)', async () => {
  const calls = [];
  await npmPublish({ tgz: join('artifacts', 'npm', FILE), tag: 'alpha', registry: REGISTRY, token: 't', spawn: (command, args) => { calls.push(args); return { status: 0, stdout: '' }; } });
  assert.equal(calls[0][1], resolve('artifacts', 'npm', FILE));
  assert.deepEqual(calls[0].slice(2, 6), ['--tag', 'alpha', '--registry', REGISTRY]);
});

// npm-release.mjs dist-tag and release-notes (round 11: the npm-dist-tag-release workflow).
const PACKAGE = '@officesdk/officesdk-plugin-dsh';

test('dist-tag rules: next only for test builds; alpha/beta/rc take their stage or a more mature one; latest a release or a preview', () => {
  const ok = [['0.0.1-rc.1', 'latest'], ['0.0.0-alpha.3', 'latest'], ['1.0.0', 'latest'], ['0.0.1-rc.1', 'rc'], ['1.0.0', 'rc'], ['0.0.1-rc.1', 'alpha'], ['0.0.1-beta.2', 'beta'], ['0.1.1-test.4', 'next']];
  for (const [version, tag] of ok) assert.equal(distTagProblem(version, tag), undefined, `${tag} ← ${version}`);
  const refused = [
    ['0.0.0-alpha.3', 'rc', /rc takes rc versions or more mature ones, not the alpha version/],
    ['0.0.1-beta.1', 'rc', /not the beta version/],
    ['0.0.1-alpha.1', 'beta', /not the alpha version/],
    ['1.0.0', 'next', /next takes test builds/],
    ['0.0.1-rc.1', 'next', /next takes test builds/],
    ['0.1.1-test.4', 'latest', /test build: it goes under next only/],
    ['0.1.1-test.4', 'alpha', /test build/],
    ['0.0.1-rc.1', 'stable', /--tag takes alpha, beta, rc, next, latest, not "stable"/],
    ['0.0.1-rc.1', '', /--tag takes/],
    ['0.0.1-win32', 'latest', /is not X\.Y\.Z/],
    ['v0.0.1', 'latest', /is not X\.Y\.Z/],
    [undefined, 'latest', /"" is not X\.Y\.Z/],
  ];
  for (const [version, tag, pattern] of refused) assert.match(distTagProblem(version, tag) ?? '', pattern, `${tag} ← ${version}`);
});

test('the npm CLI moves dist-tags through trusted publishing from 11.21.0 on (12.2.0 on npm 12)', () => {
  for (const version of ['11.21.0', '11.21.3', '11.30.1', '12.2.0', '12.10.0', '13.0.0', '11.21.3\n']) assert.equal(distTagOidcCli(version), true, version);
  for (const version of ['11.20.0', '11.19.1', '11.5.1', '12.0.0', '12.1.9', '10.9.4', '', 'npm', undefined]) assert.equal(distTagOidcCli(version), false, String(version));
});

test('the OIDC outcome comes from npm\'s verbose oidc lines', () => {
  assert.equal(oidcOutcome(['Successfully retrieved and set token']), 'exchanged');
  assert.equal(oidcOutcome(['Failed token exchange request with body message: No trusted publisher configuration found']), 'failed: Failed token exchange request with body message: No trusted publisher configuration found');
  assert.equal(oidcOutcome(['Failure with message: connect ECONNREFUSED']), 'failed: Failure with message: connect ECONNREFUSED');
  assert.equal(oidcOutcome(['Skipped because incorrect permissions for id-token within GitHub workflow']), 'not-attempted');
  assert.equal(oidcOutcome([]), 'not-attempted');
});

/** A registry with `versions` published and the dist-tags `tags` (a live object: a stub `npm dist-tag add` may change it). */
function tagRegistry(versions, tags) {
  const calls = [];
  const fetchImpl = async url => {
    calls.push(url);
    if (url.endsWith('/dist-tags')) return new Response(JSON.stringify(tags), { status: 200 });
    return new Response(JSON.stringify({ 'dist-tags': { ...tags }, versions: Object.fromEntries(versions.map(v => [v, { dist: { integrity: 'sha512-x' } }])) }), { status: 200 });
  };
  return { fetchImpl, calls };
}

/** A stub npm: `--version` answers `npmVersion`; dist-tag ls / add as `behaviour` says; every call recorded with its userconfig. */
function npmStub({ npmVersion = '11.21.3', tags, oidcLine = 'Successfully retrieved and set token', addStatus = 0, addStderr = '', move = true } = {}) {
  const calls = [];
  const spawn = (command, args, options) => {
    assert.equal(command, 'npm');
    if (args[0] === '--version') { calls.push({ args }); return { status: 0, stdout: `${npmVersion}\n` }; }
    const userconfig = args[args.indexOf('--userconfig') + 1];
    calls.push({ args, userconfig, text: spawnSync('cat', [userconfig], { encoding: 'utf8' }).stdout, env: options.env.npm_config_userconfig });
    const stderr = `npm verbose cli /usr/bin/node npm\nnpm http fetch GET 200 https://token.actions.example/idtoken?audience=npm 5ms\n${oidcLine ? `npm verbose oidc ${oidcLine}\n` : ''}`;
    if (args[1] === 'ls') return { status: 0, stdout: Object.entries(tags).map(([k, v]) => `${k}: ${v}`).sort().join('\n') + '\n', stderr };
    if (args[1] === 'add') {
      if (addStatus) return { status: addStatus, stdout: '', stderr: stderr + addStderr };
      const [spec, tag] = args.slice(2, 4);
      if (move) tags[tag] = spec.slice(spec.lastIndexOf('@') + 1);
      return { status: 0, stdout: `+${tag}: ${spec}\n`, stderr };
    }
    throw new Error(`unexpected npm ${args.join(' ')}`);
  };
  return { spawn, calls };
}
const NOW = () => ({ 'latest': '0.0.0-alpha.3', alpha: '0.0.0-alpha.3', rc: '0.0.1-rc.1' });
const VERSIONS = ['0.0.0-alpha.3', '0.0.1-rc.1'];

test('dist-tag dry run (OIDC): lists the tags with npm dist-tag ls, reports the exchange, changes nothing', async () => {
  const tags = NOW();
  const { fetchImpl } = tagRegistry(VERSIONS, tags);
  const npm = npmStub({ tags });
  const result = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, dryRun: true, fetchImpl, spawn: npm.spawn, log: () => {} });
  assert.deepEqual(result, { result: 'dry-run', package: `${PACKAGE}@0.0.1-rc.1`, version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, from: '0.0.0-alpha.3', auth: 'oidc', change: 'move', oidc: 'exchanged', tags: NOW() });
  assert.deepEqual(npm.calls.map(call => call.args.slice(0, 3)), [['--version'], ['dist-tag', 'ls', PACKAGE]]);
  const ls = npm.calls[1];
  assert.deepEqual(ls.args.slice(3), ['--registry', NPMJS_URL, '--userconfig', ls.userconfig, '--loglevel', 'verbose']);
  assert.equal(ls.text, '', 'no _authToken: the npm CLI trades the OIDC token itself');
  assert.equal(ls.env, ls.userconfig, 'npm reads only this userconfig, not ~/.npmrc');
  assert.equal(spawnSync('test', ['-e', ls.userconfig]).status, 1, 'the userconfig is removed');
  assert.deepEqual(tags, NOW(), 'nothing moved');
  const failed = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, dryRun: true, fetchImpl, spawn: npmStub({ tags, oidcLine: 'Failed token exchange request with body message: No trusted publisher configuration found' }).spawn, log: () => {} });
  assert.equal(failed.oidc, 'failed: Failed token exchange request with body message: No trusted publisher configuration found');
  const same = await distTag({ version: '0.0.1-rc.1', tag: 'rc', registry: NPMJS_URL, token: '', oidc: true, dryRun: true, fetchImpl, spawn: npmStub({ tags }).spawn, log: () => {} });
  assert.deepEqual([same.result, same.change, same.from], ['dry-run', 'none', '0.0.1-rc.1']);
});

test('dist-tag dry run without any credential still lists the tags (a local check); a move without one is skipped with a warning', async () => {
  const tags = NOW();
  const { fetchImpl } = tagRegistry(VERSIONS, tags);
  const npm = npmStub({ tags });
  const dry = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: false, dryRun: true, fetchImpl, spawn: npm.spawn, log: () => {} });
  assert.deepEqual([dry.result, dry.auth, dry.change, dry.from, 'oidc' in dry], ['dry-run', 'none', 'move', '0.0.0-alpha.3', false]);
  assert.deepEqual(npm.calls.map(call => call.args[0] === '--version' ? 'version' : call.args[1]), ['ls'], 'no npm version requirement without OIDC');
  const logs = [];
  const skipped = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: false, fetchImpl, spawn: () => assert.fail('npm called'), log: m => logs.push(m) });
  assert.deepEqual([skipped.result, skipped.reason], [SKIPPED, 'no-token']);
  assert.deepEqual(logs, [NO_AUTH_DIST_TAG_WARNING]);
  // Trusted publishing is npmjs's: a drill registry with OIDC but no token is skipped too.
  const drill = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: REGISTRY, token: '', oidc: true, fetchImpl, spawn: () => assert.fail('npm called'), log: () => {} });
  assert.deepEqual([drill.result, drill.reason], [SKIPPED, 'no-token']);
});

test('the warning without a credential points at the npm-publish workflow\'s dist-tag mode and its trusted publisher, not GitLab', () => {
  assert.match(NO_AUTH_DIST_TAG_WARNING, /^WARNING: NPM_TOKEN is not set/);
  assert.match(NO_AUTH_DIST_TAG_WARNING, /no dist-tag is changed/);
  assert.match(NO_AUTH_DIST_TAG_WARNING, /npm-publish workflow in mode dist-tag/);
  assert.match(NO_AUTH_DIST_TAG_WARNING, /trusted publisher for npm-publish\.yml with "Allow npm dist-tag"/);
  assert.match(NO_AUTH_DIST_TAG_WARNING, /npm dist-tag add @officesdk\/officesdk-plugin-dsh@<version> <tag>/);
  assert.match(NO_AUTH_DIST_TAG_WARNING, /docs\/NPM\.md/);
  assert.doesNotMatch(NO_AUTH_DIST_TAG_WARNING, /CI\/CD|GitLab/);
});

test('dist-tag move (OIDC): npm dist-tag add <package>@<version> <tag>, then the registry is read back', async () => {
  const tags = NOW();
  const { fetchImpl, calls: reads } = tagRegistry(VERSIONS, tags);
  const npm = npmStub({ tags });
  const result = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, fetchImpl, spawn: npm.spawn, log: () => {}, sleep: async () => {} });
  assert.deepEqual(result, { result: 'moved', package: `${PACKAGE}@0.0.1-rc.1`, version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, from: '0.0.0-alpha.3', auth: 'oidc' });
  const add = npm.calls.find(call => call.args[1] === 'add');
  assert.deepEqual(add.args.slice(0, 4), ['dist-tag', 'add', `${PACKAGE}@0.0.1-rc.1`, 'latest']);
  assert.equal(add.text, '', 'no _authToken with OIDC');
  assert.deepEqual(reads.filter(url => url.endsWith('/dist-tags')), [`${NPMJS_URL}-/package/@officesdk%2Fofficesdk-plugin-dsh/dist-tags`, `${NPMJS_URL}-/package/@officesdk%2Fofficesdk-plugin-dsh/dist-tags`], 'before and after');
  assert.equal(tags.latest, '0.0.1-rc.1');
  // Again: the tag is there already, npm is not called.
  const again = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, fetchImpl, spawn: () => assert.fail('npm called'), log: () => {} });
  assert.deepEqual([again.result, again.from], ['unchanged', '0.0.1-rc.1']);
});

test('dist-tag move: the registry not showing the move fails after its retries; npm failing names the trusted publisher', async () => {
  const tags = NOW();
  const { fetchImpl } = tagRegistry(VERSIONS, tags);
  const waits = [];
  await assert.rejects(distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, fetchImpl, spawn: npmStub({ tags, move: false }).spawn, log: () => {}, sleep: async ms => waits.push(ms) }), /reported success, but https:\/\/registry\.npmjs\.org\/ still has latest at 0\.0\.0-alpha\.3, not 0\.0\.1-rc\.1/);
  assert.deepEqual(waits, [2000, 2000, 2000, 2000], 'five reads, two seconds apart');
  const refused = npmStub({ tags, oidcLine: 'Failed token exchange request with body message: No trusted publisher configuration found', addStatus: 1, addStderr: 'npm error code E401\nnpm error 401 Unauthorized - PUT https://registry.npmjs.org/-/package/x/dist-tags/latest\nnpm error A complete log of this run can be found in: /home/runner/.npm/_logs/x.log\n' });
  await assert.rejects(distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, fetchImpl, spawn: refused.spawn, log: () => {} }), error =>
    /npm dist-tag add failed: oidc: Failed token exchange request with body message: No trusted publisher configuration found code E401 401 Unauthorized/.test(error.message)
    && /did not authenticate: check the trusted publisher on npmjs\.com \(workflow filename npm-publish\.yml, repository officesdk\/officesdk-plugin-dsh, Environment empty unless the publish and dist-tag jobs both run in it\) and npm 11\.21\.0 or later/.test(error.message)
    && !/complete log/.test(error.message) && !/token\.actions\.example/.test(error.message));
  assert.equal(tags.latest, '0.0.0-alpha.3');
});

test('the hint for a failed move through trusted publishing follows npm\'s error code', () => {
  assert.match(distTagOidcHint(['code EOTP', 'This operation requires a one-time password']), /one-time password, which npm trusted publishing does not need: check the package's Publishing access/);
  assert.match(distTagOidcHint(['code E403', '403 Forbidden - PUT https://registry.npmjs.org/-/package/x/dist-tags/latest']), /trusted publisher for npm-publish\.yml may not move dist-tags: tick "Allow npm dist-tag"/);
  assert.match(distTagOidcHint(['code ENEEDAUTH']), /did not authenticate: check the trusted publisher on npmjs\.com \(workflow filename npm-publish\.yml/);
  assert.match(distTagOidcHint(['code E401', '401 Unauthorized']), /did not authenticate/);
  assert.doesNotMatch(distTagOidcHint([]), /Expired|npm-dist-tag-release/);
});

test('dist-tag (OIDC) refuses an npm CLI older than 11.21.0 before calling npm dist-tag (it would not trade the token)', async () => {
  const tags = NOW();
  const { fetchImpl } = tagRegistry(VERSIONS, tags);
  for (const dryRun of [true, false]) {
    const npm = npmStub({ tags, npmVersion: '11.19.1' });
    await assert.rejects(distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: '', oidc: true, dryRun, fetchImpl, spawn: npm.spawn, log: () => {} }), /npm 11\.19\.1 does not move dist-tags through npm trusted publishing: it needs npm 11\.21\.0 or later/);
    assert.deepEqual(npm.calls.map(call => call.args[0]), ['--version']);
  }
});

test('dist-tag with NPM_TOKEN: the userconfig holds it, an error never shows it, no npm version requirement', async () => {
  const tags = NOW();
  const { fetchImpl } = tagRegistry(VERSIONS, tags);
  const npm = npmStub({ tags, npmVersion: '10.9.4' });
  const result = await distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: 'secret-t', oidc: true, fetchImpl, spawn: npm.spawn, log: () => {}, sleep: async () => {} });
  assert.deepEqual([result.result, result.auth], ['moved', 'token']);
  const add = npm.calls.find(call => call.args[1] === 'add');
  assert.equal(add.text, '//registry.npmjs.org/:_authToken=secret-t\n');
  assert.equal(npm.calls.some(call => call.args[0] === '--version'), false);
  const failing = npmStub({ tags: NOW(), oidcLine: '', addStatus: 1, addStderr: 'npm error code E403\nnpm error 403 Forbidden - secret-t is not allowed\n' });
  await assert.rejects(distTag({ version: '0.0.1-rc.1', tag: 'latest', registry: NPMJS_URL, token: 'secret-t', fetchImpl: tagRegistry(VERSIONS, NOW()).fetchImpl, spawn: failing.spawn, log: () => {} }), error => !error.message.includes('secret-t') && /\*\*\* is not allowed/.test(error.message) && !/trusted publisher/.test(error.message));
});

test('dist-tag refuses a version not on the registry, a move back, and a version the tag may not take', async () => {
  const tags = { latest: '0.0.1-rc.1', alpha: '0.0.0-alpha.3', rc: '0.0.1-rc.1' };
  const { fetchImpl } = tagRegistry(VERSIONS, tags);
  const never = () => assert.fail('npm called');
  await assert.rejects(distTag({ version: '0.0.2-rc.1', tag: 'latest', registry: NPMJS_URL, token: 't', fetchImpl, spawn: never, log: () => {} }), /@officesdk\/officesdk-plugin-dsh@0\.0\.2-rc\.1 is not published on https:\/\/registry\.npmjs\.org\//);
  await assert.rejects(distTag({ version: '0.0.0-alpha.3', tag: 'latest', registry: NPMJS_URL, token: 't', dryRun: true, fetchImpl, spawn: never, log: () => {} }), /latest is at 0\.0\.1-rc\.1, newer than 0\.0\.0-alpha\.3: moving it back .* roll back instead/);
  await assert.rejects(distTag({ version: '0.0.0-alpha.3', tag: 'rc', registry: NPMJS_URL, token: 't', fetchImpl, spawn: never, log: () => {} }), /rc takes rc versions or more mature ones/);
  // A tag whose version cannot be compared (not a version of ours) does not block the move.
  const odd = { ...tags, beta: 'garbage' };
  const replaced = await distTag({ version: '0.0.1-rc.1', tag: 'beta', registry: NPMJS_URL, token: 't', fetchImpl: tagRegistry(VERSIONS, odd).fetchImpl, spawn: npmStub({ tags: odd }).spawn, log: () => {}, sleep: async () => {} });
  assert.deepEqual([replaced.result, replaced.from], ['moved', 'garbage']);
  // A tag the registry does not have yet: from is null, the move goes ahead.
  const result = await distTag({ version: '0.0.1-rc.1', tag: 'beta', registry: NPMJS_URL, token: 't', fetchImpl, spawn: npmStub({ tags }).spawn, log: () => {}, sleep: async () => {} });
  assert.deepEqual([result.result, result.from], ['moved', null]);
});

test('release notes: the attached package must be npm\'s (name, version, integrity); the notes link the npm version page', async () => {
  const tgz = await packed('0.0.1-rc.1', { legacy: true });
  await writeFile(join(dirname(tgz), 'package', 'CHANGELOG.md'), '# Changelog\n\n## 0.0.1-rc.10 — 2026-12-01\n\n- later\n\n## 0.0.1-rc.1 — 2026-11-23\n\n初始开源版本。\n\n### Highlights\n\n- First release.\n\n## 0.0.0-alpha.3 — 2026-11-13\n\n- older\n');
  assert.equal(spawnSync('tar', ['-czf', tgz, '-C', dirname(tgz), 'package']).status, 0);
  const integrity = sri512(await readFile(tgz));
  const npmjs = versions => async () => new Response(JSON.stringify({ 'dist-tags': {}, versions }), { status: 200 });
  const notes = await releaseNotes({ version: '0.0.1-rc.1', tgz, registry: NPMJS_URL, fetchImpl: npmjs({ '0.0.1-rc.1': { dist: { integrity } } }) });
  assert.match(notes, /\*\*npm:\*\* \[@officesdk\/officesdk-plugin-dsh@0\.0\.1-rc\.1\]\(https:\/\/www\.npmjs\.com\/package\/@officesdk\/officesdk-plugin-dsh\/v\/0\.0\.1-rc\.1\)/);
  assert.match(notes, /dsh plugin --profile desktop add @officesdk\/officesdk-plugin-dsh@0\.0\.1-rc\.1/);
  assert.ok(notes.includes(`(integrity \`${integrity}\`)`), 'the integrity of the attached file');
  assert.match(notes, /The tag `v0\.0\.1-rc\.1` marks this repository's main when the release was made .* not necessarily the source this version was built from: the package is the attached file\./);
  assert.match(notes, /^## Changes\n\n初始开源版本。\n\n### Highlights\n\n- First release\.\n\n\*\*npm:/, 'its own CHANGELOG section only (not rc.10, not alpha.3)');
  await assert.rejects(releaseNotes({ version: '0.0.1-rc.1', tgz, registry: NPMJS_URL, fetchImpl: npmjs({ '0.0.1-rc.1': { dist: { integrity: 'sha512-other' } } }) }), /is not the package https:\/\/registry\.npmjs\.org\/ has for 0\.0\.1-rc\.1/);
  await assert.rejects(releaseNotes({ version: '0.0.1-rc.1', tgz, registry: NPMJS_URL, fetchImpl: npmjs({}) }), /@officesdk\/officesdk-plugin-dsh@0\.0\.1-rc\.1 is not published/);
  await assert.rejects(releaseNotes({ version: '0.0.1-rc.2', tgz, registry: NPMJS_URL, fetchImpl: npmjs({}) }), /is @officesdk\/officesdk-plugin-dsh@0\.0\.1-rc\.1, not @officesdk\/officesdk-plugin-dsh@0\.0\.1-rc\.2/);
  await assert.rejects(releaseNotes({ version: '0.0.1-rc.1', tgz: join(dirname(tgz), 'missing.tgz'), registry: NPMJS_URL, fetchImpl: npmjs({}) }), /missing\.tgz does not exist \(npm pack @officesdk\/officesdk-plugin-dsh@0\.0\.1-rc\.1 writes it\)/);
  // No section for the version: the notes go without one.
  const bare = await packed('0.0.1-rc.1', { legacy: true });
  await writeFile(join(dirname(bare), 'package', 'CHANGELOG.md'), '# Changelog\n');
  assert.equal(spawnSync('tar', ['-czf', bare, '-C', dirname(bare), 'package']).status, 0);
  const plain = await releaseNotes({ version: '0.0.1-rc.1', tgz: bare, registry: NPMJS_URL, fetchImpl: npmjs({ '0.0.1-rc.1': { dist: { integrity: sri512(await readFile(bare)) } } }) });
  assert.doesNotMatch(plain, /## Changes/);
  const modern = await packed('0.1.2');
  const body = await releaseNotes({ version: '0.1.2', tgz: modern, registry: NPMJS_URL, fetchImpl: npmjs({ '0.1.2': { dist: { integrity: sri512(await readFile(modern)) } } }) });
  assert.match(body, /^## Changes\n\n### 更新要点\n\n- 版本说明\n\n### Highlights\n\n- Version notes\n\n\*\*npm:/);
});

test('the CLI: dist-tag and release-notes need their arguments; an unknown command lists the four', () => {
  const script = fileURLToPath(new URL('../npm-release.mjs', import.meta.url));
  const cli = (...args) => spawnSync(process.execPath, [script, ...args], { encoding: 'utf8', env: { ...process.env, NPM_TOKEN: '', NPM_REGISTRY: 'http://127.0.0.1:1/' } });
  assert.match(cli('dist-tag', '--version', '0.0.1-rc.1').stderr, /dist-tag needs --version <version> and --tag <tag>/);
  assert.match(cli('dist-tag', '--version', '0.0.1-rc.1', '--tag', 'stable', '--dry-run').stderr, /--tag takes alpha, beta, rc, next, latest, not "stable"/);
  assert.match(cli('release-notes', '--version', '0.0.1-rc.1').stderr, /release-notes needs --version <version> and --package <tgz>/);
  const unknown = cli('move');
  assert.notEqual(unknown.status, 0);
  assert.match(unknown.stderr, /unknown command move \(publish, rollback, dist-tag, release-notes\)/);
});
