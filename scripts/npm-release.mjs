#!/usr/bin/env node
/**
 * npm releases of @officesdk/officesdk-plugin-dsh (spec 261111-npm-publish-prep §5, §13.1, §13.5, §13.6, §13.10).
 *
 *   publish (--package <tgz> | --version <version>) --tag alpha|beta|rc|latest|next [--registry <url>] [--dry-run]
 *       --version takes the build of that version made for the registry (buildFor): npmjs → the public build
 *       artifacts/npm/<file>, any other registry → the build made for it (pack-universal --registry) artifacts/<file>; the
 *       npm-publish workflow (.github/workflows/npm-publish.yml) uses it. --dry-run runs the same choice and gate (it reads
 *       the registry) without NPM_TOKEN and without `npm publish`, and writes no artifacts/npm-publish.json.
 *       The tag is the version's stage (packages/dsh-adapter/src/host/version.mjs): alpha|beta|rc for X.Y.Z-<stage>.N,
 *       latest for X.Y.Z, next for X.Y.Z-test.N. Authentication: NPM_TOKEN, else (npmjs only) the GitHub Actions OIDC token
 *       of a job with `id-token: write`, which the npm CLI trades itself (npm trusted publishing); neither → warning,
 *       exit 0. `next` before the registry has any `latest` → warning, exit 0 (npm would make the test build latest); a
 *       preview stage may come first. The release gate must pass (target npmjs for npmjs, local for any other registry);
 *       the same version with the same bytes is already done, with other bytes it fails.
 *   rollback --to <version> [--registry <url>] --out <tgz>
 *       That version's npm package again as the next number of its stage (a release: the next patch; 0.0.0-alpha.3 after
 *       0.0.0-alpha.2), unless its office-kit may not run after the stage's current one; prints { version, tag, out }:
 *       publish it with `publish --tag <tag>`.
 *   dist-tag --version <version> --tag alpha|beta|rc|latest|next [--registry <url>] [--dry-run]
 *       Points the tag at a version already on the registry (distTag): the npm-publish workflow's mode dist-tag runs it.
 *       Authentication as publish (NPM_TOKEN, else the job's OIDC token, which npm 11.21.0 or later trades for
 *       `npm dist-tag`); --dry-run lists the tags with `npm dist-tag ls` and changes nothing. Prints the result as JSON.
 *   release-notes --version <version> --package <tgz> [--registry <url>]
 *       Checks that the tgz is the package the registry has for the version and prints the notes of its GitHub
 *       Release (releaseNotes; the github-release workflow).
 *
 * NPM_REGISTRY replaces https://registry.npmjs.org/ (drills: a local registry). Writes artifacts/npm-publish.json (with the
 * source commit, its tree and the GitHub Actions run, sourceOf).
 */
import { maintenanceNotes, parseWhatsNew, releasePreview, renderWhatsNew } from './whats-new.mjs';
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { gate, readRegistry, sri512 } from './npm-gate.mjs';
import { NPMJS, PACKAGE_NAME } from './universal-package.mjs';
import { stampCapabilities } from './capabilities.mjs';
import { downgradeProblem } from '../packages/dsh-adapter/src/host/fetch/office-kit.mjs';
import { DIST_TAGS, compareVersions as compareAny, distTagOf, newer as newerVersion, parseVersion, publishableVersion, stageOf } from '../packages/dsh-adapter/src/host/version.mjs';

/** Runs a command and returns its stdout; a failure's message hides credentials in URLs (`//user:token@` → `//***@`). */
export function run(command, args, options = {}) {
  const result = spawnSync(command, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 256 << 20, ...options });
  if (result.status !== 0) {
    const stderr = String(result.stderr ?? '');
    throw Object.assign(new Error(`${command} ${args.map(arg => String(arg).replace(/\/\/[^/@]+@/, '//***@')).join(' ')} failed:\n${(stderr || String(result.stdout ?? '')).replace(/\/\/[^/@\s]+@/g, '//***@')}`), { stderr, status: result.status });
  }
  return result.stdout;
}

/** Re-versions an unpacked package: package.json version and capabilities.json (version, channel, highlights). */
export async function restamp(packageRoot, version) {
  const file = path.join(packageRoot, 'package.json');
  const manifest = JSON.parse(await fs.readFile(file, 'utf8'));
  manifest.version = version;
  const notes = parseWhatsNew(await fs.readFile(path.join(packageRoot, 'CHANGELOG.md'), 'utf8'));
  manifest.officeSdkReleaseNotes = releasePreview(notes, version);
  await fs.writeFile(path.join(packageRoot, 'whats-new.json'), renderWhatsNew(notes));
  await fs.writeFile(file, JSON.stringify(manifest, null, 2) + '\n');
  stampCapabilities(packageRoot, { version });
}

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SKIPPED = 'skipped';
const slash = url => url.replace(/\/?$/, '/');
const stable = version => /^\d+\.\d+\.\d+$/.test(version ?? '');
const triple = version => version.split('.').map(Number);
/** The public npm registry (the gate's strict `npmjs` target and the public build); every other registry is `local`. */
export const isNpmjs = registry => slash(registry) === NPMJS;

/**
 * The build of `version` that goes to `registry` (tooling/suite/pack-universal.mjs makes either): npmjs takes the public
 * npm build (`--sources public --out <artifacts>/npm`), any other registry the build made for it (`--registry`, <artifacts>/).
 */
export function buildFor({ version, registry, artifacts = 'artifacts' }) {
  if (!version) throw new Error('buildFor needs a version');
  const file = `${PACKAGE_NAME.replace(/^@/, '').replace('/', '-')}-${version}.tgz`;
  return isNpmjs(registry) ? path.join(artifacts, 'npm', file) : path.join(artifacts, file);
}
const higher = (a, b) => { const [x, y] = [triple(a), triple(b)]; for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] > y[i]; return false; };

async function packageJsonOf(tgz) {
  const work = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-release-'));
  try { run('tar', ['-xzf', tgz, '-C', work, 'package/package.json']); return JSON.parse(await fs.readFile(path.join(work, 'package/package.json'), 'utf8')); }
  finally { await fs.rm(work, { recursive: true, force: true }); }
}

/**
 * The job can get a GitHub Actions OIDC token (GitHub sets these two only for a job with `permissions: id-token: write`):
 * the npm CLI (11.5.1 or later) then publishes through npm trusted publishing, without NPM_TOKEN.
 */
export const oidcAvailable = (env = process.env) => env.GITHUB_ACTIONS === 'true' && Boolean(env.ACTIONS_ID_TOKEN_REQUEST_URL) && Boolean(env.ACTIONS_ID_TOKEN_REQUEST_TOKEN);

/**
 * `npm publish` with a throwaway userconfig (never written into the repository or a profile, and the user's ~/.npmrc
 * stays out): it holds the token, or without one nothing, and the npm CLI trades the job's OIDC token itself. No
 * --provenance: npm adds it to a trusted publish from a public repository and refuses it from a private one.
 */
export async function npmPublish({ tgz, tag, registry, token, spawn = spawnSync }) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-release-auth-'));
  const userconfig = path.join(dir, '.npmrc');
  const host = new URL(registry);
  await fs.writeFile(userconfig, token ? `//${host.host}${host.pathname.replace(/\/?$/, '/')}:_authToken=${token}\n` : '', { mode: 0o600 });
  try {
    // An absolute path: npm 11 reads a relative `artifacts/x.tgz` as the GitHub shorthand `artifacts/x.tgz` (git ls-remote).
    const result = spawn('npm', ['publish', path.resolve(tgz), '--tag', tag, '--registry', registry, '--access', 'public', '--userconfig', userconfig], { encoding: 'utf8', env: { ...process.env, npm_config_userconfig: userconfig } });
    if (result.status !== 0) {
      const stderr = String(result.stderr ?? '');
      const hint = token ? '' : ' (npm trusted publishing: check the trusted publisher on npmjs.com — workflow filename npm-publish.yml, Environment, Allowed actions with `npm publish` — docs/NPM.md)';
      throw new Error(`npm publish failed: ${(token ? stderr.replaceAll(token, '***') : stderr).trim().split('\n').slice(-3).join(' ')}${hint}`);
    }
    return String(result.stdout ?? '').trim();
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
}

/** What a publish without NPM_TOKEN and without a GitHub Actions OIDC token says (and then publishes nothing). */
export const NO_TOKEN_WARNING = 'WARNING: NPM_TOKEN is not set and this is no GitHub Actions job with an OIDC token: nothing is published to npm. Publish with the npm-publish workflow (npm trusted publishing, or the NPM_TOKEN secret in the repository\'s Settings → Secrets and variables → Actions), or by hand: NPM_TOKEN=<token> node scripts/npm-release.mjs publish … (docs/NPM.md)';

/**
 * Publishes `tgz` under `tag`; returns what happened (`published` with `auth` — `token` or `oidc`, what this script
 * supplied — `already-published`, or SKIPPED with the reason).
 */
export async function publish({ tgz, tag, registry = process.env.NPM_REGISTRY || NPMJS, token = process.env.NPM_TOKEN, oidc = oidcAvailable(), fetchImpl, publishImpl = npmPublish, log = message => console.warn(message), allowFile, kitManifestFile, internalHosts, rollback = false, dryRun = false }) {
  if (!DIST_TAGS.includes(tag)) throw new Error(`--tag takes ${DIST_TAGS.join(', ')}, not ${tag}`);
  registry = slash(registry);
  // Trusted publishing is npmjs's: a drill registry still needs NPM_TOKEN.
  const viaOidc = !token && oidc && isNpmjs(registry);
  if (!token && !viaOidc && !dryRun) { log(NO_TOKEN_WARNING); return { result: SKIPPED, reason: 'no-token' }; }
  const integrity = sri512(await fs.readFile(tgz));
  const manifest = await packageJsonOf(tgz);
  const { packument } = await readRegistry(registry, manifest.name, fetchImpl);
  if (tag === 'next' && !packument?.['dist-tags']?.latest) {
    log(`WARNING: ${registry} has no ${manifest.name}@latest yet: the first publish must be a release, so ${manifest.version} is not published under next`);
    return { result: SKIPPED, reason: 'no-latest' };
  }
  const target = isNpmjs(registry) ? 'npmjs' : 'local';
  const report = await gate(tgz, { target, registry, tag, fetchImpl, rollback, ...allowFile && { allowFile }, ...kitManifestFile && { kitManifestFile }, ...internalHosts && { internalHosts } });
  if (!report.ok) throw new Error(`the npm release gate refuses ${report.package}:\n${report.errors.join('\n')}`);
  for (const warning of report.warnings) log(`WARNING: ${warning}`);
  if (dryRun) return { result: 'dry-run', package: report.package, tgz, tag, registry, target, alreadyPublished: Boolean(report.alreadyPublished), integrity };
  if (report.alreadyPublished) return { result: 'already-published', package: report.package, integrity };
  await publishImpl({ tgz, tag, registry, token: token || undefined });
  return { result: 'published', package: report.package, tag, registry, integrity, auth: viaOidc ? 'oidc' : 'token' };
}

/**
 * The version a rollback publishes: for a release, one patch above the highest release on npm (and `gitTags`, when given); for a preview
 * stage (`stage`: alpha, beta, rc), the next number of that stage on the highest X.Y.Z with it (0.0.0-alpha.3 after
 * 0.0.0-alpha.2), published under the stage's tag.
 */
export function rollbackVersion(npmVersions, gitTags = [], stage = 'release') {
  if (stage !== 'release') {
    const line = npmVersions.map(parseVersion).filter(v => v?.stage === stage);
    if (!line.length) throw new Error(`no ${stage} version to roll back from`);
    const top = line.reduce((best, v) => (v.major - best.major || v.minor - best.minor || v.patch - best.patch || v.number - best.number) > 0 ? v : best);
    const version = `${top.major}.${top.minor}.${top.patch}-${stage}.${top.number + 1}`;
    // The stage's channel also takes every more mature version (version.mjs channelTags): a rollback below one never arrives.
    const ahead = npmVersions.filter(other => newerVersion(other, version) && ['alpha', 'beta', 'rc', 'release'].indexOf(stageOf(other)) > ['alpha', 'beta', 'rc'].indexOf(stage));
    if (ahead.length) throw new Error(`${version} would be below ${ahead.sort(compareAny).at(-1)}, which ${stage} users already take: roll back that stage instead`);
    return version;
  }
  const candidates = [...npmVersions.filter(stable), ...gitTags.map(tag => /^v(\d+\.\d+\.\d+)$/.exec(tag)?.[1]).filter(Boolean)];
  if (!candidates.length) throw new Error('no release to roll back from');
  const highest = candidates.reduce((best, version) => (higher(version, best) ? version : best));
  const [x, y, z] = triple(highest);
  return `${x}.${y}.${z + 1}`;
}

/**
 * Rollback: the npm package of `to`, re-versioned as the next patch, written to `out`. Refuses a target whose office-kit
 * may not run after the current release's (the git rollback's check, §4.3 of 261110). Returns `{ version, out }`.
 */
export async function rollback({ to, registry = process.env.NPM_REGISTRY || NPMJS, out, gitTags = [], fetchImpl = fetch }) {
  registry = slash(registry);
  const { packument } = await readRegistry(registry, PACKAGE_NAME, fetchImpl);
  const target = packument?.versions?.[to];
  const stage = stageOf(to);
  if (!publishableVersion(to) || stage === 'test' || !target) throw new Error(`${PACKAGE_NAME}@${to} is not a release or preview version on ${registry}`);
  const tag = distTagOf(to);
  const current = packument['dist-tags']?.[tag];
  const version = rollbackVersion(Object.keys(packument.versions), gitTags, stage);
  const work = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-rollback-'));
  const fetchPackage = async (meta, name) => {
    const response = await fetchImpl(meta.dist.tarball);
    if (!response.ok) throw new Error(`${meta.dist.tarball}: HTTP ${response.status}`);
    const bytes = Buffer.from(await response.arrayBuffer());
    if (meta.dist.integrity && sri512(bytes) !== meta.dist.integrity) throw new Error(`${name}: the downloaded package does not match its integrity`);
    const file = path.join(work, `${name}.tgz`);
    await fs.writeFile(file, bytes);
    await fs.mkdir(path.join(work, name));
    run('tar', ['-xzf', file, '-C', path.join(work, name)]);
    return path.join(work, name, 'package');
  };
  const targetRoot = await fetchPackage(target, 'target');
  if (current && packument.versions[current]) {
    const currentRoot = await fetchPackage(packument.versions[current], 'current');
    const read = root => fs.readFile(path.join(root, 'kit-manifest.json'), 'utf8').then(JSON.parse);
    const refused = downgradeProblem(await read(targetRoot), (await read(currentRoot)).officeKit);
    if (refused) throw new Error(`rolling back to ${to} would run an older office-kit after ${current}: ${refused.message.split('。')[0]}. Release ${to}'s plugin code with the current office-kit instead.`);
  }
  const changelog = path.join(targetRoot, 'CHANGELOG.md');
  const text = await fs.readFile(changelog, 'utf8');
  const at = text.indexOf('\n## ');
  await fs.writeFile(changelog, `${text.slice(0, at + 1)}## ${version} — ${new Date().toISOString().slice(0, 10)}\n\n### 更新要点\n\n- ${maintenanceNotes.highlights[0]}\n\n### Highlights\n\n- ${maintenanceNotes.highlightsEn[0]}\n\n### 技术记录\n\n回滚到 ${to}。\n${text.slice(at)}`);
  await restamp(targetRoot, version);
  await fs.mkdir(path.dirname(path.resolve(out)), { recursive: true });
  run('tar', ['-czf', path.resolve(out), '-C', path.dirname(targetRoot), 'package']);
  return { version, tag, out: path.resolve(out) };
}

/**
 * Where a publish comes from: the commit (GITHUB_SHA on GitHub Actions, else this checkout's HEAD), its tree (the
 * public main is rebuilt without history, so a commit may be gone where its tree is not) and the GitHub Actions run.
 * A field git cannot answer is left out.
 */
export function sourceOf(env = process.env, git = args => run('git', args, { cwd: repo }).trim()) {
  const ask = args => { try { return git(args) || undefined; } catch { return undefined; } };
  const commit = env.GITHUB_SHA || ask(['rev-parse', 'HEAD']);
  const actions = env.GITHUB_ACTIONS === 'true' && env.GITHUB_SERVER_URL && env.GITHUB_REPOSITORY && env.GITHUB_RUN_ID;
  return { commit, tree: commit ? ask(['rev-parse', `${commit}^{tree}`]) : undefined, run: actions ? `${env.GITHUB_SERVER_URL}/${env.GITHUB_REPOSITORY}/actions/runs/${env.GITHUB_RUN_ID}` : undefined };
}

async function record(file, value) {
  await fs.mkdir(path.dirname(file), { recursive: true });
  await fs.writeFile(file, `${JSON.stringify({ ...value, ...sourceOf(), at: new Date().toISOString() }, null, 2)}\n`);
}

/** The package `publish` sends: --package as given, or --version's build for the registry (buildFor), which must exist. */
export function publishPackage({ package: tgz, version }, registry, exists = existsSync) {
  if (tgz && version) throw new Error('publish takes --package or --version, not both');
  if (tgz) return tgz;
  if (!version) throw new Error('publish needs --package <tgz> or --version <version>');
  const chosen = buildFor({ version, registry });
  if (!exists(chosen)) throw new Error(`${chosen} does not exist: ${isNpmjs(registry) ? 'the public npm build (pack-universal.mjs --sources public --out artifacts/npm)' : 'the build for this registry (pack-universal.mjs --registry <url> --out artifacts)'} of ${version} is missing`);
  return chosen;
}

/**
 * Whether `version` may take the dist-tag `tag` (`npm-release.mjs dist-tag`); the problem in words, or undefined. `next`
 * takes test builds (X.Y.Z-test.N) only and a test build goes under `next` only; `alpha`, `beta` and `rc` take their own
 * stage or a more mature one (the channels users follow, version.mjs channelTags: an rc user must never get an alpha);
 * `latest` takes a release or any preview stage (before the first release, so that an install without a tag gets the
 * preview).
 */
export function distTagProblem(version, tag) {
  if (!DIST_TAGS.includes(tag)) return `--tag takes ${DIST_TAGS.join(', ')}, not ${JSON.stringify(tag ?? '')}`;
  if (!publishableVersion(version)) return `${JSON.stringify(version ?? '')} is not X.Y.Z or X.Y.Z-alpha|beta|rc|test.N`;
  const stage = stageOf(version);
  if (tag === 'next' && stage !== 'test') return `next takes test builds (X.Y.Z-test.N) only, not ${version}`;
  if (tag !== 'next' && stage === 'test') return `${version} is a test build: it goes under next only`;
  const maturity = ['alpha', 'beta', 'rc', 'release'];
  if (tag !== 'latest' && tag !== 'next' && maturity.indexOf(stage) < maturity.indexOf(tag)) return `${tag} takes ${tag} versions or more mature ones, not the ${stage} version ${version}`;
  return undefined;
}

/** The registry's dist-tags of `name` (`/-/package/<name>/dist-tags`: no login for a public package, read fresh). */
export async function readDistTags(registry, name, fetchImpl = fetch) {
  const response = await fetchImpl(`${slash(registry)}-/package/${name.replace('/', '%2F')}/dist-tags`, { headers: { Accept: 'application/json', 'Cache-Control': 'no-cache' }, signal: AbortSignal.timeout(30_000) });
  if (!response.ok) throw new Error(`${slash(registry)}: the dist-tags of ${name}: HTTP ${response.status}`);
  const tags = await response.json();
  return tags && typeof tags === 'object' ? tags : {};
}

/** An npm CLI that trades a CI OIDC token for `npm dist-tag` (npm trusted publishing): 11.21.0 or later, 12.2.0 or later. */
export function distTagOidcCli(version) {
  const match = /^(\d+)\.(\d+)\.(\d+)/.exec(String(version ?? '').trim());
  if (!match) return false;
  const [major, minor] = [Number(match[1]), Number(match[2])];
  return major === 11 ? minor >= 21 : major === 12 ? minor >= 2 : major > 12;
}

/**
 * Runs `npm <args>` with a throwaway userconfig (the token, or nothing for the npm CLI's own OIDC exchange; the user's
 * ~/.npmrc stays out; removed afterwards) at --loglevel verbose. Returns the exit status, stdout, and from stderr only
 * the `npm verbose oidc` and `npm error` lines (the token replaced by ***): verbose logs carry other detail not shown.
 */
async function npmCommand(args, { registry, token, spawn }) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-release-auth-'));
  const userconfig = path.join(dir, '.npmrc');
  const host = new URL(registry);
  await fs.writeFile(userconfig, token ? `//${host.host}${host.pathname.replace(/\/?$/, '/')}:_authToken=${token}\n` : '', { mode: 0o600 });
  try {
    const result = spawn('npm', [...args, '--registry', registry, '--userconfig', userconfig, '--loglevel', 'verbose'], { encoding: 'utf8', maxBuffer: 64 << 20, env: { ...process.env, npm_config_userconfig: userconfig } });
    const hide = text => (token ? String(text ?? '').replaceAll(token, '***') : String(text ?? ''));
    const lines = hide(result.stderr).split('\n');
    const oidc = lines.filter(line => line.startsWith('npm verbose oidc ')).map(line => line.slice('npm verbose oidc '.length).trim());
    const errors = lines.filter(line => /^npm (error|ERR!) /.test(line) && !/complete log of this run/.test(line)).map(line => line.replace(/^npm (error|ERR!) /, '').trim());
    return { status: result.status, error: result.error, stdout: hide(result.stdout), oidc, errors };
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
}

/** What the npm CLI's OIDC exchange did, from its verbose `oidc` lines (it never fails a command by itself). */
export function oidcOutcome(lines) {
  if (lines.some(line => line.startsWith('Successfully retrieved and set token'))) return 'exchanged';
  const failed = lines.filter(line => /^Fail/.test(line));
  return failed.length ? `failed: ${failed.at(-1)}` : 'not-attempted';
}

/** What a dist-tag move without NPM_TOKEN and without a GitHub Actions OIDC token says (and then changes nothing). */
export const NO_AUTH_DIST_TAG_WARNING = 'WARNING: NPM_TOKEN is not set and this is no GitHub Actions job with an OIDC token: no dist-tag is changed. Move it with the npm-publish workflow in mode dist-tag (npm trusted publishing: the trusted publisher for npm-publish.yml with "Allow npm dist-tag"), or by hand with your own npm login: npm dist-tag add @officesdk/officesdk-plugin-dsh@<version> <tag> (docs/NPM.md)';
/** What an `npm dist-tag add` failing through npm trusted publishing most likely means, by npm's error code. */
export function distTagOidcHint(errors) {
  const text = errors.join(' ');
  if (/\bEOTP\b/.test(text)) return ' — npm asked for a one-time password, which npm trusted publishing does not need: check the package\'s Publishing access and its trusted publisher on npmjs.com (docs/NPM.md)';
  if (/\bE403\b|403 Forbidden/.test(text)) return ' — the trusted publisher for npm-publish.yml may not move dist-tags: tick "Allow npm dist-tag" on it on npmjs.com (docs/NPM.md)';
  return ' — npm trusted publishing did not authenticate: check the trusted publisher on npmjs.com (workflow filename npm-publish.yml, repository officesdk/officesdk-plugin-dsh, Environment empty unless the publish and dist-tag jobs both run in it) and npm 11.21.0 or later (docs/NPM.md)';
}

/**
 * Points the dist-tag `tag` at the published `version` (`npm-release.mjs dist-tag`). The rules are distTagProblem's; the
 * version must be on the registry; a tag already there is `unchanged` (npm is not called); a tag at a newer version is
 * not moved back (users who updated keep theirs, the plugin never downgrades: roll back instead). Authentication as
 * publish: NPM_TOKEN, else (npmjs only) the job's GitHub Actions OIDC token, which the npm CLI (11.21.0 or later) trades
 * itself; neither → warning, `skipped`. A dry run runs `npm dist-tag ls` (with OIDC it also shows whether the exchange
 * works) and changes nothing; a move runs `npm dist-tag add` and reads the registry's dist-tags back until the tag is
 * there. Returns `{ result: 'dry-run' | 'unchanged' | 'moved' | 'skipped', … }`.
 */
export async function distTag({ version, tag, registry = process.env.NPM_REGISTRY || NPMJS, token = process.env.NPM_TOKEN, oidc = oidcAvailable(), dryRun = false, fetchImpl = fetch, spawn = spawnSync, log = message => console.warn(message), sleep = ms => new Promise(resolve => setTimeout(resolve, ms)), attempts = 5 }) {
  const problem = distTagProblem(version, tag);
  if (problem) throw new Error(problem);
  registry = slash(registry);
  const viaOidc = !token && oidc && isNpmjs(registry);
  const auth = token ? 'token' : viaOidc ? 'oidc' : 'none';
  if (auth === 'none' && !dryRun) { log(NO_AUTH_DIST_TAG_WARNING); return { result: SKIPPED, reason: 'no-token' }; }
  const { packument } = await readRegistry(registry, PACKAGE_NAME, fetchImpl);
  if (!packument?.versions?.[version]) throw new Error(`${PACKAGE_NAME}@${version} is not published on ${registry}: only a published version takes a dist-tag`);
  const from = (await readDistTags(registry, PACKAGE_NAME, fetchImpl))[tag] ?? null;
  if (from && from !== version && compareAny(version, from) < 0) throw new Error(`${tag} is at ${from}, newer than ${version}: moving it back does not reach users who updated (the plugin never downgrades); roll back instead (docs/NPM.md, 回滚)`);
  const base = { package: `${PACKAGE_NAME}@${version}`, version, tag, registry, from, auth };
  if (!dryRun && from === version) return { result: 'unchanged', ...base };
  if (viaOidc) {
    const cli = spawn('npm', ['--version'], { encoding: 'utf8' });
    const have = String(cli.stdout ?? '').trim();
    if (cli.status !== 0 || !distTagOidcCli(have)) throw new Error(`npm ${have || '(not found)'} does not move dist-tags through npm trusted publishing: it needs npm 11.21.0 or later (12.2.0 or later on npm 12)`);
  }
  if (dryRun) {
    const listed = await npmCommand(['dist-tag', 'ls', PACKAGE_NAME], { registry, token, spawn });
    if (listed.status !== 0) throw new Error(`npm dist-tag ls failed: ${[...listed.errors.slice(-3), listed.error?.message].filter(Boolean).join(' ')}`);
    const tags = Object.fromEntries(listed.stdout.split('\n').map(line => /^([^:\s]+): (\S+)$/.exec(line.trim())).filter(Boolean).map(match => [match[1], match[2]]));
    return { result: 'dry-run', ...base, change: from === version ? 'none' : 'move', ...viaOidc && { oidc: oidcOutcome(listed.oidc) }, tags };
  }
  const added = await npmCommand(['dist-tag', 'add', `${PACKAGE_NAME}@${version}`, tag], { registry, token, spawn });
  if (added.status !== 0) {
    const detail = [...added.oidc.map(line => `oidc: ${line}`), ...added.errors.slice(-3), added.error?.message].filter(Boolean).join(' ');
    throw new Error(`npm dist-tag add failed: ${detail}${viaOidc ? distTagOidcHint(added.errors) : ''}`);
  }
  for (let attempt = 1; ; attempt++) {
    const now = (await readDistTags(registry, PACKAGE_NAME, fetchImpl))[tag];
    if (now === version) return { result: 'moved', ...base };
    if (attempt >= attempts) throw new Error(`npm dist-tag add reported success, but ${registry} still has ${tag} at ${now ?? '(none)'}, not ${version}`);
    await sleep(2000);
  }
}

/** The `## <version>` section of a CHANGELOG (without its heading), or undefined. */
export function changelogSection(text, version) {
  const lines = String(text ?? '').split('\n');
  const at = lines.findIndex(line => line.startsWith(`## ${version}`) && /^(\s|$)/.test(line.slice(3 + version.length)));
  if (at < 0) return undefined;
  const end = lines.findIndex((line, index) => index > at && line.startsWith('## '));
  const body = lines.slice(at + 1, end < 0 ? undefined : end).join('\n').trim();
  return body || undefined;
}

/**
 * The notes of the GitHub Release of `version` (`npm-release.mjs release-notes`), whose asset is `tgz`: it must be the
 * package npm has for that version (name, version and integrity against the registry), or this fails. The notes link
 * the npm version page, give the install command, name the asset's integrity, say what the tag marks, and carry the
 * package's own CHANGELOG section for the version.
 */
export async function releaseNotes({ version, tgz, registry = NPMJS, fetchImpl = fetch }) {
  if (!publishableVersion(version)) throw new Error(`${JSON.stringify(version ?? '')} is not X.Y.Z or X.Y.Z-alpha|beta|rc|test.N`);
  registry = slash(registry);
  if (!existsSync(tgz)) throw new Error(`${tgz} does not exist (npm pack ${PACKAGE_NAME}@${version} writes it)`);
  const manifest = await packageJsonOf(tgz);
  if (manifest.name !== PACKAGE_NAME || manifest.version !== version) throw new Error(`${tgz} is ${manifest.name}@${manifest.version}, not ${PACKAGE_NAME}@${version}`);
  const integrity = sri512(await fs.readFile(tgz));
  const { packument } = await readRegistry(registry, PACKAGE_NAME, fetchImpl);
  const published = packument?.versions?.[version]?.dist?.integrity;
  if (!published) throw new Error(`${PACKAGE_NAME}@${version} is not published on ${registry}`);
  if (published !== integrity) throw new Error(`${tgz} is not the package ${registry} has for ${version} (integrity ${integrity}, the registry's ${published})`);
  let changes;
  if (manifest.officeSdkReleaseNotes) {
    const preview = releasePreview({ entries: [manifest.officeSdkReleaseNotes] }, version);
    changes = ['### 更新要点', '', ...preview.highlights.map(text => `- ${text}`), '', '### Highlights', '', ...preview.highlightsEn.map(text => `- ${text}`)].join('\n');
  } else {
    const work = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-release-notes-'));
    try {
      run('tar', ['-xzf', tgz, '-C', work, 'package/CHANGELOG.md']);
      changes = changelogSection(await fs.readFile(path.join(work, 'package/CHANGELOG.md'), 'utf8'), version);
    } catch { changes = undefined; } finally { await fs.rm(work, { recursive: true, force: true }); }
  }
  const page = `https://www.npmjs.com/package/${PACKAGE_NAME}/v/${version}`;
  return [
    ...changes ? ['## Changes', '', changes, ''] : [],
    `**npm:** [${PACKAGE_NAME}@${version}](${page})`,
    '',
    'Install (quit the DeepSeek Harness desktop app first):',
    '',
    '```sh',
    `dsh plugin --profile desktop add ${PACKAGE_NAME}@${version}`,
    '```',
    '',
    `The attached \`${path.basename(tgz)}\` is the package published to npm (integrity \`${integrity}\`).`,
    '',
    `The tag \`v${version}\` marks this repository's main when the release was made (main is rebuilt as one commit without history, so the tag keeps that snapshot). It is not necessarily the source this version was built from: the package is the attached file.`,
    '',
  ].join('\n');
}

async function main() {
  const [command, ...rest] = process.argv.slice(2);
  const { values } = parseArgs({ args: rest, options: { package: { type: 'string' }, version: { type: 'string' }, tag: { type: 'string' }, registry: { type: 'string' }, out: { type: 'string' }, to: { type: 'string' }, rollback: { type: 'boolean' }, 'dry-run': { type: 'boolean' } } });
  // An empty --registry or NPM_REGISTRY means npmjs.
  const registry = slash(values.registry || process.env.NPM_REGISTRY || NPMJS);
  switch (command) {
    case 'publish': {
      const tgz = publishPackage(values, registry);
      const dryRun = values['dry-run'] === true;
      console.warn(`npm-release: ${registry} takes ${path.resolve(tgz)} (gate ${isNpmjs(registry) ? 'npmjs' : 'local'})${dryRun ? ', dry run: nothing is published' : ''}`);
      // --rollback: the package npm-release.mjs rollback wrote, an older release as it was (its own kit-manifest.json stands).
      const result = await publish({ tgz, tag: values.tag, registry, rollback: values.rollback === true, dryRun });
      if (!dryRun) await record(path.join(repo, 'artifacts', 'npm-publish.json'), result);
      return console.log(JSON.stringify(result));
    }
    case 'rollback': {
      const result = await rollback({ to: values.to, registry, out: values.out ?? path.join(repo, 'artifacts', 'npm-rollback.tgz') });
      return console.log(JSON.stringify(result));
    }
    case 'dist-tag': {
      if (!values.version || !values.tag) throw new Error('dist-tag needs --version <version> and --tag <tag>');
      const dryRun = values['dry-run'] === true;
      console.warn(`npm-release: ${registry} ${values.tag} → ${values.version}${dryRun ? ', dry run: no dist-tag is changed' : ''}`);
      const result = await distTag({ version: values.version, tag: values.tag, registry, dryRun });
      return console.log(JSON.stringify(result));
    }
    case 'release-notes': {
      if (!values.version || !values.package) throw new Error('release-notes needs --version <version> and --package <tgz>');
      return process.stdout.write(await releaseNotes({ version: values.version, tgz: values.package, registry }));
    }
    default: throw new Error(`unknown command ${command} (publish, rollback, dist-tag, release-notes)`);
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main();
