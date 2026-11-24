import { readFileSync } from 'node:fs';
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { DIST_TAGS, channelTags, distTagOf } from './version.mjs';

/**
 * npm installs of the plugin (spec 261111-npm-publish-prep §3, §13.3, §13.7). DSH's `dsh plugin add <name>@<spec>` writes
 * the exact version into the profile's package.json (no registry, no channel), so:
 *   - the channel (`next` test builds, `latest` releases) is kept in the plugin data directory per profile;
 *   - the registry is OFFICESDK_NPM_REGISTRY, else the profile .npmrc (`@officesdk:registry=` or `registry=`), else
 *     npmjs, falling back to npmmirror when npmjs cannot be reached;
 *   - an update or a downgrade is `dsh plugin add <name>@<exact version>` (`dsh plugin update` never leaves an exact pin).
 */
export const PACKAGE = '@officesdk/officesdk-plugin-dsh';
export const NPMJS = 'https://registry.npmjs.org/';
export const NPMMIRROR = 'https://registry.npmmirror.com/';
/** Channels are dist-tags: alpha, beta, rc (preview stages), next (test builds), latest (releases). */
export const CHANNELS = DIST_TAGS;
const SCOPE_LINE = /^\s*@officesdk:registry\s*=\s*(\S+)\s*$/m;
const DEFAULT_LINE = /^\s*registry\s*=\s*(\S+)\s*$/m;
const SEMVER = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/;
const PROBE_TIMEOUT_MS = 5000;

const slash = url => url.replace(/\/?$/, '/');
export const channelOf = version => distTagOf(version) ?? 'latest';

/**
 * This copy as an npm install: the profile depends on @officesdk/officesdk-plugin-dsh by an exact version (or `next`/`latest`).
 * @returns {{ kind: 'npm', profileDir, profile, packageDir, dependencies: string[], spec } | undefined}
 */
export function npmInstall(packageRoot) {
  const packageDir = resolve(packageRoot);
  if (basename(dirname(dirname(packageDir))) !== 'node_modules' || `${basename(dirname(packageDir))}/${basename(packageDir)}` !== PACKAGE) return undefined;
  const profileDir = dirname(dirname(dirname(packageDir)));
  let manifest;
  try { manifest = JSON.parse(readFileSync(join(profileDir, 'package.json'), 'utf8')); } catch { return undefined; }
  const spec = manifest.dependencies?.[PACKAGE];
  if (typeof spec !== 'string' || !(SEMVER.test(spec) || CHANNELS.includes(spec))) return undefined;
  return { kind: 'npm', profileDir, profile: basename(profileDir), packageDir, dependencies: [PACKAGE], spec };
}

/** The registry the profile's .npmrc names for @officesdk (its scope line first), or undefined. */
export function npmrcRegistry(profileDir) {
  let text;
  try { text = readFileSync(join(profileDir, '.npmrc'), 'utf8'); } catch { return undefined; }
  const url = SCOPE_LINE.exec(text)?.[1] ?? DEFAULT_LINE.exec(text)?.[1];
  return url && /^https?:\/\//.test(url) ? slash(url) : undefined;
}

/**
 * Adds or replaces the `@officesdk:registry=` line of the profile .npmrc, keeping every other line (auth, other
 * scopes). Only for a registry the user chose (OFFICESDK_NPM_REGISTRY): later `dsh plugin install` runs then resolve
 * the package without our environment. Returns whether the file changed.
 */
export async function setScopedRegistry(profileDir, url) {
  const file = join(profileDir, '.npmrc');
  let text = '';
  try { text = await readFile(file, 'utf8'); } catch (error) { if (error?.code !== 'ENOENT') throw error; }
  const line = `@officesdk:registry=${slash(url)}`;
  const next = SCOPE_LINE.test(text) ? text.replace(SCOPE_LINE, line) : `${text}${text && !text.endsWith('\n') ? '\n' : ''}${line}\n`;
  if (next === text) return false;
  const tmp = `${file}.${randomUUID()}.tmp`;
  await writeFile(tmp, next, { mode: 0o600 });
  await rename(tmp, file);
  return true;
}

/** The registries to ask, in order: `[{ url, from: 'env'|'npmrc'|'npmjs'|'npmmirror' }]`. */
export function registryCandidates(profileDir, env = process.env) {
  const configured = (env.OFFICESDK_NPM_REGISTRY || '').trim();
  if (/^https?:\/\//.test(configured)) return [{ url: slash(configured), from: 'env' }];
  const npmrc = profileDir && npmrcRegistry(profileDir);
  if (npmrc) return [{ url: npmrc, from: 'npmrc' }];
  return [{ url: NPMJS, from: 'npmjs' }, { url: NPMMIRROR, from: 'npmmirror' }];
}

/**
 * The package's packument from the first registry that answers: `{ registry, from, packument }`. A registry that
 * answers 404 (the package is not there) ends the search only when it was configured; npmjs → npmmirror only on a
 * network failure or a timeout. `full`: the whole document, which also has each version's publish `time` (the
 * abbreviated install document does not).
 */
export async function fetchPackument({ profileDir, env = process.env, fetchImpl = fetch, timeoutMs = PROBE_TIMEOUT_MS, name = PACKAGE, full = false } = {}) {
  const errors = [];
  for (const candidate of registryCandidates(profileDir, env)) {
    let response;
    try { response = await fetchImpl(`${candidate.url}${name.replace('/', '%2F')}`, { headers: { Accept: full ? 'application/json' : 'application/vnd.npm.install-v1+json, application/json' }, signal: AbortSignal.timeout(timeoutMs) }); }
    catch (error) { errors.push(`${candidate.url}: ${error?.cause?.code ?? error?.name ?? error?.message}`); continue; }
    if (!response.ok) { errors.push(`${candidate.url}: HTTP ${response.status}`); if (candidate.from !== 'npmjs') break; continue; }
    return { registry: candidate.url, from: candidate.from, packument: await response.json() };
  }
  throw Object.assign(new Error(`${name} could not be read from npm (${errors.join('; ')})`), { code: 'registry-unavailable' });
}

/**
 * The version a channel offers: the newest of its own tag and every more mature one (version.mjs channelTags): alpha
 * users get beta, rc and releases too, next users releases; latest only its tag.
 */
export function channelVersion(packument, channel, newer) {
  const tags = packument?.['dist-tags'] ?? {};
  return channelTags(channel).map(tag => tags[tag]).filter(Boolean).reduce((best, version) => (best === undefined || newer(version, best) ? version : best), undefined);
}

/** `<data>/npm/channels.json`: `{ "<profileDir>": "alpha"|"beta"|"rc"|"next"|"latest" }`. */
export class ChannelStore {
  constructor(dataRoot) { this.file = dataRoot && join(dataRoot, 'npm', 'channels.json'); }
  read() { try { return JSON.parse(readFileSync(this.file, 'utf8')); } catch { return {}; } }
  get(profileDir) { const value = this.file && this.read()[resolve(profileDir)]; return CHANNELS.includes(value) ? value : undefined; }
  async set(profileDir, channel) {
    if (!this.file || !CHANNELS.includes(channel) || this.get(profileDir) === channel) return;
    await mkdir(dirname(this.file), { recursive: true });
    const tmp = `${this.file}.${randomUUID()}.tmp`;
    await writeFile(tmp, JSON.stringify({ ...this.read(), [resolve(profileDir)]: channel }, null, 2));
    await rename(tmp, this.file);
  }
  /**
   * The profile's channel: OFFICESDK_CHANNEL (written back), else the stored one, else the current version's
   * (its dist-tag: `-alpha.N` → alpha, `-test.N` → next), stored from then on — so a preview user who received a
   * release stays on the preview channel.
   */
  async resolve(profileDir, version, env = process.env) {
    const forced = (env.OFFICESDK_CHANNEL || '').trim();
    const channel = CHANNELS.includes(forced) ? forced : this.get(profileDir) ?? channelOf(version);
    await this.set(profileDir, channel).catch(() => {});
    return channel;
  }
}
