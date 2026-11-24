#!/usr/bin/env node
/**
 * office-kit's editor bundle, the Sheet host fonts and kit-manifest.json are not part of this repository: office-kit is
 * a separate product (see NOTICE). components/office-kit.json pins where each of them is taken from (a published npm
 * tarball with its sha512, or a file of office-kit's GitHub release, cross-checked against that release's SHA256SUMS)
 * and every file's sha256; this script fetches them into the paths the build reads (ignored by git):
 *
 *   node tooling/kit/vendor.mjs fetch     download the pinned sources (or reuse .vendor-cache/), check them, write the files
 *   node tooling/kit/vendor.mjs check     the files on disk are exactly the pinned ones (exit 1 otherwise)
 *
 * OFFICESDK_VENDOR_DIR=<dir> reads local copies instead of downloading (still checked), named as in .vendor-cache/: an
 * npm source as `<unscoped package>-<version>.tgz`, a release file (and its SHA256SUMS) as `<source id>-<file>`.
 * tooling/suite/pack-universal.mjs refuses to build until `check` passes.
 */
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { Readable } from 'node:stream';
import { fileURLToPath } from 'node:url';
import { extractTar } from '../../packages/dsh-adapter/src/host/fetch/tar.mjs';

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
export const PIN_FILE = 'components/office-kit.json';
/** The pin format this script reads (v1, one npm tarball for every file, is gone with office-kit 0.3). */
export const PIN_VERSION = 2;

export const readPin = (repo = REPO) => JSON.parse(readFileSync(path.join(repo, PIN_FILE), 'utf8'));

/** Where the build reads the fetched files (the same paths the editor bundle and fonts always had). */
export function paths(repo = REPO) {
  const vendor = path.join(repo, 'components', 'vendor');
  return { pin: path.join(repo, PIN_FILE), kitManifest: path.join(repo, 'components', 'kit-manifest.json'), vendor: { root: vendor, editorBundle: path.join(vendor, 'editor', 'bundle'), hostFonts: path.join(vendor, 'host-fonts') } };
}

const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const sri512 = bytes => `sha512-${createHash('sha512').update(bytes).digest('base64')}`;
const unscoped = name => name.replace(/^@[^/]+\//, '');
/** A relative `/`-separated path that stays inside its root on every platform (no `..`, no `\\`, no drive or scheme). */
const SAFE_PATH = value => typeof value === 'string' && value !== '' && !/^\//.test(value) && !/[\\:]/.test(value) && !value.split('/').some(part => part === '..' || part === '');
const SOURCE_ID = /^[a-z0-9][a-z0-9-]*$/;
/** The registry tarball address of an npm source (`<registry><name>/-/<unscoped>-<version>.tgz`). */
export const tarballUrl = source => `${source.registry.replace(/\/*$/, '/')}${source.package}/-/${unscoped(source.package)}-${source.version}.tgz`;
/** The kind of a source: `npm` (registry, package, version, integrity) or `url` (a release directory). */
export const sourceKind = source => (source?.url !== undefined ? 'url' : 'npm');
/** How a source is named in messages. */
export const describe = (id, source) => (sourceKind(source) === 'url' ? `${id} (${source.url})` : `${source.package}@${source.version}`);

/** What is wrong with the pin file itself (empty when nothing). */
export function pinProblems(pin) {
  const problems = [];
  if (pin?.v !== PIN_VERSION) return [`${PIN_FILE}: v must be ${PIN_VERSION}`];
  const sources = Object.entries(pin.sources ?? {});
  if (!sources.length) problems.push(`${PIN_FILE}: sources is empty`);
  for (const [id, source] of sources) {
    if (!SOURCE_ID.test(id)) problems.push(`${PIN_FILE}: sources.${id}: an id is lower-case letters, digits and -`);
    if (sourceKind(source) === 'url') {
      if (typeof source.url !== 'string' || !/^https:\/\/[^/]+\/.*\/$/.test(source.url)) problems.push(`${PIN_FILE}: sources.${id}.url must be an https directory (ending in /)`);
      if (source.sums !== undefined && !SAFE_PATH(source.sums)) problems.push(`${PIN_FILE}: sources.${id}.sums must be a file name in that directory`);
      continue;
    }
    for (const key of ['registry', 'package', 'version', 'integrity']) if (typeof source?.[key] !== 'string' || !source[key]) problems.push(`${PIN_FILE}: sources.${id}.${key} is missing`);
    if (source?.registry && !/^https:\/\//.test(source.registry)) problems.push(`${PIN_FILE}: sources.${id}.registry must be https`);
    if (source?.integrity && !/^sha512-/.test(source.integrity)) problems.push(`${PIN_FILE}: sources.${id}.integrity must be sha512`);
  }
  const entries = Object.entries(pin.files ?? {});
  if (!entries.length) problems.push(`${PIN_FILE}: files is empty`);
  const targets = new Set();
  for (const [name, entry] of entries) {
    if (!pin.sources?.[entry?.from]) problems.push(`${PIN_FILE}: files.${name}.from must name one of sources`);
    if (!SAFE_PATH(entry?.path)) problems.push(`${PIN_FILE}: files.${name}.path must be a relative path in its source`);
    if (!/^[0-9a-f]{64}$/.test(entry?.sha256 ?? '')) problems.push(`${PIN_FILE}: files.${name}.sha256 must be 64 hex digits`);
    if (typeof entry?.to !== 'string' || !entry.to.startsWith('components/') || !SAFE_PATH(entry.to)) problems.push(`${PIN_FILE}: files.${name}.to must stay under components/`);
    else if (targets.has(entry.to)) problems.push(`${PIN_FILE}: files.${name}.to is written twice`);
    targets.add(entry?.to);
  }
  if (typeof pin.officeKit?.version !== 'string' || !pin.officeKit.version) problems.push(`${PIN_FILE}: officeKit.version is missing`);
  if (pin.officeKit?.commit !== undefined && typeof pin.officeKit.commit !== 'string') problems.push(`${PIN_FILE}: officeKit.commit must be a string when present`);
  if (!Array.isArray(pin.officeKit?.platforms)) problems.push(`${PIN_FILE}: officeKit.platforms must be a list`);
  return problems;
}

/** The fetched files against the pin: missing, changed, or extra files under components/vendor (empty when exact). */
export function vendorProblems(repo = REPO) {
  let pin;
  try { pin = readPin(repo); } catch { return [`${PIN_FILE} is missing or unreadable`]; }
  const problems = pinProblems(pin);
  if (problems.length) return problems;
  const expected = new Set();
  for (const [name, entry] of Object.entries(pin.files)) {
    const file = path.join(repo, entry.to);
    expected.add(path.resolve(file));
    if (!existsSync(file)) { problems.push(`${entry.to} is missing (node tooling/kit/vendor.mjs fetch)`); continue; }
    if (sha256(readFileSync(file)) !== entry.sha256) problems.push(`${entry.to} differs from ${describe(entry.from, pin.sources[entry.from])} ${entry.path} (${name})`);
  }
  const { vendor, kitManifest } = paths(repo);
  for (const file of filesUnder(vendor.root)) if (!expected.has(path.resolve(file))) problems.push(`${path.relative(repo, file).split(path.sep).join('/')} is not in ${PIN_FILE}`);
  if (!problems.length && existsSync(kitManifest)) {
    const manifest = JSON.parse(readFileSync(kitManifest, 'utf8'));
    // office-kit 0.3 manifests carry no commit; the pin names one only when its manifest does.
    if (manifest.officeKit?.version !== pin.officeKit.version || (manifest.officeKit?.commit ?? undefined) !== (pin.officeKit.commit ?? undefined)) problems.push(`kit-manifest.json is office-kit ${manifest.officeKit?.version}${manifest.officeKit?.commit ? ` (${manifest.officeKit.commit})` : ''}, the pin says ${pin.officeKit.version}${pin.officeKit.commit ? ` (${pin.officeKit.commit})` : ''}`);
    // The platforms the plugin can run a Kit on: those with an executable (a library-only build, such as OpenHarmony's, is not one).
    const platforms = Object.entries(manifest.kit ?? {}).filter(([, kit]) => kit?.binary).map(([name]) => name).sort();
    if (JSON.stringify(platforms) !== JSON.stringify([...pin.officeKit.platforms].sort())) problems.push(`kit-manifest.json has runnable Kits for ${platforms.join(', ')}, the pin lists ${pin.officeKit.platforms.join(', ')}`);
    // The editor page the package carries is the one office-kit's manifest names, when it names one.
    const bundle = Object.values(pin.files).find(entry => entry.to === 'components/vendor/editor/bundle/offline-editor.js');
    const named = /^sha256-/.test(manifest.editor?.bundle?.sha256 ?? '') ? Buffer.from(manifest.editor.bundle.sha256.slice(7), 'base64').toString('hex') : undefined;
    if (bundle && named && bundle.sha256 !== named) problems.push(`components/vendor/editor/bundle/offline-editor.js is not the editor bundle kit-manifest.json names (editor.bundle.sha256)`);
  }
  return problems;
}

/** Every file under `dir` (none when it does not exist). */
function filesUnder(dir) {
  if (!existsSync(dir)) return [];
  return readdirSync(dir, { withFileTypes: true }).flatMap(entry => (entry.isDirectory() ? filesUnder(path.join(dir, entry.name)) : [path.join(dir, entry.name)]));
}

/** `<hex>  <name>` lines of a SHA256SUMS file, as name → hex. */
export function parseSums(text) {
  const sums = new Map();
  for (const line of String(text).split(/\r?\n/)) { const match = /^([0-9a-f]{64})\s+\*?(.+?)\s*$/.exec(line); if (match) sums.set(match[2], match[1]); }
  return sums;
}

/**
 * Reads the pinned sources (from `dir` when given, else .vendor-cache/ when its copy is still the pinned one, else the
 * network), checks each source and each file, and only then writes the files to their `to` paths: a bad source never
 * leaves a half-written vendor tree. `fetchImpl` is for tests. Returns the files written.
 */
export async function fetchVendor({ repo = REPO, dir = process.env.OFFICESDK_VENDOR_DIR, fetchImpl = fetch, log = text => process.stderr.write(`${text}\n`) } = {}) {
  const pin = readPin(repo);
  const problems = pinProblems(pin);
  if (problems.length) throw new Error(problems.join('\n'));
  const cache = path.join(repo, '.vendor-cache');
  const download = async url => {
    log(`fetching ${url}`);
    const response = await fetchImpl(url);
    if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`);
    return Buffer.from(await response.arrayBuffer());
  };
  /** A source file's bytes: a local copy when `dir` is given; else the cache when `valid` accepts it; else downloaded and cached. */
  const read = async (name, url, valid) => {
    if (dir) return fs.readFile(path.join(dir, name));
    const cached = path.join(cache, name);
    if (existsSync(cached)) { const bytes = readFileSync(cached); if (valid(bytes)) return bytes; }
    const bytes = await download(url);
    if (valid(bytes)) { await fs.mkdir(cache, { recursive: true }); await fs.writeFile(cached, bytes); }
    return bytes;
  };
  const work = await fs.mkdtemp(path.join(repo, '.vendor-cache-'));
  try {
    const checked = [];
    const byFrom = Object.entries(pin.files).reduce((map, [name, entry]) => map.set(entry.from, [...map.get(entry.from) ?? [], [name, entry]]), new Map());
    for (const [id, entries] of byFrom) {
      const source = pin.sources[id];
      if (sourceKind(source) === 'npm') {
        const file = `${unscoped(source.package)}-${source.version}.tgz`;
        const bytes = await read(file, tarballUrl(source), bytes => sri512(bytes) === source.integrity);
        if (sri512(bytes) !== source.integrity) throw new Error(`${describe(id, source)}: the tarball's sha512 is not the pinned ${source.integrity}`);
        const into = path.join(work, id);
        const wanted = new Set(entries.map(([, entry]) => entry.path));
        const { files } = await extractTar(Readable.from([bytes]), into, { gzip: true, strip: 'package/', keep: name => wanted.has(name) });
        for (const [name, entry] of entries) {
          if (!files[entry.path]) throw new Error(`${describe(id, source)} has no ${entry.path}`);
          const content = await fs.readFile(path.join(into, ...entry.path.split('/')));
          if (sha256(content) !== entry.sha256) throw new Error(`${entry.path} in ${describe(id, source)} is not the pinned sha256 (${name})`);
          checked.push([entry.to, content]);
        }
      } else {
        // A release file: its own sha256, and (when the source names one) the release's SHA256SUMS must list the same.
        // SHA256SUMS is cached once it lists every pinned file of this source as pinned (it then checks itself).
        const listsPins = bytes => { const listed = parseSums(bytes.toString('utf8')); return entries.every(([, entry]) => listed.get(entry.path) === entry.sha256); };
        const sums = source.sums ? parseSums((await read(`${id}-${source.sums}`, `${source.url}${source.sums}`, listsPins)).toString('utf8')) : undefined;
        for (const [name, entry] of entries) {
          const content = await read(`${id}-${entry.path.replaceAll('/', '_')}`, `${source.url}${entry.path}`, bytes => sha256(bytes) === entry.sha256);
          if (sha256(content) !== entry.sha256) throw new Error(`${source.url}${entry.path} is not the pinned sha256 (${name})`);
          if (sums && sums.get(entry.path) !== entry.sha256) throw new Error(`${source.url}${source.sums} does not list ${entry.path} with the pinned sha256`);
          checked.push([entry.to, content]);
        }
      }
    }
    const { vendor, kitManifest } = paths(repo);
    await fs.rm(vendor.root, { recursive: true, force: true });
    await fs.rm(kitManifest, { force: true });
    for (const [to, content] of checked) {
      const target = path.join(repo, to);
      await fs.mkdir(path.dirname(target), { recursive: true });
      await fs.writeFile(target, content);
    }
    return checked.map(([to]) => to);
  } finally { await fs.rm(work, { recursive: true, force: true }); }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const command = process.argv[2];
  if (command === 'fetch') {
    const written = await fetchVendor();
    const problems = vendorProblems();
    if (problems.length) { console.error(problems.join('\n')); process.exit(1); }
    console.log(`fetched ${written.join(', ')}`);
  } else if (command === 'check') {
    const problems = vendorProblems();
    if (problems.length) { console.error(problems.join('\n')); process.exit(1); }
    console.log(`components/vendor and kit-manifest.json match ${PIN_FILE}`);
  } else { console.error('usage: node tooling/kit/vendor.mjs fetch|check'); process.exit(2); }
}
