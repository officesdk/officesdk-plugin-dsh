import { execFile } from 'node:child_process';
import { chmod, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import { DEFAULT_SOURCES, NPM_DEFAULTS, npmOptions, selectProfile } from './download-sources.mjs';

/**
 * office-kit downloads for the universal package (spec 261110-universal-package §3.3, §4.2, §4.3): which Kit and
 * editor resources this plugin version needs (its kit-manifest.json), where they live under the plugin data directory,
 * how a Kit is made ready to run, and the downgrade guard.
 */

export const platformKey = ({ platform = process.platform, arch = process.arch } = {}) => `${platform}-${arch}`;

export function readManifest(packageRoot) {
  try { return JSON.parse(readFileSync(join(packageRoot, 'kit-manifest.json'), 'utf8')); } catch { return undefined; }
}

/**
 * The sources one item tries, in the order its profile lists them (download-sources.json, spec 261111-npm-publish-prep
 * §6): `mirror` (`OFFICESDK_KIT_MIRROR`, a directory with office-kit's release layout; only when set), `release` (the
 * manifest's own sources: office-kit's GitHub release), `npm` (the editor only: the manifest's `editor.npm`).
 */
function sourcesOf(ids, manifest, { mirror, npm, log }) {
  const version = manifest.officeKit.version;
  const out = [];
  for (const entry of ids) {
    if (entry.id === 'mirror') { if (mirror) out.push({ id: 'mirror', kind: 'url', base: `${mirror.replace(/\/+$/, '')}/${version}/` }); }
    else if (entry.id === 'release') out.push(...manifest.sources);
    else if (entry.id === 'npm') { const source = editorNpmSource(manifest, npm, log, entry); if (source) out.push(source); }
  }
  return out;
}

/** The Kit's sources (the default profile unless `profile` says otherwise). */
export function kitSources(manifest, { mirror, profile = selectProfile(DEFAULT_SOURCES) } = {}) {
  return sourcesOf(profile.kit, manifest, { mirror });
}

/**
 * office-kit's editor package on npm (kit-manifest.json `editor.npm`), the editor's default source: NPM_DEFAULTS limits
 * (download-sources.json may set its own). `npm`: `off` leaves it out; a URL replaces the registry
 * (`OFFICESDK_EDITOR_NPM`, e.g. a registry mirror). The bytes are trusted only through the manifest's sha512, as from
 * any source.
 */
export const NPM_CONNECT_TIMEOUT_MS = NPM_DEFAULTS.connectTimeoutMs;
export const NPM_IDLE_MS = NPM_DEFAULTS.idleMs;
export function editorNpmSource(manifest, npm, log = () => {}, entry = {}) {
  const pinned = manifest.editor?.npm;
  const value = typeof npm === 'string' ? npm.trim() : '';
  if (!pinned || /^(off|false|0|no|none)$/i.test(value)) return undefined;
  const url = /^https?:\/\//i.test(value);
  if (value && !url) log('warn', `OFFICESDK_EDITOR_NPM=${value} is neither off nor an http(s) URL; the manifest's registry is used`);
  const registry = url ? value : pinned.registry;
  return { id: 'npm', kind: 'npm', registry, package: pinned.package, version: pinned.version, ...npmOptions(entry) };
}

const run = (file, args, { timeout }) => new Promise((resolve, reject) => execFile(file, args, { timeout, windowsHide: true, maxBuffer: 1 << 20 }, (error, stdout, stderr) => (error ? reject(Object.assign(error, { stderr: String(stderr) })) : resolve(String(stdout)))));

/**
 * Makes an unpacked Kit runnable and proves it: executable bits; on macOS no `com.apple.quarantine` (a Kit copied in
 * from a browser download would otherwise be stopped by Gatekeeper, §1.4); then `--version --json` must name this
 * manifest's version and protocol. Windows gets a longer first start (Defender scans a new exe).
 */
export function kitPreparer(manifest, platform, { runFile = run, host = process.platform } = {}) {
  const entry = manifest.kit[platform];
  return async dir => {
    for (const [name, file] of Object.entries(entry.files)) if (file.executable) await chmod(join(dir, ...name.split('/')), 0o755);
    if (host === 'darwin') await runFile('xattr', ['-dr', 'com.apple.quarantine', dir], { timeout: 30_000 }).catch(() => {});
    let info;
    try { info = JSON.parse(await runFile(join(dir, entry.binary), ['--version', '--json'], { timeout: host === 'win32' ? 120_000 : 30_000 })); }
    // Not retried by downloading again: the files were already checked; a slow first start (Defender) or a broken
    // machine is not fixed by the same bytes. The next open runs the check again.
    catch (error) { throw Object.assign(new Error(`the document engine does not start (${error?.code ?? ''} ${String(error?.stderr ?? error?.message ?? error).trim().slice(0, 200)})`), { code: 'kit-selftest', retry: false, final: true }); }
    if (info?.version !== manifest.officeKit.version || info?.compat?.protocolVersion !== manifest.officeKit.protocol) {
      throw Object.assign(new Error(`the document engine reports ${info?.version}/protocol ${info?.compat?.protocolVersion}, expected ${manifest.officeKit.version}/protocol ${manifest.officeKit.protocol}`), { code: 'kit-selftest', retry: false, final: true });
    }
  };
}

/**
 * The two download items of this plugin version on this machine, or `{ unsupported }` when the manifest has no Kit for
 * it. `root` is `<data>/office-kit`; `packageRoot/resources/{kit,office-kit}` is used in place when an offline package
 * carries them.
 */
export function officeKitItems(manifest, { root, packageRoot, host = {}, mirror, npm, runFile, prepare, log, profile = selectProfile(DEFAULT_SOURCES) } = {}) {
  const platform = platformKey(host);
  const version = manifest.officeKit.version;
  const kit = manifest.kit?.[platform];
  const sources = sourcesOf(profile.kit, manifest, { mirror, log });
  // The editor: in its profile's order (a mirror when set, then npm). The Kit never comes from npm.
  const editorSources = sourcesOf(profile.editor, manifest, { mirror, npm, log });
  const editor = {
    key: `editor/${version}`, label: 'office-kit editor resources', dir: join(root, 'editor', version),
    archive: { file: manifest.editor.file, size: manifest.editor.size, integrity: manifest.editor.integrity },
    files: manifest.editor.files, strip: manifest.editor.root, outside: 'skip', sources: editorSources,
    bundledDir: packageRoot && join(packageRoot, 'resources', 'office-kit'),
  };
  // What the manifest does have, for the message (test platforms are named as such).
  const supported = Object.entries(manifest.kit ?? {}).map(([name, row]) => ({ platform: name, status: row.status }));
  // A platform office-kit builds only a library for (no executable, e.g. OpenHarmony) is not one the plugin can run.
  if (!kit?.binary) return { platform, editor, unsupported: { platform, reason: 'platform', supported: supported.filter(row => manifest.kit[row.platform]?.binary) } };
  return {
    platform, editor,
    kit: {
      key: `kit/${version}/${platform}`, label: `office-kit ${version} (${platform})`, dir: join(root, 'kit', version, platform),
      archive: { file: kit.file, size: kit.size, integrity: kit.integrity }, files: kit.files, sources,
      bundledDir: packageRoot && join(packageRoot, 'resources', 'kit'),
      binary: kit.binary, status: kit.status,
      prepare: prepare ?? kitPreparer(manifest, platform, { runFile, host: host.platform ?? process.platform }),
    },
  };
}

const semver = version => /^(\d+)\.(\d+)\.(\d+)/.exec(version ?? '')?.slice(1).map(Number);
export const versionBelow = (a, b) => {
  const [x, y] = [semver(a), semver(b)];
  if (!x || !y) return false;
  for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] < y[i];
  return false;
};

/**
 * §4.3 downgrade guard. `highest` is what the newest Kit that ever ran here recorded: `{ version, indexVersion,
 * downgradeFloor }`. A Kit may start when its version is at least that floor and it reads that index version.
 */
export function downgradeProblem(manifest, highest) {
  if (!highest?.version) return undefined;
  const kit = manifest.officeKit;
  const reads = kit.readsIndexVersions ?? [];
  if ((highest.downgradeFloor && versionBelow(kit.version, highest.downgradeFloor)) || (highest.indexVersion !== undefined && reads.length && !reads.includes(highest.indexVersion))) {
    return { version: highest.version, message: `这台机器用过更新的文档引擎 ${highest.version}，当前插件的引擎 ${kit.version} 不能再打开它整理过的文档。请更新插件（或安装使用同一引擎版本的插件）。This machine has run the newer document engine ${highest.version}; this plugin's engine ${kit.version} cannot read what it wrote. Update the plugin.` };
  }
  return undefined;
}

/** `<root>/highest.json`, written after a Kit started and opened its work directory (never before). */
export async function readHighest(root) { try { return JSON.parse(await readFile(join(root, 'highest.json'), 'utf8')); } catch { return undefined; } }
export async function recordHighest(root, manifest) {
  const current = await readHighest(root);
  const kit = manifest.officeKit;
  if (current?.version && !versionBelow(current.version, kit.version)) return current;
  const next = { version: kit.version, indexVersion: kit.indexVersion, downgradeFloor: kit.downgradeFloor };
  await mkdir(root, { recursive: true });
  const tmp = join(root, `highest.json.${randomUUID()}.tmp`);
  await writeFile(tmp, JSON.stringify(next, null, 2));
  await rename(tmp, join(root, 'highest.json'));
  return next;
}
