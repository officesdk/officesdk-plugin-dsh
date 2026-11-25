import { execFile } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve, win32 } from 'node:path';
import { ChannelStore, channelVersion, fetchPackument, npmInstall, setScopedRegistry } from './npm.mjs';
import { newer as newerVersion, parseVersion } from './version.mjs';

// How long a release check counts: matches the panel's background check, so each one asks the registry again.
const CHECK_TTL_MS = 600_000;
const RELOAD_MARKER = /^# officesdk: loaded .*\n?/m;

// X.Y.Z, a preview X.Y.Z-alpha|beta|rc.N, or a test build X.Y.Z-test.N; all of them come before X.Y.Z itself
// (version.mjs).
export const parse = parseVersion;
export const newer = newerVersion;
/** One change to a DSH profile at a time in this process (review R1-F6): updates rewrite the profile's package.json and lock through `dsh plugin`. */
let profileQueue = Promise.resolve();
export function withProfileLock(task) {
  const run = profileQueue.then(task, task);
  profileQueue = run.catch(() => {});
  return run;
}

/** Drops `user:token@` from URLs in text shown to the client (a registry URL may carry credentials). */
export const redactCredentials = text => String(text ?? '').replace(/(\/\/)[^/@\s]+@/g, '$1');

const readVersion = file => { try { return JSON.parse(readFileSync(file, 'utf8')).version; } catch { return undefined; } };

// windowsHide: the DSH command line runs from a desktop app, so no console window may flash up on Windows.
export function runFile(file, args, { cwd, timeout = 120_000, env, windowsVerbatimArguments } = {}) {
  return new Promise((done, fail) => execFile(file, args, { cwd, timeout, maxBuffer: 4 << 20, env: { ...process.env, GIT_TERMINAL_PROMPT: '0', ...env }, windowsHide: true, windowsVerbatimArguments },
    (error, stdout, stderr) => error ? fail(Object.assign(error, { stderr: String(stderr ?? '') })) : done(String(stdout))));
}

/** The desktop app's CLI entry inside its archive, as its own `bin/dsh` launcher runs it (ELECTRON_RUN_AS_NODE). */
const DESKTOP_CLI = ['app.asar', 'dsh', 'node_modules', '@deepseek-ai', 'dsh-desktop-host', 'lib', 'cli.js'];

/**
 * How to run the DSH command line next to the running app: `{ file, args }` (args go before ours), with `env`, and
 * `cmd: true` for a Windows `.cmd` launcher. macOS runs the app's `bin/dsh` script (or `dsh` on PATH). Windows runs
 * `bin/dsh.cmd`, which needs cmd.exe (execFile cannot start a .cmd by itself); its `fallback` is the app's own
 * executable running the CLI entry as Node — that ties us to DSH's internal layout, so it is used only when the
 * launcher is missing or our arguments cannot pass through cmd.exe unchanged (see cliCommand).
 */
export function findCli({ env = process.env, platform = process.platform, resourcesPath = process.resourcesPath, execPath = process.execPath, exists = existsSync } = {}) {
  if (platform !== 'win32') {
    const candidates = [env.DSH_CLI, resourcesPath && join(resourcesPath, 'runtime', 'cli', 'bin', 'dsh'), '/Applications/DeepSeek Harness.app/Contents/Resources/runtime/cli/bin/dsh'];
    return { file: candidates.find(path => path && exists(path)) ?? 'dsh', args: [] };
  }
  const direct = resourcesPath && execPath ? { file: execPath, args: ['--expose-internals', win32.join(resourcesPath, ...DESKTOP_CLI)], env: { ELECTRON_RUN_AS_NODE: '1' } } : undefined;
  if (env.DSH_CLI) return /\.(cmd|bat)$/i.test(env.DSH_CLI) ? { file: env.DSH_CLI, args: [], cmd: true, fallback: direct } : { file: env.DSH_CLI, args: [] };
  const launcher = resourcesPath && win32.join(resourcesPath, 'runtime', 'cli', 'bin', 'dsh.cmd');
  if (launcher && exists(launcher)) return { file: launcher, args: [], cmd: true, fallback: direct };
  return direct ?? { file: 'dsh.cmd', args: [], cmd: true };
}

// What passes through cmd.exe unchanged, even to a launcher reading %1…: no spaces, quotes, %, !, ^, &, |, <, >,
// parentheses, or the ',' ';' '=' cmd.exe also splits arguments on.
const CMD_SAFE = /^[A-Za-z0-9_@.\/:+-]+$/;

/**
 * The process to start for `dsh <args>`: `{ file, args, env?, windowsVerbatimArguments? }`. A .cmd launcher runs as
 * `cmd.exe /d /s /c ""<launcher>" <args>"`; cmd.exe parses that line again, so the launcher path may hold no quote,
 * %, ! or ^ and each argument only plain characters — otherwise the direct entry runs instead, or nothing does.
 */
export function cliCommand(cli, args, { comspec = process.env.ComSpec || 'cmd.exe' } = {}) {
  if (typeof cli === 'string') cli = { file: cli, args: [] };
  const all = [...cli.args, ...args];
  if (!cli.cmd) return { file: cli.file, args: all, ...(cli.env && { env: cli.env }) };
  if (/["%!^\r\n]/.test(cli.file) || !all.every(arg => CMD_SAFE.test(arg))) {
    if (cli.fallback) return cliCommand(cli.fallback, args, { comspec });
    throw Object.assign(new Error('The DSH command line cannot be run safely here.'), { code: 'unsupported-operation' });
  }
  return { file: comspec, args: ['/d', '/s', '/c', `""${cli.file}" ${all.join(' ')}"`], windowsVerbatimArguments: true };
}

/**
 * Installs an update with Kit out of its way (spec §4.3). Every Kit view, in every window and for every principal,
 * must have nothing unsaved and no AI writing into it — otherwise the update is refused (`unsaved-changes`): the
 * update reloads them all. On Windows the running Kit executables cannot be replaced, so Kit is stopped and kept
 * stopped for the update; a failed update lets it start again (the next request starts it), as does a successful
 * one that DSH does not load until a restart. One that reloads the plugin leaves this instance's Kit stopped.
 */
export async function updateWithKit({ updater, views = [], kit, platform = process.platform, target }) {
  const unsaved = [...views].filter(view => view.unsaved || view.aiActive).length;
  if (unsaved) throw Object.assign(new Error(`${unsaved} open document${unsaved === 1 ? ' has' : 's have'} unsaved changes (perhaps in another window); save ${unsaved === 1 ? 'it' : 'them'} before updating.`), { code: 'unsaved-changes' });
  // Whatever would refuse the update anyway (not an npm install, no safe way to run the CLI) refuses before Kit stops.
  updater.command?.();
  const hold = platform === 'win32' && typeof kit?.suspend === 'function';
  if (hold) await kit.suspend();
  let result;
  try { result = await updater.update(target); }
  catch (error) { if (hold) kit.resume(); throw error; }
  if (hold && !result?.reloading) kit.resume();
  return result;
}

/**
 * Checks the npm channel for a newer release and updates through `dsh plugin add <name>@<version>`, which takes the
 * profile's package lock. With DSH's HMR the new copy loads right away (the open tabs reload); without it, on the next
 * restart.
 */
export class PluginUpdater {
  /**
   * `cli`: findCli()'s answer (a bare path string also works). An npm install (spec 261111-npm-publish-prep §3) checks its
   * channel's dist-tag and installs that exact version. `dataRoot` keeps each profile's npm channel.
   */
  constructor({ packageRoot, source, cli = findCli(), run = runFile, now = Date.now, ttlMs = CHECK_TTL_MS, dataRoot, env = process.env, fetchImpl }) {
    Object.assign(this, { cli, run, now, ttlMs, env, fetchImpl });
    // Read again each time: the profile may change under this running copy.
    this.sourceOf = source !== undefined ? () => source : () => npmInstall(packageRoot);
    this.packageDir = packageRoot && resolve(packageRoot);
    this.current = readVersion(join(packageRoot, 'package.json'));
    this.channels = new ChannelStore(dataRoot);
  }

  get source() { return this.sourceOf(); }

  /**
   * `{ current, latest?, releasedAt?, available, installed?, channel? }`; `installed` is a newer copy already on disk,
   * waiting for a restart; `releasedAt` is when `latest` was published.
   */
  async status({ refresh = false } = {}) {
    const source = this.source;
    if (!source) return { current: this.current, available: false };
    const installed = readVersion(join(source.packageDir, 'package.json'));
    const pending = newer(installed, this.current) ? installed : undefined;
    if (refresh || !this.checked || this.now() - this.checked.at > this.ttlMs) {
      this.checking ??= this.latest().then(latest => { this.checked = { at: this.now(), latest, releasedAt: latest !== undefined ? this.published?.[latest] : undefined }; }).finally(() => { this.checking = undefined; });
      await this.checking.catch(() => {});
    }
    const latest = this.checked?.latest;
    const releasedAt = this.checked?.releasedAt;
    const entry = this.registry?.packument?.versions?.[latest]?.officeSdkReleaseNotes;
    const validList = list => Array.isArray(list) && list.length > 0 && list.every(text => typeof text === 'string' && text.trim());
    const notes = latest && entry?.version === latest && validList(entry.highlights) && validList(entry.highlightsEn) && entry.highlights.length === entry.highlightsEn.length
      ? { highlights: entry.highlights, highlightsEn: entry.highlightsEn } : {};
    return { current: this.current, latest, ...notes, ...releasedAt && { releasedAt }, available: !pending && newer(latest, this.current), ...pending && { installed: pending }, channel: this.channel };
  }

  /** The newest version on offer: the npm channel's version. */
  async latest() { return this.npmLatest(this.source); }

  async npmLatest(source) {
    this.channel = await this.channels.resolve(source.profileDir, this.current, this.env);
    const answer = await fetchPackument({ profileDir: source.profileDir, env: this.env, full: true, ...this.fetchImpl && { fetchImpl: this.fetchImpl } });
    this.registry = answer;
    // Publish times by version, for the update bubble's date; a registry without them shows no date.
    const time = answer.packument?.time;
    this.published = time && typeof time === 'object' ? Object.fromEntries(Object.entries(time).filter(([, at]) => typeof at === 'string' && !Number.isNaN(Date.parse(at)))) : undefined;
    const version = channelVersion(answer.packument, this.channel, newer);
    return parse(version) ? version : undefined;
  }

  /** The `dsh plugin` process to start, or why this copy cannot update (`unsupported-operation`). */
  command(target) {
    const source = this.source;
    if (!source) throw Object.assign(new Error('This copy was not installed from npm, so it cannot update itself.'), { code: 'unsupported-operation' });
    return cliCommand(this.cli, ['plugin', '--profile', source.profile, 'add', `${source.dependencies[0]}@${target ?? 'latest'}`]);
  }

  /**
   * Installs the newest release into the profile; DSH's HMR loads it, or the next restart does. It takes `target` (the
   * version the panel showed), else the version its last check found, else a fresh check's; only a newer one unless
   * `downgrade`.
   */
  async update(target, { downgrade = false } = {}) {
    const source = this.source;
    let version;
    if (source) {
      version = target ?? this.checked?.latest ?? await this.latest();
      if (!parse(version)) throw Object.assign(new Error('No release of the plugin was found on npm.'), { code: 'unavailable' });
      // An update only moves forward; going back is `dsh plugin add <name>@<version>` (or `downgrade`), never a click (R2-F9).
      if (!downgrade && !newer(version, this.current)) throw Object.assign(new Error(`${version} is not newer than the installed ${this.current}.`), { code: 'unavailable' });
      // A registry the user chose is written to the profile (scoped) so DSH's own installs find the package later.
      if (this.registry?.from === 'env') await setScopedRegistry(source.profileDir, this.registry.registry);
    }
    const { file, args, ...options } = this.command(version);
    this.updating ??= withProfileLock(() => this.run(file, args, { cwd: source.profileDir, timeout: 600_000, ...options })).finally(() => { this.updating = undefined; });
    await this.updating;
    const after = this.source ?? source;
    const installed = readVersion(join(after.packageDir, 'package.json'));
    this.checked = undefined;
    // Recomposing reloads changed files in place; a copy in another directory (the profile was renamed under this
    // running copy) loads only at the next start.
    const inPlace = !this.packageDir || after.packageDir === this.packageDir;
    const reloading = newer(installed, this.current) && inPlace && this.reloadProfile(installed);
    return { current: this.current, installed, restartRequired: newer(installed, this.current) && !reloading, ...reloading && { reloading } };
  }

  /**
   * A package update alone does not recompose a running DSH profile; a changed profile patch does, and the
   * recomposition replaces plugins whose files changed on disk. One marker comment line, replaced per release,
   * changes the patch text without changing any setting. Returns whether the patch was touched.
   */
  reloadProfile(version, { profileDir = this.source?.profileDir, tag } = {}) { return touchReloadMarker(profileDir, version, tag); }

}

/**
 * Touches the profile patch's marker line (`# officesdk: loaded <version>[ <tag>]`) so DSH recomposes the profile and
 * loads the plugin files that changed on disk. Returns whether the patch changed.
 */
export function touchReloadMarker(profileDir, version, tag) {
  const patch = join(profileDir, 'cordis.patch.yml');
  let text;
  try { text = readFileSync(patch, 'utf8'); } catch { return false; }
  const marker = `# officesdk: loaded ${version}${tag ? ` ${tag}` : ''}\n`;
  const next = RELOAD_MARKER.test(text) ? text.replace(RELOAD_MARKER, marker) : `${text}${text.endsWith('\n') || !text ? '' : '\n'}${marker}`;
  if (next === text) return false;
  try { writeFileSync(patch, next); return true; } catch { return false; }
}
