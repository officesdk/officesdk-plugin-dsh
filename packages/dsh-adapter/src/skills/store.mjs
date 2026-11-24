import { randomUUID } from 'node:crypto';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { mkdir, rename, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { Downloader, COMPLETE } from '../host/fetch/downloader.mjs';
import { CATALOG_FORMAT, availableVersions, skillFields } from './catalog.mjs';

/**
 * Skill resources on demand (spec 261110-universal-package §4.5). The package ships each skill's SKILL.md (the light
 * layer) and skills-catalog.json; the rest of a skill's directory (style packs, templates, scripts, examples) downloads
 * on first use into `<data>/skills/<name>/<version>/`, through the same downloader as office-kit. The skills
 * component (DSH's provider) and the host component (style and template tools) each hold one store over the same
 * directory; they agree through the files there (`.complete`, `settings.json`, `catalog-cache.json`, leases).
 *
 * Without a packaged catalog (the repository itself, tests) every skill is its package directory, as before.
 */
/** The skills ship in the package unless a download-sources.json names a skill mirror (`public`, drills). */
export const DEFAULT_SKILL_SOURCES = Object.freeze({ sources: ['bundled'], catalog: ['bundled'] });
const PUBLIC_CONNECT_TIMEOUT_MS = 5000;
const LEASE_RENEW_MS = 10_000;
const LEASE_STALE_MS = 60_000;
const readJson = file => { try { return JSON.parse(readFileSync(file, 'utf8')); } catch { return undefined; } };
const writeJson = async (file, value) => {
  const tmp = `${file}.${randomUUID()}.tmp`;
  await writeFile(tmp, `${JSON.stringify(value, null, 2)}\n`);
  await rename(tmp, file);
};

export class SkillStore {
  /**
   * `sources`: the profile's skill sources (download-sources.json, spec 261111-npm-publish-prep §6): `{ sources, catalog,
   * publicBase }` with ids `public` (a skill mirror at `publicBase`) and `bundled` (in the package: no source).
   */
  constructor({ packageRoot, dataRoot, pluginVersion, names, downloader, log = () => {}, now = Date.now, fetchImpl, sources = DEFAULT_SKILL_SOURCES }) {
    Object.assign(this, { packageRoot, pluginVersion, names, log, now, fetchImpl });
    this.sourceIds = [...new Set(sources.sources ?? [])];
    this.catalogIds = [...new Set(sources.catalog ?? [])];
    // An empty public base means no public mirror yet: the source is left out rather than tried and timed out.
    this.publicBase = (sources.publicBase ?? '').replace(/\/+$/, '');
    this.root = join(dataRoot, 'skills');
    this.packaged = readJson(join(packageRoot, 'skills-catalog.json'));
    this.dev = !this.packaged;
    // A package lists the skills it has (the public build ships two of the three, spec 261113-bundle-public-skills): one
    // its catalog does not name is not registered, shown, downloaded or pruned here.
    if (this.packaged?.skills && names) this.names = names.filter(name => Object.hasOwn(this.packaged.skills, name));
    this.downloader = downloader ?? new Downloader({ root: this.root, log });
    this.pins = new Map();
    this.leaseFile = join(this.root, 'leases', `${process.pid}-${randomUUID()}.json`);
  }

  get settingsFile() { return join(this.root, 'settings.json'); }
  get cacheFile() { return join(this.root, 'catalog-cache.json'); }
  settings() { return readJson(this.settingsFile) ?? { disabled: [] }; }
  enabled(name) { return !(this.settings().disabled ?? []).includes(name); }
  async setEnabled(name, on) {
    await mkdir(this.root, { recursive: true });
    const settings = this.settings();
    const disabled = new Set(settings.disabled ?? []);
    if (on) disabled.delete(name); else disabled.add(name);
    await writeJson(this.settingsFile, { ...settings, disabled: [...disabled].sort() });
  }
  remoteCatalog() { return this.packaged?.remoteCatalog === false ? undefined : readJson(this.cacheFile); }

  /** The package's own copy of a skill: complete in an offline/bridge package (`layer: full`) and in the repository. */
  packageDir(name) { return join(this.packageRoot, 'skills', name); }
  bundledFull() { return this.dev || this.packaged?.layer === 'full'; }
  versionDir(name, version) { return join(this.root, name, version); }

  /** Available versions, highest first (catalog.mjs availableVersions; the plugin's own when nothing else is). */
  versions(name) {
    if (this.dev) return [];
    const list = availableVersions(name, { packaged: this.packaged, remote: this.remoteCatalog(), pluginVersion: this.pluginVersion });
    return list.length ? list : (this.packaged.skills?.[name] ?? []).map(entry => ({ ...entry, packaged: true }));
  }

  /**
   * The download item of a skill version. Source: a skill mirror (`public`), `<base>/<name>-<version>/<pack file>`,
   * pinned by the pack's sha512; every file is checked.
   */
  item(name, entry) {
    const pack = entry.pack ?? this.remoteCatalog()?.skills?.[name]?.find(item => item.version === entry.version)?.pack;
    const genericFile = `${name}-${entry.version}.tgz`;
    const sources = [];
    for (const id of this.sourceIds) {
      if (id === 'public' && this.publicBase && pack) sources.push({ id: 'public', kind: 'url', base: `${this.publicBase}/${encodeURIComponent(`${name}-${entry.version}`)}/`, connectTimeoutMs: PUBLIC_CONNECT_TIMEOUT_MS });
    }
    return {
      key: `skill/${name}/${entry.version}`, label: `${name} ${entry.version}`, dir: this.versionDir(name, entry.version), files: entry.files,
      file: genericFile,
      ...pack && { archive: { file: pack.file, size: pack.size, integrity: pack.integrity } },
      bundledDir: entry.packaged && this.bundledFull() ? this.packageDir(name) : undefined, sources,
    };
  }

  /** Whether `entry`'s files are here (downloaded, or carried by the package). */
  ready(name, entry) {
    const item = this.item(name, entry);
    if (item.bundledDir && Object.keys(item.files).every(file => existsSync(join(item.bundledDir, ...file.split('/'))))) return true;
    return this.downloader.isComplete(item);
  }

  /**
   * `{ version, dir?, entry?, selected }`: the effective version (spec §4.5) — the selected version when its files are
   * here, else the highest available version whose files are, else the packaged one (dir only when it is complete).
   */
  effective(name) {
    if (this.dev) return { version: skillFields(readFileSync(join(this.packageDir(name), 'SKILL.md'), 'utf8')).version, dir: this.packageDir(name), dev: true };
    const versions = this.versions(name);
    const selected = versions[0];
    const here = versions.find(entry => this.ready(name, entry));
    const entry = here ?? versions.find(item => item.packaged) ?? selected;
    if (!entry) return { version: undefined, selected: undefined };
    const ready = Boolean(here);
    const item = this.item(name, entry);
    const dir = !ready ? undefined : item.bundledDir && Object.keys(item.files).every(file => existsSync(join(item.bundledDir, ...file.split('/')))) ? item.bundledDir : item.dir;
    return { version: entry.version, entry, dir, selected: selected?.version, ready };
  }

  /** SKILL.md text of the effective version: its own file when here, else the package's. */
  skillText(name, effective = this.effective(name)) {
    const file = effective.dir ? join(effective.dir, 'SKILL.md') : join(this.packageDir(name), 'SKILL.md');
    return readFileSync(file, 'utf8');
  }

  /** Downloads `version` (default: the selected one). Options as Downloader.ensure (`budgetMs`, `signal`, `firstSourceOnly`). */
  async ensure(name, { version, ...options } = {}) {
    if (this.dev) return { dir: this.packageDir(name), version: this.effective(name).version };
    const versions = this.versions(name);
    const entry = version ? versions.find(item => item.version === version) ?? this.remoteEntry(name, version) : versions[0];
    if (!entry) throw Object.assign(new Error(`${name} ${version ?? ''} is not in the skill catalog`), { code: 'not-found' });
    let item = this.item(name, entry);
    // The public mirror serves packs only, pinned by the catalog: before the first catalog arrived there is nothing to
    // ask, so fetch it now instead of failing with no source (smoke 261111).
    if (!item.sources.length && this.sourceIds.includes('public') && this.publicBase && !this.remoteCatalog()) {
      await this.refresh({ timeoutMs: 15_000 }).catch(error => this.log(`skill catalog: ${error?.message}`));
      item = this.item(name, entry);
    }
    const got = await this.downloader.ensure(item, options);
    await this.touchLease();
    return { dir: got.dir, version: entry.version, source: got.source };
  }
  remoteEntry(name, version) { return [...(this.packaged?.skills?.[name] ?? []), ...(this.remoteCatalog()?.skills?.[name] ?? [])].find(entry => entry.version === version); }

  /**
   * The directory a host tool uses in session `sessionId` (spec §4.5 point 3): the first call in a session pins the
   * effective version, later calls reuse it; a pinned version that lost its files falls back to the effective one.
   * Downloads when nothing is here yet (`budgetMs`, default 120 s).
   */
  async resolve(name, sessionId, { budgetMs = 120_000, signal } = {}) {
    if (this.dev) return this.packageDir(name);
    const key = `${sessionId ?? ''}\0${name}`;
    const pinned = this.pins.get(key);
    if (pinned && existsSync(join(pinned.dir, 'SKILL.md'))) return pinned.dir;
    let effective = this.effective(name);
    if (!effective.ready) {
      await this.ensure(name, { budgetMs, signal, version: effective.selected });
      effective = this.effective(name);
    }
    if (!effective.dir) throw Object.assign(new Error(`${name} resources are not available yet`), { code: 'unavailable' });
    this.pins.set(key, { version: effective.version, dir: effective.dir });
    await this.touchLease();
    return effective.dir;
  }
  /** The effective directory now, without waiting (contexts built synchronously); undefined when nothing is here. */
  resolveNow(name, sessionId) {
    if (this.dev) return this.packageDir(name);
    const pinned = this.pins.get(`${sessionId ?? ''}\0${name}`);
    if (pinned && existsSync(join(pinned.dir, 'SKILL.md'))) return pinned.dir;
    return this.effective(name).dir;
  }
  /** Re-pins a session to the effective version (a style or template the pinned version lacks, spec §4.5 J1). */
  repin(name, sessionId) {
    const effective = this.effective(name);
    if (effective.dir) this.pins.set(`${sessionId ?? ''}\0${name}`, { version: effective.version, dir: effective.dir });
    return effective.dir;
  }

  /** Fetches the skill mirror's catalog into catalog-cache.json; returns whether any selected version changed. */
  async refresh({ timeoutMs = 60_000 } = {}) {
    if (this.dev || this.packaged?.remoteCatalog === false) return false;
    const before = Object.fromEntries((this.names ?? []).map(name => [name, this.versions(name)[0]?.version]));
    const text = await this.fetchCatalog(timeoutMs);
    const catalog = JSON.parse(text);
    if (catalog?.v !== CATALOG_FORMAT || typeof catalog.skills !== 'object') throw new Error('the remote skill catalog has an unknown format');
    await mkdir(this.root, { recursive: true });
    await writeJson(this.cacheFile, catalog);
    return (this.names ?? []).some(name => this.versions(name)[0]?.version !== before[name]);
  }
  /** The newest catalog from the profile's catalog sources: a skill mirror (`public`, `<base>/catalog/skills-catalog.json`). */
  async fetchCatalog(timeoutMs) {
    const errors = [];
    for (const id of this.catalogIds) {
      try {
        if (id === 'public') {
          if (!this.publicBase) continue;
          const response = await (this.fetchImpl ?? fetch)(`${this.publicBase}/catalog/skills-catalog.json`, { signal: AbortSignal.timeout(Math.min(timeoutMs, 30_000)) });
          if (!response.ok) throw new Error(`public catalog: HTTP ${response.status}`);
          return await response.text();
        }
      } catch (error) { errors.push(error?.message ?? String(error)); }
    }
    throw new Error(errors.join('; ') || 'no skill catalog source');
  }

  /** Per-process lease listing the versions this process pinned; other processes' clear-cache keeps them. */
  async touchLease() {
    try {
      await mkdir(join(this.root, 'leases'), { recursive: true });
      await writeJson(this.leaseFile, { pid: process.pid, at: this.now(), dirs: [...new Set([...this.pins.values()].map(pin => pin.dir))] });
    } catch { /* best effort */ }
  }
  startLease() {
    this.leaseTimer = setInterval(() => { void this.touchLease(); }, LEASE_RENEW_MS);
    this.leaseTimer.unref?.();
  }
  async dispose() { clearInterval(this.leaseTimer); await rm(this.leaseFile, { force: true }).catch(() => {}); }
  leasedDirs() {
    const dirs = new Set([...this.pins.values()].map(pin => pin.dir));
    let names = [];
    try { names = readdirSync(join(this.root, 'leases')); } catch {}
    for (const name of names) {
      const file = join(this.root, 'leases', name);
      // This process's own pins are read from memory (they may have changed since the lease file was written).
      if (file === this.leaseFile) continue;
      try { if (this.now() - statSync(file).mtimeMs > LEASE_STALE_MS) continue; } catch { continue; }
      for (const dir of readJson(file)?.dirs ?? []) dirs.add(dir);
    }
    return dirs;
  }

  /**
   * Clear cache (settings): every downloaded version of `name` except those leased by a process or downloading, then the
   * effective version downloads again in the background. Returns the removed directories.
   */
  async clear(name) {
    if (this.dev) return [];
    const keep = this.leasedDirs();
    const removed = [];
    let versions = [];
    try { versions = readdirSync(join(this.root, name)).filter(version => !version.startsWith('.')); } catch {}
    for (const version of versions) {
      const dir = this.versionDir(name, version);
      if (keep.has(dir) || this.downloader.inflight.has(`skill/${name}/${version}`)) continue;
      await rm(dir, { recursive: true, force: true }).then(() => removed.push(dir), () => {});
    }
    void this.ensure(name).catch(error => this.log('warn', `${name}: ${error.message}`));
    return removed;
  }

  /** Automatic cleanup: versions neither effective, selected, leased, nor used within 30 days. */
  async prune() {
    if (this.dev) return [];
    const keep = this.leasedDirs();
    for (const name of this.names ?? []) {
      const effective = this.effective(name);
      for (const version of [effective.version, effective.selected]) if (version) keep.add(this.versionDir(name, version));
    }
    const dirs = [];
    for (const name of this.names ?? []) { try { for (const version of readdirSync(join(this.root, name))) if (!version.startsWith('.')) dirs.push(this.versionDir(name, version)); } catch {} }
    return this.downloader.prune({ dirs, keep: dir => keep.has(dir) });
  }

  /** One row per skill for the settings panel. */
  status() {
    return (this.names ?? []).map(name => {
      const effective = this.effective(name);
      const progress = this.downloader.inflight.has(`skill/${name}/${effective.selected}`) ? 'downloading' : undefined;
      const failure = effective.selected ? this.downloader.lastError(`skill/${name}/${effective.selected}`) : undefined;
      return { name, enabled: this.enabled(name), version: effective.version, latest: effective.selected ?? effective.version,
        state: this.dev ? 'bundled' : progress ?? (effective.ready && effective.version === effective.selected ? 'ready' : failure ? 'failed' : effective.ready ? 'update' : 'missing'),
        ...failure && { error: failure.message } };
    });
  }
}

export { COMPLETE };
