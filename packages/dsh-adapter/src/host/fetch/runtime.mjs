import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Downloader } from './downloader.mjs';
import { downgradeProblem, officeKitItems, readHighest, recordHighest } from './office-kit.mjs';

/**
 * The document engine of the universal package (spec 261110-universal-package §4.2–§4.3): office-kit's Kit and editor
 * resources for this plugin version, downloaded on first need (or right after start, in the background), checked,
 * and handed to KitProcess as a path. Progress and failures are kept for the Client (`status()`).
 */
const PREFETCH_DELAY_MS = 5000;
const mb = bytes => (bytes / 1048576).toFixed(0);
/** Whether the package carries every file of `item` (the same test `Downloader.ensure` uses to take the packaged copy). */
const packaged = item => Boolean(item?.bundledDir && Object.keys(item.files).every(name => existsSync(join(item.bundledDir, ...name.split('/')))));

export class OfficeKitRuntime {
  constructor({ manifest, dataRoot, packageRoot, host = {}, mirror, npm, profile, downloader, log = () => {}, runFile }) {
    this.manifest = manifest;
    this.root = join(dataRoot, 'office-kit');
    this.log = log;
    this.downloader = downloader ?? new Downloader({ root: this.root, log });
    const items = officeKitItems(manifest, { root: this.root, packageRoot, host, mirror, npm, runFile, log, ...profile && { profile } });
    Object.assign(this, { kitItem: items.kit, editorItem: items.editor, platform: items.platform, unsupported: items.unsupported });
    this.progress = {};
    this.lifetime = new AbortController();
    this.downloader.on('progress', event => {
      const previous = this.progress[event.key] ?? {};
      this.progress[event.key] = { ...previous, ...event, ...event.phase !== 'failed' && { error: undefined } };
    });
  }

  /**
   * Whether both items are already here (packaged — the editor resources in every package since 261112-bundle-editor, the
   * Kit in an offline package — or an earlier download): opening a document will not wait.
   */
  ready() {
    const here = item => item && (this.packaged(item) || this.downloader.isComplete(item));
    return Boolean(!this.unsupported && here(this.kitItem) && here(this.editorItem));
  }

  /**
   * Whether the package carries `item` (1041 files for the editor): checked again by binary() and prune(), remembered
   * for status() and ready(), which the Client polls while a document waits.
   */
  packaged(item, { fresh = false } = {}) {
    this.packagedSeen ??= new Map();
    if (fresh || !this.packagedSeen.has(item)) this.packagedSeen.set(item, packaged(item));
    return this.packagedSeen.get(item);
  }

  /** The editor resources directory once known; the package's own copy (an offline package) until then. */
  get editorDir() { return this.editorPath; }

  /**
   * The Kit executable, downloading Kit and the editor resources first when needed (KitProcess calls this before each
   * start). Refuses a Kit below what a newer Kit on this machine requires (§4.3).
   */
  async binary() {
    if (this.unsupported) throw Object.assign(new Error(`OfficeSDK has no document engine for ${this.unsupported.platform}.`), { code: 'unsupported-platform' });
    // A platform office-kit releases a Kit for but nobody has verified here yet: it runs, and says so once.
    if (this.kitItem.status === 'test' && !this.testNoted) { this.testNoted = true; this.log('warn', `the document engine for ${this.platform} is in test (kit-manifest.json status test)`); }
    const refused = downgradeProblem(this.manifest, await readHighest(this.root));
    if (refused) throw Object.assign(new Error(refused.message), { code: 'kit-downgrade' });
    try {
      const signal = this.lifetime.signal;
      for (const item of [this.kitItem, this.editorItem]) this.packaged(item, { fresh: true });
      const [kit, editor] = await Promise.all([this.downloader.ensure(this.kitItem, { signal }), this.downloader.ensure(this.editorItem, { signal })]);
      this.editorPath = editor.dir;
      this.kitDir = kit.dir;
      return join(kit.dir, this.kitItem.binary);
    } catch (error) {
      if (this.lifetime.signal.aborted) throw this.lifetime.signal.reason;
      // The engine downloaded and checked out but would not start (or reports another version): say that, not "download failed".
      const selftest = [error, ...(error?.failures ?? [])].find(item => item?.code === 'kit-selftest');
      if (selftest) throw Object.assign(new Error(`文档引擎下载完成但无法启动：${selftest.message}。可重试；Windows 上可能是安全软件正在检查。The document engine was downloaded but does not start (${selftest.message}).`), { code: 'kit-selftest' });
      // The reason that matters: a source that was tried (network, HTTP, checksum), not one skipped for lack of a token.
      const failures = error?.failures ?? [];
      const reason = (failures.filter(failure => failure.code !== 'no-token').at(-1) ?? failures.at(-1))?.message ?? error?.message ?? String(error);
      throw Object.assign(new Error(`文档引擎下载失败：${reason}。请检查网络后重试。The document engine could not be downloaded (${reason}); retry when online.`), { code: error?.code === 'denied' ? 'denied' : 'download-failed', cause: error });
    }
  }

  /** Starts the download in the background shortly after start, so the first document seldom waits. */
  prefetch({ delayMs = PREFETCH_DELAY_MS } = {}) {
    if (this.unsupported || this.prefetching) return;
    this.prefetching = setTimeout(() => { void this.binary().catch(error => this.log('warn', error.message)); }, delayMs);
    this.prefetching.unref?.();
  }

  /** Stops waiting and downloading (the plugin unloads); a later instance resumes from the .part files. */
  dispose() { clearTimeout(this.prefetching); this.lifetime.abort(Object.assign(new Error('The plugin was unloaded.'), { code: 'kit-suspended' })); }

  /** After a Kit started and opened its work directory: the downgrade record moves up (never on a failed start). */
  async markStarted() { try { await recordHighest(this.root, this.manifest); } catch (error) { this.log('warn', `highest.json not written: ${error?.message}`); } }

  /**
   * Removes other versions' Kit and editor directories not used for 30 days (§4.3 cleanup). Always kept: this version,
   * the most recently used other version (a rollback needs no download), and a Kit that is running (the work root
   * records which binary it started, `kitWork/.officesdk-kit.running`). Editor directories are not kept when this
   * package carries the editor resources itself.
   */
  async prune({ kitWork, isAlive = pid => { try { process.kill(pid, 0); return true; } catch (error) { return error.code === 'EPERM'; } } } = {}) {
    const dirs = [];
    // Hidden names are downloads in progress (`.x.tmp-…`, `.x.old-…`); they are never pruned here.
    const list = dir => { try { return readdirSync(dir).filter(name => !name.startsWith('.')); } catch { return []; } };
    for (const version of list(join(this.root, 'kit'))) for (const platform of list(join(this.root, 'kit', version))) dirs.push(join(this.root, 'kit', version, platform));
    for (const version of list(join(this.root, 'editor'))) dirs.push(join(this.root, 'editor', version));
    // A packaged editor (261112-bundle-editor) needs no data directory copy: an older plugin's download, this version's
    // included, ages out. While the packaged copy is incomplete the download in use stays.
    const editorPackaged = this.packaged(this.editorItem, { fresh: true });
    const keep = new Set([this.kitItem?.dir, !editorPackaged && this.editorItem.dir].filter(Boolean));
    let running;
    try { running = JSON.parse(readFileSync(join(kitWork, '.officesdk-kit.running'), 'utf8')); } catch {}
    if (running?.binary && running.pid && isAlive(running.pid)) keep.add(dirname(running.binary));
    const used = Object.values(this.downloader.readState().items ?? {}).filter(entry => entry?.dir && !keep.has(entry.dir)).sort((a, b) => (b.lastUsed ?? 0) - (a.lastUsed ?? 0));
    for (const kind of editorPackaged ? ['kit'] : ['kit', 'editor']) { const previous = used.find(entry => entry.dir.startsWith(join(this.root, kind))); if (previous) keep.add(previous.dir); }
    return this.downloader.prune({ dirs, keep: dir => keep.has(dir) });
  }

  /** What the Client shows while a document waits: `{ state, received, total, label?, error? }`. */
  status() {
    if (this.unsupported) return { state: 'unsupported', platform: this.unsupported.platform };
    // Only what has to be downloaded counts: a packaged item is ready and adds nothing to the sizes shown.
    const rows = [this.kitItem, this.editorItem].filter(item => !this.packaged(item)).map(item => ({ item, progress: this.progress[item.key] }));
    const failed = rows.find(row => row.progress?.phase === 'failed');
    const total = rows.reduce((sum, row) => sum + (row.item.archive?.size ?? 0), 0);
    const received = rows.reduce((sum, row) => sum + (this.downloader.isComplete(row.item) || row.progress?.phase === 'ready' ? row.item.archive.size : row.progress?.received ?? 0), 0);
    if (failed) return { state: 'failed', received, total, error: failed.progress.error };
    if (rows.every(row => this.downloader.isComplete(row.item) || row.progress?.phase === 'ready')) return { state: 'ready', received: total, total };
    const busy = rows.find(row => row.progress && row.progress.phase !== 'ready');
    if (!busy) return { state: 'idle', received, total };
    // `source`: where the busy item is coming from; `fallback`: sources it gave up on (a mirror or npm that did not answer).
    const fallback = rows.flatMap(row => (row.progress?.phase !== 'ready' && row.progress?.fallback) || []);
    const testing = this.kitItem?.status === 'test';
    const text = `正在下载文档引擎 ${mb(received)}/${mb(total)} MB${fallback.includes('npm') ? '（npm 源连不上，已改用其他来源下载）' : ''}${testing ? '（此平台测试中）' : ''}`;
    return { state: busy.progress.phase === 'waiting' ? 'waiting' : 'downloading', phase: busy.progress.phase, received, total, text,
      ...busy.progress.source && { source: busy.progress.source }, ...fallback.length && { fallback: [...new Set(fallback)] }, ...testing && { platformStatus: 'test' } };
  }
}
