import { randomBytes } from 'node:crypto';
import { cp, mkdir, readdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fail } from './errors.mjs';
import { measure } from './paths.mjs';

// Drafts are listed per principal and Session; the source path is kept for display and save-as.
const ownerKey = owner => JSON.stringify([owner.principal, owner.sessionId]);

/**
 * Drafts belong to (principal, Session, source path), not to a connection, so a reload or
 * reconnect can still recover them. Office drafts are exported files; snapshot drafts keep
 * the raw editor package and convert only when someone asks for them.
 */
export class DraftStore {
  constructor({ root, limits, now }) {
    this.root = root;
    this.limits = limits;
    this.now = now;
    this.drafts = new Map();
    this.loading = undefined;
  }
  async load() {
    this.loading ??= (async () => {
      await mkdir(this.root, { recursive: true, mode: 0o700 });
      for (const id of await readdir(this.root)) {
        const directory = join(this.root, id);
        try {
          const meta = JSON.parse(await readFile(join(directory, 'meta.json'), 'utf8'));
          if (meta.id !== id) throw new Error('mismatch');
          this.drafts.set(id, { ...meta, directory });
        } catch { await rm(directory, { recursive: true, force: true }); }
      }
      await this.evict();
    })();
    return this.loading;
  }
  get bytes() { let total = 0; for (const draft of this.drafts.values()) total += draft.bytes; return total; }
  /** source is a file (office) or a directory (snapshot); it is copied, never moved, so callers keep ownership. */
  async add({ owner, format, kind, source, name, sourceVersion, content, assets }) {
    await this.load();
    const id = randomBytes(16).toString('hex');
    const staging = join(this.root, `.${id}`);
    const directory = join(this.root, id);
    await mkdir(staging, { recursive: true, mode: 0o700 });
    try {
      await cp(source, join(staging, 'data'), { recursive: true, verbatimSymlinks: true });
      const copied = join(staging, 'data');
      const bytes = (await stat(copied)).isDirectory()
        ? await measure(copied, { maxBytes: this.limits.maxDraftBytes, maxEntries: this.limits.maxOutputFiles })
        : (await stat(copied)).size;
      const meta = { id, owner, format, kind, name, sourceVersion, content, assets, bytes, createdAt: this.now() };
      await writeFile(join(staging, 'meta.json'), JSON.stringify(meta), { mode: 0o600 });
      await rename(staging, directory);
      this.drafts.set(id, { ...meta, directory });
      await this.evict(id);
      return this.public(this.drafts.get(id));
    } catch (error) { await rm(staging, { recursive: true, force: true }); throw error; }
  }
  /** Oldest first, until count, bytes and age are all within budget; the newest draft is never evicted. */
  async evict(keep) {
    const expired = [...this.drafts.values()].filter(d => d.createdAt + this.limits.draftTtlMs <= this.now());
    for (const draft of expired) await this.remove(draft.id);
    const ordered = [...this.drafts.values()].sort((a, b) => a.createdAt - b.createdAt);
    while (ordered.length && (this.drafts.size > this.limits.maxDrafts || this.bytes > this.limits.maxDraftBytes)) {
      const draft = ordered.shift();
      if (draft.id === keep) continue;
      await this.remove(draft.id);
    }
  }
  async remove(id) {
    const draft = this.drafts.get(id);
    if (!draft) return;
    this.drafts.delete(id);
    await rm(draft.directory, { recursive: true, force: true });
  }
  async list(owner) {
    await this.load();
    await this.evict();
    const key = ownerKey(owner);
    return [...this.drafts.values()].filter(d => ownerKey(d.owner) === key).sort((a, b) => b.createdAt - a.createdAt).map(d => this.public(d));
  }
  async get(owner, id) {
    await this.load();
    const draft = this.drafts.get(id);
    if (!draft || ownerKey(draft.owner) !== ownerKey(owner) || draft.createdAt + this.limits.draftTtlMs <= this.now()) fail('not-found', 'Draft not found.');
    return draft;
  }
  async discard(owner, id) { await this.get(owner, id); await this.remove(id); }
  public(draft) {
    return { id: draft.id, format: draft.format, kind: draft.kind, name: draft.name, path: draft.owner.path, sourceVersion: draft.sourceVersion, bytes: draft.bytes, createdAt: draft.createdAt, expiresAt: draft.createdAt + this.limits.draftTtlMs };
  }
}
