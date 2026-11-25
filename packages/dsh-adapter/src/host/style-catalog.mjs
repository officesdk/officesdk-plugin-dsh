import { readFileSync, realpathSync } from 'node:fs';
import { mkdir, readFile } from 'node:fs/promises';
import { dirname, join, sep } from 'node:path';

/**
 * What the input box's PPT style picker shows of each style pack, and which pack a workspace uses when the user names
 * none. A pack describes itself in `style.json` `meta`, whose fields the style distill toolkit owns and validates
 * (its spec §1.1: id, name, aliases, summary, scenarios, requires, thumb, previews, fonts, design_source; fields are only
 * ever added). The picker reads id, name, aliases, summary, scenarios, thumb, previews and the photo note; everything is read defensively, so a
 * pack without meta still lists under its style.json title.
 */
export const THUMB_MAX_BYTES = 512 * 1024;
export const PREVIEWS_MAX = 4;
const PNG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/** A file name inside a pack (letters, digits, `.`, `_`, `-`; ends in .png; no leading dot), else undefined. */
const packFile = name => (typeof name === 'string' && /^[\w.-]+\.png$/i.test(name) && !name.startsWith('.') ? name : undefined);
/** A PNG within THUMB_MAX_BYTES as a data URI (the Client shows it without fetching from the Host), else null. */
function pngDataUri(path) {
  try {
    const bytes = readFileSync(path);
    return bytes.length <= THUMB_MAX_BYTES && PNG.every((value, index) => bytes[index] === value) ? `data:image/png;base64,${bytes.toString('base64')}` : null;
  } catch { return null; }
}

/** `{ zh, en }` from a string or a partial `{ zh, en }`; either language falls back to the other, then to `fallback`. */
export function localized(value, fallback = '') {
  if (typeof value === 'string') return { zh: value.trim() || fallback, en: value.trim() || fallback };
  const zh = typeof value?.zh === 'string' ? value.zh.trim() : '';
  const en = typeof value?.en === 'string' ? value.en.trim() : '';
  return { zh: zh || en || fallback, en: en || zh || fallback };
}

/** One pack's card; throws when its style.json cannot be read. */
export function styleCard(root, id, { thumb: withThumb = true } = {}) {
  const style = JSON.parse(readFileSync(join(root, id, 'style.json'), 'utf8'));
  const meta = style?.meta && typeof style.meta === 'object' && !Array.isArray(style.meta) ? style.meta : {};
  const name = localized(meta.name, typeof style?.title === 'string' && style.title.trim() ? style.title.trim() : id);
  const strings = list => (Array.isArray(list) ? [...new Set(list.filter(item => typeof item === 'string' && item.trim()).map(item => item.trim()))] : []);
  const scenarios = { zh: strings(meta.scenarios?.zh), en: strings(meta.scenarios?.en) };
  const photos = meta.requires?.photos;
  const count = value => (Number.isInteger(value) && value >= 0 ? value : null);
  const primary = style?.palettes?.[style?.default_palette]?.primary;
  // `thumb` names a file inside the pack (no folders, no ..); thumb.png when meta does not say.
  const thumbFile = packFile(meta.thumb) ?? 'thumb.png';
  const thumb = withThumb ? pngDataUri(join(root, id, thumbFile)) : null;
  // `previews`: optional page images (2–4 named files) the gallery shows under the thumbnail; unreadable ones are left out.
  const previews = withThumb && Array.isArray(meta.previews)
    ? [...new Set(meta.previews.map(packFile).filter(Boolean))].slice(0, PREVIEWS_MAX).map(file => pngDataUri(join(root, id, file))).filter(Boolean)
    : [];
  return {
    id, name, aliases: strings(meta.aliases), summary: localized(meta.summary),
    scenarios: { zh: scenarios.zh.length ? scenarios.zh : scenarios.en, en: scenarios.en.length ? scenarios.en : scenarios.zh },
    requires: { note: localized(meta.requires?.note), photos: photos && typeof photos === 'object' ? { slots: count(photos.slots), required: count(photos.required), perSlideMax: count(photos.per_slide_max) } : null },
    primary: typeof primary === 'string' ? primary : null, thumb, previews,
  };
}

/** Cards for the named packs, in order; a pack that cannot be read is left out, never blocking the others. */
export function styleCards(root, names, warn = () => {}) {
  const cards = [];
  for (const id of names) {
    try { cards.push(styleCard(root, id)); }
    catch (error) { warn(`style ${id} skipped: ${error?.message ?? error}`); }
  }
  return cards;
}

const canonical = path => { try { return realpathSync(path); } catch { return path; } };
const covers = (root, path) => path === root || path.startsWith(root.endsWith(sep) ? root : root + sep);

/**
 * The style each workspace remembers, in `<dataRoot>/style-defaults.json` (`{ v: 1, workspaces: { <root>: { style,
 * updated } } }`), kept in memory so a synchronous reader (the model's runtime context) can ask. Roots and queries are
 * real paths; a query takes the longest remembered root that is it or contains it.
 */
export class StyleMemory {
  constructor({ file, atomic, now = () => new Date().toISOString(), warn = () => {} }) {
    Object.assign(this, { file, atomic, now, warn });
    this.workspaces = this.read(() => readFileSync(file, 'utf8'));
  }
  read(load) {
    try {
      const parsed = JSON.parse(load());
      const workspaces = parsed?.v === 1 && parsed.workspaces && typeof parsed.workspaces === 'object' ? parsed.workspaces : {};
      return Object.fromEntries(Object.entries(workspaces).filter(([, entry]) => typeof entry?.style === 'string'));
    } catch (error) {
      if (error?.code !== 'ENOENT') this.warn(`style defaults unreadable, starting empty: ${error?.message ?? error}`);
      return {};
    }
  }
  /** The remembered style id for a path (workspace root or a folder inside it), or undefined. */
  rememberedFor(path) { return this.lookup(path)?.style; }
  /** `{ style, root }` of the longest remembered root covering `path`, and whether that root is `path` itself. */
  lookup(path) {
    if (typeof path !== 'string' || !path) return undefined;
    const target = canonical(path);
    let best;
    for (const root of Object.keys(this.workspaces)) if (covers(root, target) && (!best || root.length > best.length)) best = root;
    return best && { style: this.workspaces[best].style, root: best, own: best === target };
  }
  /** Remember `style` for the workspace at `root`, or forget it with null. Re-reads the file under the lock first. */
  async set(root, style) {
    const key = canonical(root);
    await mkdir(dirname(this.file), { recursive: true, mode: 0o700 });
    await this.atomic.withFileLock(`${this.file}.lock`, async () => {
      const workspaces = this.read(() => readFileSync(this.file, 'utf8'));
      if (style === null) delete workspaces[key];
      else workspaces[key] = { style, updated: this.now() };
      await this.atomic.writeFileAtomic(this.file, `${JSON.stringify({ v: 1, workspaces }, null, 2)}\n`, { mode: 0o600 });
      this.workspaces = workspaces;
    });
  }
}

/** Test helper and diagnostics: the stored file as it is on disk. */
export const readStyleDefaults = async file => JSON.parse(await readFile(file, 'utf8'));
