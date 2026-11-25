import { readFileSync } from 'node:fs';
import { mkdir, readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { baseOf, compareVersions as compare, parseVersion, stageOf } from './version.mjs';

/**
 * The CHANGELOG section a package version belongs to: a release or preview stage keeps its own (0.0.0-alpha.2 is
 * 0.0.0-alpha.2); a test build and the old Windows forms are their release (0.0.47-test.3, 0.0.47-win32 and
 * 0.0.47-win32-test.2 are all 0.0.47).
 */
export function releaseOf(version) {
  const parsed = parseVersion(version);
  if (!parsed) return undefined;
  const base = `${parsed.major}.${parsed.minor}.${parsed.patch}`;
  return ['alpha', 'beta', 'rc'].includes(parsed.stage) ? `${base}-${parsed.stage}.${parsed.number}` : base;
}
/** Ordering of the sections two versions belong to (releaseOf; previews before their release); NaN for a non-version. */
export const compareVersions = (left, right) => compare(releaseOf(left), releaseOf(right));

/** whats-new.json as shipped, or no notes when it is missing or not the expected shape. */
export function readNotes(file) {
  try {
    const parsed = JSON.parse(readFileSync(file, 'utf8'));
    if (parsed?.v !== 1 || !Array.isArray(parsed.entries)) return { entries: [] };
    const points = list => Array.isArray(list) && list.length > 0 && list.every(text => typeof text === 'string' && text.trim());
    // highlightsEn (the English points, one per Chinese point) is optional: a malformed one is dropped, the entry kept.
    const entries = parsed.entries.filter(entry => (entry?.version === 'unreleased' || releaseOf(entry?.version)) && points(entry.highlights))
      .map(({ version, highlights, highlightsEn }) => ({ version, highlights, ...points(highlightsEn) && highlightsEn.length === highlights.length && { highlightsEn } }));
    return { latestReleased: releaseOf(parsed.latestReleased), entries };
  } catch { return { entries: [] }; }
}

/**
 * The notes for an update from `previous` to `current`, newest first (`{ version, highlights, highlightsEn? }`; a version
 * has English points only when every entry that makes it up has them): versions above `previous` up to `current`, or only
 * `current`'s own when there is no previous (a new install, or one from before this feature). 未发布 points belong to
 * `current` when it is newer than every released section (the package was built before the section got its number).
 */
export function notesBetween(notes, previous, current) {
  const target = releaseOf(current);
  if (!target) return [];
  const versionOf = entry => entry.version !== 'unreleased' ? releaseOf(entry.version)
    : !notes.latestReleased || compareVersions(target, notes.latestReleased) > 0 ? target : undefined;
  const picked = new Map();
  for (const entry of notes.entries) {
    const version = versionOf(entry);
    if (!version || compareVersions(version, target) > 0) continue;
    if (previous ? compareVersions(version, previous) <= 0 : version !== target) continue;
    const before = picked.get(version) ?? { highlights: [], highlightsEn: [] };
    picked.set(version, { highlights: [...before.highlights, ...entry.highlights],
      highlightsEn: before.highlightsEn && entry.highlightsEn ? [...before.highlightsEn, ...entry.highlightsEn] : undefined });
  }
  return [...picked].sort(([a], [b]) => compareVersions(b, a))
    .map(([version, { highlights, highlightsEn }]) => ({ version, highlights, ...highlightsEn && { highlightsEn } }));
}

/**
 * 「本次更新」: what changed since the version this machine last showed notes for, kept in `<dataRoot>/whats-new-seen.json`
 * (`{ v: 1, version, at }`). Each version is shown once: the client marks it seen when its card appears.
 */
export class WhatsNew {
  constructor({ packageRoot, file, atomic, now = () => new Date().toISOString() }) {
    Object.assign(this, { file, atomic, now });
    this.notesFile = join(packageRoot, 'whats-new.json');
    try { this.current = JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8')).version; } catch {}
  }
  notes() { return this.cached ??= readNotes(this.notesFile); }
  async seen() {
    try {
      const parsed = JSON.parse(await readFile(this.file, 'utf8'));
      return parsed?.v === 1 && releaseOf(parsed.version) ? parsed.version : undefined;
    } catch { return undefined; }
  }
  /** `{ current, previous?, entries }` to show, or null; nothing to show (same version, a downgrade, no points) records this version. */
  async pending() {
    if (!releaseOf(this.current)) return null;
    const seen = await this.seen();
    const order = seen === undefined ? undefined : compareVersions(seen, this.current);
    if (order === 0) return null;
    const previous = order < 0 ? releaseOf(seen) : undefined;
    // A preview of a lower X.Y.Z than what was seen is the new numbering, not a downgrade: 0.0.0-alpha.1 installed over
    // the git releases (0.0.50) shows its own points, like a new install.
    const restart = order > 0 && ['alpha', 'beta', 'rc'].includes(stageOf(this.current)) && baseOf(seen) !== baseOf(this.current);
    const entries = order > 0 && !restart ? [] : notesBetween(this.notes(), previous, this.current);
    if (!entries.length) { await this.markSeen(this.current).catch(() => {}); return null; }
    return { current: releaseOf(this.current), ...previous && { previous }, entries };
  }
  /** Records `version` (this package's own, any suffix) as shown. */
  async markSeen(version) {
    if (!releaseOf(this.current) || compareVersions(version, this.current) !== 0) throw new Error(`not this plugin's version: ${version}`);
    await mkdir(dirname(this.file), { recursive: true, mode: 0o700 });
    await this.atomic.withFileLock(`${this.file}.lock`, () =>
      this.atomic.writeFileAtomic(this.file, `${JSON.stringify({ v: 1, version: this.current, at: this.now() }, null, 2)}\n`, { mode: 0o600 }));
  }
}
