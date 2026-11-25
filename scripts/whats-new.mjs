#!/usr/bin/env node
/**
 * The plugin's 「本次更新」 notes: packages/dsh-adapter/whats-new.json, made from the `### 更新要点` list in each
 * CHANGELOG.md version section and shipped in the package (package.json files), so every platform's package carries it
 * without a build step and the Host reads it offline.
 *
 *   node scripts/whats-new.mjs           regenerate whats-new.json and capabilities.json (scripts/capabilities.mjs)
 *   node scripts/whats-new.mjs --check   exit 1 when either is out of date
 *
 * `{ v: 1, latestReleased, entries: [{ version: 'unreleased' | 'X.Y.Z', highlights, highlightsEn? }] }`, newest first, only
 * versions with highlights. `highlightsEn` is the section's `### Highlights` list: the same points in English, one for each
 * Chinese point, shown when DSH's interface language is English. The committed CHANGELOG must give every section with
 * highlights its English list (missingEnglish); the parser itself only checks that a given list matches. `latestReleased` is the newest (first) dated section, with or without highlights: the Host counts the
 * 未发布 points as the running package's own only when that package is newer than it.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { compareVersions as compareAny } from '../packages/dsh-adapter/src/host/version.mjs';

export const HIGHLIGHTS_HEADING = '### 更新要点';
export const HIGHLIGHTS_EN_HEADING = '### Highlights';
export const maintenanceNotes = {
  highlights: ['本次更新包含重要的安全修复和改进'],
  highlightsEn: ['This update includes important security fixes and improvements.'],
};

/** Only a dated, bilingual entry is eligible for a published package's update preview. */
export function releasePreview(notes, version) {
  const entry = notes?.entries?.find(item => item.version === version)
    ?? (/-test\.\d+$/.test(version) && notes?.entries?.find(item => item.version === version.replace(/-test\.\d+$/, '')));
  const points = list => Array.isArray(list) && list.length && list.every(text => typeof text === 'string' && text.trim());
  if (!points(entry?.highlights) || !points(entry.highlightsEn) || entry.highlights.length !== entry.highlightsEn.length)
    throw new Error(`Missing bilingual release notes for ${version}`);
  const genericZh = entry.highlights.includes(maintenanceNotes.highlights[0]);
  const genericEn = entry.highlightsEn.includes(maintenanceNotes.highlightsEn[0]);
  if ((genericZh || genericEn) && !(genericZh && genericEn && entry.highlights.length === 1))
    throw new Error(`Use the maintenance note as the sole point in both languages for ${version}`);
  return { version, highlights: entry.highlights, highlightsEn: entry.highlightsEn };
}
const UNRELEASED = '未发布';

/** Markdown emphasis and code marks dropped: the card shows plain text. */
const plain = text => text.replace(/\*\*(.+?)\*\*/g, '$1').replace(/`([^`]+)`/g, '$1').replace(/\s+/g, ' ').trim();

/** `{ latestReleased, entries }` from CHANGELOG.md; throws on a highlights list that is empty. */
export function parseWhatsNew(markdown) {
  const lines = markdown.split(/\r?\n/);
  const entries = [];
  let latestReleased;
  let section;
  let inHighlights = false;
  let list;
  const finish = () => {
    if (section?.highlightsEn && !section.highlights) throw new Error(`CHANGELOG ${section.name}: ${HIGHLIGHTS_EN_HEADING} without ${HIGHLIGHTS_HEADING}`);
    if (!section?.highlights) return;
    if (!section.highlights.length) throw new Error(`CHANGELOG ${section.name}: ${HIGHLIGHTS_HEADING} has no points`);
    for (const [heading, points] of [[HIGHLIGHTS_HEADING, section.highlights], [HIGHLIGHTS_EN_HEADING, section.highlightsEn ?? []]]) {
      if (points.some(text => !text)) throw new Error(`CHANGELOG ${section.name}: ${heading} has a blank point`);
    }
    if (section.highlightsEn && section.highlightsEn.length !== section.highlights.length) throw new Error(`CHANGELOG ${section.name}: ${HIGHLIGHTS_EN_HEADING} has ${section.highlightsEn.length} points, ${HIGHLIGHTS_HEADING} has ${section.highlights.length}; give one English point for each`);
    entries.push({ version: section.version, highlights: section.highlights, ...section.highlightsEn && { highlightsEn: section.highlightsEn } });
  };
  for (const line of lines) {
    if (line.startsWith('## ')) {
      finish();
      inHighlights = false;
      const name = line.slice(3).trim();
      const version = name === UNRELEASED ? 'unreleased' : /^(\d+\.\d+\.\d+(?:-(?:alpha|beta|rc)\.\d+)?) — /.exec(name)?.[1];
      section = version ? { name, version } : undefined;
      // The newest dated section: the first one (newest first). Not the highest number: 0.0.0-alpha.1 comes after 0.0.50.
      if (version && version !== 'unreleased' && !latestReleased) latestReleased = version;
      continue;
    }
    if (line.startsWith('#')) {
      const key = !section ? undefined : line.trim() === HIGHLIGHTS_HEADING ? 'highlights' : line.trim() === HIGHLIGHTS_EN_HEADING ? 'highlightsEn' : undefined;
      inHighlights = Boolean(key);
      if (key) list = section[key] ??= [];
      continue;
    }
    if (!inHighlights) continue;
    if (/^- /.test(line)) list.push(plain(line.slice(2)));
    else if (/^\s+\S/.test(line) && list.length) list[list.length - 1] = plain(`${list.at(-1)} ${line}`);
    else if (line.trim()) inHighlights = false;
  }
  finish();
  return { latestReleased, entries };
}

/** Version ordering, previews (alpha < beta < rc) before their release (packages/dsh-adapter/src/host/version.mjs). */
export const compareVersions = compareAny;

/** Versions whose highlights have no English list (the committed CHANGELOG must have none). */
export const missingEnglish = ({ entries }) => entries.filter(entry => !entry.highlightsEn).map(entry => entry.version);

export const renderWhatsNew = ({ latestReleased, entries }) => `${JSON.stringify({ v: 1, latestReleased, entries }, null, 2)}\n`;

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const adapter = path.join(repo, 'packages/dsh-adapter');
  const { buildCapabilities, renderCapabilities } = await import('./capabilities.mjs');
  const { PLATFORMS } = await import('./platform-package.mjs');
  const notes = parseWhatsNew(await fs.readFile(path.join(adapter, 'CHANGELOG.md'), 'utf8'));
  const untranslated = missingEnglish(notes);
  if (untranslated.length) {
    console.error(`packages/dsh-adapter/CHANGELOG.md: add a ${HIGHLIGHTS_EN_HEADING} list (the same points in English) to ${untranslated.join(', ')}`);
    process.exit(1);
  }
  const outputs = {
    'whats-new.json': renderWhatsNew(notes),
    'capabilities.json': renderCapabilities(buildCapabilities(repo, { platforms: Object.keys(PLATFORMS) })),
  };
  for (const [name, expected] of Object.entries(outputs)) {
    const target = path.join(adapter, name);
    if (process.argv.includes('--check')) {
      const current = await fs.readFile(target, 'utf8').catch(() => '');
      if (current !== expected) { console.error(`packages/dsh-adapter/${name} is out of date: run node scripts/whats-new.mjs`); process.exitCode = 1; }
      else console.log(`${name} is up to date`);
    } else {
      await fs.writeFile(target, expected);
      console.log(target);
    }
  }
}
