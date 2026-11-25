import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

/**
 * skills-catalog.json (spec 261110-universal-package §3.4): the skill versions there are, each with its description,
 * where its files come from (a source commit and path, a pack) and every file's sha512. The build writes one into the
 * package (the versions it ships, `layer: light|full`); a skill mirror may keep a remote one. One generator for both, so a
 * packaged entry and the remote entry of the same version are equal file for file.
 */
export const CATALOG_FORMAT = 1;
export const SKILLS_PATH = 'packages/dsh-adapter/skills';
// The universal package (light skills downloaded on first use) starts at 0.0.0-alpha.1; compareBase compares X.Y.Z only.
export const DEFAULT_MIN_PLUGIN = '0.0.0';
const sri = bytes => `sha512-${createHash('sha512').update(bytes).digest('base64')}`;

/** `name`, `description`, `whenToUse`, `metadata.version`, `metadata.minPluginVersion` of a SKILL.md. */
export function skillFields(text) {
  const front = /^﻿?---\r?\n([\s\S]*?)\r?\n---/.exec(text)?.[1] ?? '';
  const scalar = key => {
    const match = new RegExp(`^${key}:[ \\t]*(.*)$`, 'm').exec(front)?.[1]?.trim();
    if (!match) return undefined;
    if (/^".*"$/.test(match)) return JSON.parse(match);
    if (/^'.*'$/.test(match)) return match.slice(1, -1).replaceAll("''", "'");
    return match.replace(/\s+#.*$/, '');
  };
  const meta = key => new RegExp(`^metadata:[\\s\\S]*?^[ \\t]+${key}:[ \\t]*['"]?([^'"\\s]+)`, 'm').exec(front)?.[1];
  return { name: scalar('name'), description: scalar('description'), whenToUse: scalar('whenToUse'), version: meta('version'), minPluginVersion: meta('minPluginVersion') };
}

/** Every file under `dir` as `{ [relative path]: Buffer }` (no dot files). */
export function readTree(dir, prefix = '') {
  const out = {};
  for (const entry of readdirSync(join(dir, prefix), { withFileTypes: true })) {
    if (entry.name.startsWith('.')) continue;
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) Object.assign(out, readTree(dir, rel));
    else out[rel] = readFileSync(join(dir, rel));
  }
  return out;
}

export const skillTag = (name, version) => `skill/${name}/v${version}`;

/** One catalog entry from a skill's files (`{ path: Buffer }`, SKILL.md included). */
export function skillEntry(name, files, { commit = null, channel = 'test', pack } = {}) {
  const skillMd = files['SKILL.md'];
  if (!skillMd) throw new Error(`${name}: no SKILL.md`);
  const fields = skillFields(skillMd.toString('utf8'));
  if (fields.name !== name) throw new Error(`${name}: SKILL.md names ${fields.name}`);
  if (!/^\d+\.\d+\.\d+$/.test(fields.version ?? '')) throw new Error(`${name}: SKILL.md metadata.version must be X.Y.Z`);
  const names = Object.keys(files).sort();
  const dirs = prefix => [...new Set(names.filter(file => file.startsWith(prefix)).map(file => file.slice(prefix.length).split('/')[0]).filter(part => !part.includes('.')))];
  const stems = (prefix, ext) => names.filter(file => file.startsWith(prefix) && file.endsWith(ext) && !file.slice(prefix.length).includes('/')).map(file => file.slice(prefix.length, -ext.length));
  return {
    version: fields.version, channel, minPluginVersion: fields.minPluginVersion ?? DEFAULT_MIN_PLUGIN,
    description: fields.description, ...fields.whenToUse && { whenToUse: fields.whenToUse },
    source: { commit, path: `${SKILLS_PATH}/${name}`, tag: skillTag(name, fields.version) },
    ...pack && { pack },
    ...name === 'shimo-pptx' && { stylePacks: dirs('styles/') },
    ...name === 'shimo-docx' && { templates: stems('templates/', '.docx') },
    ...name === 'shimo-xlsx' && { templates: stems('templates/', '.xlsx') },
    files: Object.fromEntries(names.map(file => [file, { sha512: sri(files[file]) }])),
  };
}

export const catalogText = catalog => `${JSON.stringify(catalog, null, 2)}\n`;
export function catalogOf(entries, { layer, remoteCatalog, updated = null }) {
  return { v: CATALOG_FORMAT, ...layer && { layer }, ...remoteCatalog !== undefined && { remoteCatalog }, updated, skills: Object.fromEntries(Object.entries(entries).map(([name, list]) => [name, [].concat(list)])) };
}

/** Problems between two entries of one version (files and descriptions must agree). */
export function entryProblems(name, expected, actual) {
  const problems = [];
  const files = new Set([...Object.keys(expected.files), ...Object.keys(actual?.files ?? {})]);
  for (const file of files) if (expected.files[file]?.sha512 !== actual?.files?.[file]?.sha512) problems.push(`${name} ${expected.version}: ${file} differs`);
  return problems;
}

const base = version => /^(\d+)\.(\d+)\.(\d+)/.exec(version ?? '')?.slice(1).map(Number);
/** Compares X.Y.Z only (a test plugin 0.1.0-test.7 counts as 0.1.0, spec §3.4). */
export function compareBase(a, b) {
  const [x, y] = [base(a), base(b)];
  if (!x || !y) return 0;
  for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] - y[i];
  return 0;
}

/**
 * The versions of `name` this plugin may use: packaged entries always; remote ones when the channel allows (a stable
 * plugin sees stable, a test plugin test and stable), `minPluginVersion` ≤ the plugin's X.Y.Z, and not yanked.
 * Highest first; equal versions keep the packaged one.
 */
export function availableVersions(name, { packaged, remote, pluginVersion }) {
  // A preview or test build (alpha, beta, rc, test) also sees test skills; a release only stable ones.
  const testPlugin = /-(?:alpha|beta|rc|test)\.\d+$/.test(pluginVersion ?? '');
  const own = (packaged?.skills?.[name] ?? []).map(entry => ({ ...entry, packaged: true }));
  const others = (remote?.skills?.[name] ?? []).filter(entry => !entry.yanked && (testPlugin || entry.channel === 'stable') && compareBase(entry.minPluginVersion ?? DEFAULT_MIN_PLUGIN, pluginVersion) <= 0);
  const yanked = new Set((remote?.skills?.[name] ?? []).filter(entry => entry.yanked).map(entry => entry.version));
  const all = [...own.filter(entry => !yanked.has(entry.version)), ...others.filter(entry => !own.some(mine => mine.version === entry.version))];
  return all.sort((a, b) => compareBase(b.version, a.version));
}
