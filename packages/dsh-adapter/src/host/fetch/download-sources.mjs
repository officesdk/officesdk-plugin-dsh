import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/**
 * Where downloads come from, as data (spec 261111-npm-publish-prep §6): `download-sources.json` at the package root
 * names an ordered list of source ids per item in its one profile, `public`: the Kit from office-kit's release (`release`:
 * the sources kit-manifest.json names, office-kit's GitHub release), the editor resources from office-kit's npm package
 * (`npm`: kit-manifest.json `editor.npm`; the release does not carry them), each after `mirror` when OFFICESDK_KIT_MIRROR
 * is set; its skills ship inside the package (`bundled`, spec 261113-bundle-public-skills): nothing to download.
 *
 * Both the Kit and the editor resources download on first use; an offline package carries both. The `public` skill
 * source (a mirror at `publicBase`) still works when a file names it (drills).
 *
 * A missing or unreadable file means the defaults below. Environment variables still override (OFFICESDK_KIT_MIRROR,
 * OFFICESDK_EDITOR_NPM, and OFFICESDK_SKILLS_BASE for `publicBase`).
 */
export const SOURCES_FILE = 'download-sources.json';
export const PROFILES = Object.freeze(['public']);
/**
 * The editor's npm registry is its last source (the release carries no editor): it gets time to answer (15 s), to
 * pause in the body (30 s), and three tries, as any last source does.
 */
export const NPM_DEFAULTS = Object.freeze({ connectTimeoutMs: 15_000, idleMs: 30_000, attempts: 3 });
export const PUBLIC_CONNECT_TIMEOUT_MS = 5000;
const KIT_IDS = ['mirror', 'release'];
const EDITOR_IDS = ['mirror', 'npm'];
/** `bundled`: the skills ship complete in the package; no download source is tried. `public`: a skill mirror (drills). */
const SKILL_IDS = ['public', 'bundled'];

export const DEFAULT_SOURCES = Object.freeze({
  v: 1,
  profile: 'public',
  profiles: {
    public: {
      kit: ['mirror', 'release'],
      editor: ['mirror', 'npm'],
      skills: { sources: ['bundled'], catalog: ['bundled'] },
    },
  },
});

const entryOf = value => (typeof value === 'string' ? { id: value } : value && typeof value === 'object' && typeof value.id === 'string' ? value : undefined);
const positive = value => (Number.isFinite(value) && value > 0 ? value : undefined);

/** What is wrong with a sources document (empty when nothing); CI and the gate refuse a package whose file has problems. */
export function sourcesProblems(doc) {
  const problems = [];
  if (doc?.v !== 1) return ['download-sources.json: v must be 1'];
  if (!PROFILES.includes(doc.profile)) problems.push(`download-sources.json: profile must be one of ${PROFILES.join(', ')}`);
  for (const name of PROFILES) {
    const profile = doc.profiles?.[name];
    if (!profile) { problems.push(`download-sources.json: profiles.${name} is missing`); continue; }
    for (const [item, allowed] of [['kit', KIT_IDS], ['editor', EDITOR_IDS]]) {
      if (!Array.isArray(profile[item])) { problems.push(`profiles.${name}.${item} must be a list`); continue; }
      for (const raw of profile[item]) { const entry = entryOf(raw); if (!entry || !allowed.includes(entry.id)) problems.push(`profiles.${name}.${item}: unknown source ${JSON.stringify(raw)}`); }
    }
    for (const list of ['sources', 'catalog']) for (const id of profile.skills?.[list] ?? []) if (!SKILL_IDS.includes(id)) problems.push(`profiles.${name}.skills.${list}: unknown source ${id}`);
    // Skills either ship in the package or download; a profile never mixes the two.
    const skillIds = [...profile.skills?.sources ?? [], ...profile.skills?.catalog ?? []];
    if (skillIds.includes('bundled') && skillIds.some(id => id !== 'bundled')) problems.push(`profiles.${name}.skills: bundled cannot be mixed with download sources`);
    const base = profile.skills?.publicBase;
    if (base && !/^https:\/\//.test(base) && !/^http:\/\/(127\.0\.0\.1|localhost)[:/]/.test(base)) problems.push(`profiles.${name}.skills.publicBase must be https (or a loopback http address for drills)`);
  }
  return problems;
}

/** The parsed document (DEFAULT_SOURCES when the package has none or it is broken: `log` says why). */
export function readDownloadSources(packageRoot, log = () => {}) {
  let doc;
  try { doc = JSON.parse(readFileSync(join(packageRoot, SOURCES_FILE), 'utf8')); }
  catch (error) { if (error?.code !== 'ENOENT') log('warn', `${SOURCES_FILE} is unreadable (${error?.message}); the built-in sources are used`); return DEFAULT_SOURCES; }
  const problems = sourcesProblems(doc);
  if (problems.length) { log('warn', `${SOURCES_FILE}: ${problems.join('; ')}; the built-in sources are used`); return DEFAULT_SOURCES; }
  return doc;
}

/**
 * The selected profile, normalised: `{ name, kit: [{id,…}], editor: [{id,…}], skills: { sources, catalog, publicBase } }`.
 * `env.OFFICESDK_SKILLS_BASE` replaces `publicBase`.
 */
export function selectProfile(doc = DEFAULT_SOURCES, env = process.env) {
  const name = PROFILES.includes(doc?.profile) ? doc.profile : 'public';
  const profile = doc?.profiles?.[name] ?? DEFAULT_SOURCES.profiles[name];
  const skills = profile.skills ?? {};
  const publicBase = (env.OFFICESDK_SKILLS_BASE || skills.publicBase || '').trim();
  return {
    name,
    kit: (profile.kit ?? []).map(entryOf).filter(Boolean),
    editor: (profile.editor ?? []).map(entryOf).filter(Boolean),
    skills: { sources: [...skills.sources ?? []], catalog: [...skills.catalog ?? []], publicBase: publicBase.replace(/\/+$/, '') },
  };
}

/** The npm source options of an editor list entry (DEFAULT timeouts unless the file sets its own). */
export function npmOptions(entry = {}) {
  return { connectTimeoutMs: positive(entry.connectTimeoutMs) ?? NPM_DEFAULTS.connectTimeoutMs, idleMs: positive(entry.idleMs) ?? NPM_DEFAULTS.idleMs, attempts: positive(entry.attempts) ?? NPM_DEFAULTS.attempts };
}

/** The document a build writes for `profile` (pack-universal --sources); `publicBase` (drills) sets a public mirror. */
export function sourcesFor(profile = 'public', { publicBase } = {}) {
  if (!PROFILES.includes(profile)) throw new Error(`unknown sources profile ${profile} (only public)`);
  const doc = JSON.parse(JSON.stringify(DEFAULT_SOURCES));
  doc.profile = profile;
  // A drill of the old public mirror: the public profile downloads its skills from `publicBase` instead of carrying them.
  if (publicBase) doc.profiles.public.skills = { sources: ['public'], catalog: ['public'], publicBase: publicBase.replace(/\/+$/, '') };
  return doc;
}
