#!/usr/bin/env node
/**
 * The npm release gate (spec 261111-npm-publish-prep §5.2, §13.4, §13.6): nothing is published to npm unless every check
 * passes. Reads a packed tarball, never the source tree.
 *
 *   node scripts/npm-gate.mjs <package.tgz> [--target ci|local|npmjs] [--registry <url>] [--tag next|latest] [--allow <json>] [--kit-manifest <json>]
 *     [--internal-hosts <json>]
 *
 *   files       the tarball holds only what package.json `files` lists (plus what npm always adds) and nothing forbidden
 *               (office-kit's editor resources are not in it: they download on first use, spec 261113-bundle-public-skills §6)
 *   secrets     no tokens, keys, passwords or credentials in URLs, in any text file
 *   internal    no private address and no host of the publisher's own organisation in the package, unless listed with its
 *               reason in scripts/npm-gate-allow.json; the download profile names only public default sources. The
 *               organisation's hosts are configuration, never committed: NPM_GATE_INTERNAL_HOSTS (comma-separated
 *               domains), else --internal-hosts / NPM_GATE_INTERNAL_HOSTS_FILE, else scripts/npm-gate-internal.local.json
 *               (git-ignored), each `{ "hosts": ["corp.example", …] }`. Without any, `npmjs` refuses to pass.
 *   license     package.json license, LICENSE, THIRD-PARTY-NOTICES.md; bundled packages only under allowed licenses
 *   audit       no dependencies (DSH provides the peers), else `npm audit` without high or critical findings
 *   version     X.Y.Z for latest, X.Y.Z-alpha|beta|rc.N for alpha|beta|rc, X.Y.Z-test.N for next; with --registry: not already published with other bytes, and
 *               next only once latest exists (npm makes the first publish latest whatever its tag)
 *   metadata    publishConfig public, engines, no os/cpu/private, no install scripts; repository (public) for npmjs
 *   package     the universal package's own checks (size, no executables, required files)
 *
 * `--target npmjs` makes every finding an error; `ci` and `local` turn the findings that need public infrastructure not
 * built yet (a repository, a non-public profile, the host list) into warnings. The public build's skills ship in the package (`bundled`,
 * spec 261113-bundle-public-skills): the packaged catalog must then carry them complete with no remote catalog. Prints a JSON report; exits 1 on any error.
 * provenance is not checked: npm provenance needs GitHub Actions or gitlab.com runners, not self-hosted ones (§5.3).
 */
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { universalProblems } from './universal-package.mjs';
import { DIST_TAGS, distTagOf, publishableVersion } from '../packages/dsh-adapter/src/host/version.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
export const DEFAULT_ALLOW = path.join(here, 'npm-gate-allow.json');
/** The publisher's own host list (git-ignored; see the header). */
export const DEFAULT_INTERNAL_HOSTS = path.join(here, 'npm-gate-internal.local.json');
/** The repository's manifest: the packaged kit-manifest.json must be this file (never the package's own copy, which could vouch for itself). */
export const DEFAULT_KIT_MANIFEST = path.join(here, '../components/kit-manifest.json');
export const ALLOWED_LICENSES = Object.freeze(['MIT', 'ISC', 'BSD-2-Clause', 'BSD-3-Clause', 'Apache-2.0', '0BSD', 'CC0-1.0', 'Unlicense', 'BlueOak-1.0.0']);
const ALWAYS = /^(package\.json|README(\.[^/]*)?|LICEN[CS]E(\.[^/]*)?|CHANGELOG(\.[^/]*)?)$/i;
const FORBIDDEN = [/(^|\/)\.npmrc$/, /(^|\/)\.env(\.|$)/, /\.(pem|key|p12|pfx)$/i, /\.map$/, /(^|\/)test\//, /(^|\/)src\//, /(^|\/)node_modules\//, /(^|\/)\.git\//, /\.tgz$/];
/** Text, whatever its name: no NUL byte in its first 8 KB (review R1-F7: not an extension list). */
export const isText = bytes => !bytes.subarray(0, 8192).includes(0);

export const SECRET_PATTERNS = Object.freeze([
  ['gitlab-token', /gl(?:pat|dt|rt|oas|ptt|cbt|ft|imt|agent|soat)-[0-9A-Za-z_-]{20,}/],
  ['npm-token', /npm_[0-9A-Za-z]{36}/],
  ['npm-auth', /_authToken\s*=\s*\S+/],
  ['private-key', /-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----/],
  ['aws-key', /\bAKIA[0-9A-Z]{16}\b/],
  ['url-credentials', /\b[a-z][a-z0-9+.-]*:\/\/[^/\s:@'"`]+:[^/\s@'"`]+@[\w.-]+/i],
  ['password', /\b(?:password|passwd|secret)\s*[:=]\s*['"][^'"\s]{6,}['"]/i],
]);

/** Addresses no public package names, whoever publishes it. */
export const PRIVATE_PATTERNS = Object.freeze([
  ['private-ip', /\b(?:10\.\d{1,3}\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.\d{1,3}\.\d{1,3})\b/g],
]);
/** Source ids a public download profile may list by default (download-sources.json). */
export const PUBLIC_SOURCE_IDS = Object.freeze({ kit: ['mirror', 'release'], editor: ['mirror', 'npm'], skills: ['bundled', 'public'] });
/**
 * Public code hosts a kit-manifest source may live on (a GitHub release): everything on them is public, so the
 * manifest-host rule (meant for a storage bucket whose other paths are not) does not apply to them.
 */
export const PUBLIC_CODE_HOSTS = Object.freeze(['github.com']);

const escapeRegExp = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** A host suffix list as it is configured: lower case, no scheme, no leading dot or `*.`; invalid entries dropped. */
export const normaliseHosts = hosts => [...new Set((hosts ?? []).map(host => String(host).trim().toLowerCase().replace(/^[a-z]+:\/\//, '').replace(/^\*?\./, '').replace(/\/.*$/, '')).filter(host => /^[a-z0-9-]+(\.[a-z0-9-]+)+$/.test(host)))];
/**
 * The patterns the gate looks for: private addresses, every host under one of `internalHosts` (the publisher's own
 * domains, from configuration), and on each host of the packaged kit-manifest.json's sources any path outside that
 * source's release directory (for a storage bucket: its public releases are fine, the rest of it is not meant for the
 * public). A source on a public code host (PUBLIC_CODE_HOSTS, e.g. a GitHub release) adds no such pattern.
 */
export function internalPatterns({ internalHosts = [], kitSources = [] } = {}) {
  const patterns = [...PRIVATE_PATTERNS];
  const hosts = normaliseHosts(internalHosts);
  if (hosts.length) patterns.push(['company-host', new RegExp(`\\b(?:[a-z0-9-]+\\.)*(?:${hosts.map(escapeRegExp).join('|')})\\b`, 'gi')]);
  for (const source of kitSources) {
    let url;
    try { url = new URL(source?.base); } catch { continue; }
    if (PUBLIC_CODE_HOSTS.includes(url.host.toLowerCase())) continue;
    // The release directory: the source's base without its last (version) segment.
    const release = url.pathname.replace(/\/?$/, '/').replace(/[^/]+\/$/, '');
    if (release === '/') continue;
    patterns.push(['manifest-host', new RegExp(`${escapeRegExp(url.host)}\\/(?!${escapeRegExp(release.slice(1))})[^\\s'"\`)]*`, 'gi')]);
  }
  return patterns;
}

/**
 * The configured host list (`[]` when none): a file given explicitly (`--internal-hosts`) alone; otherwise
 * NPM_GATE_INTERNAL_HOSTS, else NPM_GATE_INTERNAL_HOSTS_FILE or scripts/npm-gate-internal.local.json (`{ hosts: [...] }`).
 */
export async function readInternalHosts({ env = process.env, file } = {}) {
  const fromFile = async path => { try { return normaliseHosts(JSON.parse(await fs.readFile(path, 'utf8')).hosts); } catch { return []; } };
  if (file) return fromFile(file);
  if (env.NPM_GATE_INTERNAL_HOSTS?.trim()) return normaliseHosts(env.NPM_GATE_INTERNAL_HOSTS.split(/[\s,]+/));
  return fromFile(env.NPM_GATE_INTERNAL_HOSTS_FILE || DEFAULT_INTERNAL_HOSTS);
}

/** Every download address a kit-manifest.json names: its release sources and its editor registries. */
export const kitManifestUrls = kit => [...(kit?.sources ?? []).map(source => source?.base), kit?.editor?.npm?.registry, ...(kit?.editor?.sources ?? []).map(source => source?.registry ?? source?.base)].filter(Boolean);

export const sri512 = bytes => `sha512-${createHash('sha512').update(bytes).digest('base64')}`;

async function walk(root, prefix = '') {
  const out = [];
  for (const entry of await fs.readdir(path.join(root, prefix), { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) out.push(...await walk(root, rel)); else out.push(rel);
  }
  return out;
}

/** Whether `rel` is covered by package.json `files` (entries are files or directories). */
export const listed = (files, rel) => ALWAYS.test(rel) || files.some(entry => { const e = entry.replace(/^\.?\/+|\/+$/g, ''); return rel === e || rel.startsWith(`${e}/`); });

/**
 * The gate's checks on an unpacked package at `root`. Returns `{ errors, warnings, notes, files }`. `kitManifest`: the
 * text of this repository's components/kit-manifest.json (gate() reads it); the packaged one must equal it.
 */
export async function gateChecks(root, { target = 'ci', allow = [], registryInfo, tag, tgzIntegrity, kitManifest, rollback = false, internalHosts = [] } = {}) {
  const strict = target === 'npmjs';
  const errors = []; const warnings = []; const notes = [];
  const soft = (message) => (strict ? errors : warnings).push(message);
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const files = (await walk(root)).sort();
  let packagedKit;
  try { packagedKit = JSON.parse(await fs.readFile(path.join(root, 'kit-manifest.json'), 'utf8')); } catch {}
  const hosts = normaliseHosts(internalHosts);
  const INTERNAL_PATTERNS = internalPatterns({ internalHosts: hosts, kitSources: packagedKit?.sources ?? [] });
  if (!hosts.length) soft('internal: no company host list is configured (NPM_GATE_INTERNAL_HOSTS or scripts/npm-gate-internal.local.json): only private addresses are checked');

  // files
  for (const rel of files) {
    if (FORBIDDEN.some(pattern => pattern.test(rel))) errors.push(`files: ${rel} must not ship`);
    else if (!listed(manifest.files ?? [], rel)) errors.push(`files: ${rel} is not in package.json files`);
  }

  // secrets and company addresses, in every text file
  const hits = [];
  for (const rel of files) {
    const bytes = await fs.readFile(path.join(root, rel));
    if (!isText(bytes)) continue;
    const text = bytes.toString('utf8');
    for (const [name, pattern] of SECRET_PATTERNS) { const match = pattern.exec(text); if (match) errors.push(`secrets: ${rel} has ${name} (${match[0].slice(0, 12)}…)`); }
    for (const [kind, pattern] of INTERNAL_PATTERNS) for (const match of text.matchAll(pattern)) hits.push({ file: rel, kind, host: match[0].toLowerCase() });
  }
  const unique = [...new Map(hits.map(hit => [`${hit.file}\0${hit.host}`, hit])).values()];
  // A build for a registry other than npmjs (pack-universal --registry) names that registry where it says where it
  // is published and installed from: package.json publishConfig and capabilities.json. Only that host, only there, and
  // never for npmjs.
  let ownRegistry;
  try { ownRegistry = manifest.publishConfig?.registry && !strict ? new URL(manifest.publishConfig.registry).hostname.toLowerCase() : undefined; } catch { ownRegistry = undefined; }
  const ownRegistryHit = hit => ownRegistry === hit.host && ['package.json', 'capabilities.json'].includes(hit.file);
  if (unique.some(ownRegistryHit)) notes.push(`internal: the package's own registry ${ownRegistry} is named in package.json / capabilities.json`);
  // A rule's host `*` covers, in that file, only the hosts the packaged kit-manifest.json names as office-kit's own
  // download sources and editor registries (it is office-kit's file, checked byte for byte); never a private address
  // and never one of the publisher's own (configured) hosts: the manifest cannot vouch for those.
  const kitHosts = new Set(kitManifestUrls(packagedKit).map(url => { try { return new URL(url).host.toLowerCase(); } catch { return undefined; } }).filter(Boolean));
  const starCovers = hit => !['private-ip', 'company-host'].includes(hit.kind) && kitHosts.has(hit.host.split('/')[0]);
  const allowed = hit => ownRegistryHit(hit) || allow.some(rule => (rule.file === hit.file || (rule.file.endsWith('/*') && hit.file.startsWith(rule.file.slice(0, -1)))) && ((rule.host === '*' && starCovers(hit)) || rule.host === hit.host || (rule.host.startsWith('*.') && hit.host.endsWith(rule.host.slice(1)))));
  for (const hit of unique) if (!allowed(hit)) errors.push(`internal: ${hit.file} names ${hit.host} (${hit.kind}); remove it or list it with its reason in scripts/npm-gate-allow.json`);
  notes.push(`internal: ${unique.length} company addresses found, ${unique.filter(allowed).length} listed in the allow file`);

  // default sources (rule a)
  let sources;
  try { sources = JSON.parse(await fs.readFile(path.join(root, 'download-sources.json'), 'utf8')); } catch { errors.push('sources: download-sources.json is missing or unreadable'); }
  if (sources) {
    if (sources.profile !== 'public') soft(`sources: the ${sources.profile} profile is not for a public package (pack with --sources public)`);
    const profile = sources.profiles?.[sources.profile] ?? {};
    const idOf = entry => (typeof entry === 'string' ? entry : entry?.id);
    for (const item of ['kit', 'editor']) for (const id of (profile[item] ?? []).map(idOf)) if (!PUBLIC_SOURCE_IDS[item].includes(id)) soft(`sources: the ${sources.profile} profile tries ${id} for the ${item}, not a public default source`);
    for (const id of [...profile.skills?.sources ?? [], ...profile.skills?.catalog ?? []]) if (!PUBLIC_SOURCE_IDS.skills.includes(id)) soft(`sources: skills come from ${id}, not a public default source`);
    const base = profile.skills?.publicBase ?? '';
    const skillIds = [...profile.skills?.sources ?? [], ...profile.skills?.catalog ?? []];
    if (sources.profile === 'public' && skillIds.length && skillIds.every(id => id === 'bundled')) {
      // Skills in the package: complete, listed alone, never refreshed from a remote catalog (universalProblems checks
      // every file against the catalog and the public skill set).
      let catalog;
      try { catalog = JSON.parse(await fs.readFile(path.join(root, 'skills-catalog.json'), 'utf8')); } catch {}
      if (catalog?.layer !== 'full' || catalog?.remoteCatalog !== false) errors.push('sources: skills are bundled but skills-catalog.json is not layer full with remoteCatalog false');
      else notes.push(`sources: skills bundled in the package (${Object.keys(catalog.skills ?? {}).join(', ')})`);
    } else if (sources.profile === 'public' && !base) soft('sources: no public skill mirror (publicBase): skills cannot download for public users');
    for (const [kind, pattern] of INTERNAL_PATTERNS) if (base && new RegExp(pattern.source, 'i').test(base)) errors.push(`sources: publicBase ${base} is a company address (${kind})`);
    // Where the plugin downloads from: the manifest's release sources and its editor registries, none of them the
    // publisher's own or a private address.
    if (packagedKit) {
      for (const source of packagedKit.sources ?? []) for (const [kind, pattern] of INTERNAL_PATTERNS) if (kind !== 'manifest-host' && new RegExp(pattern.source, 'i').test(source.base ?? '')) errors.push(`sources: kit-manifest.json source ${source.id} is a company address (${kind})`);
      const registries = [['editor.npm.registry', packagedKit.editor?.npm?.registry], ...(packagedKit.editor?.sources ?? []).map((source, index) => [`editor.sources[${index}].registry`, source?.registry])];
      for (const [where, registry] of registries) for (const [kind, pattern] of INTERNAL_PATTERNS) if (registry && kind !== 'manifest-host' && new RegExp(pattern.source, 'i').test(registry)) errors.push(`sources: kit-manifest.json ${where} ${registry} is a company address (${kind})`);
    }
    else errors.push('sources: kit-manifest.json is missing or unreadable');
  }

  // license
  if (!manifest.license) errors.push('license: package.json has no license');
  for (const file of ['LICENSE', 'THIRD-PARTY-NOTICES.md']) if (!files.includes(file)) errors.push(`license: ${file} is missing`);
  if (files.includes('THIRD-PARTY-NOTICES.md')) {
    const text = await fs.readFile(path.join(root, 'THIRD-PARTY-NOTICES.md'), 'utf8');
    const section = /## Bundled open-source packages\n([\s\S]*?)(?:\n## |$)/.exec(text)?.[1] ?? '';
    for (const line of section.split('\n').filter(line => line.startsWith('- '))) {
      const license = /— (.+)$/.exec(line)?.[1]?.trim();
      if (!ALLOWED_LICENSES.includes(license)) errors.push(`license: ${line.slice(2)} is not under an allowed license (${ALLOWED_LICENSES.join(', ')})`);
    }
  }

  // metadata
  if (manifest.publishConfig?.access !== 'public') errors.push('metadata: publishConfig.access must be public');
  if (manifest.publishConfig?.registry && !/^https:\/\/registry\.npmjs\.org\/?$/.test(manifest.publishConfig.registry)) soft(`metadata: publishConfig.registry is ${manifest.publishConfig.registry}`);
  for (const key of ['node', 'dsh']) if (!manifest.engines?.[key]) errors.push(`metadata: engines.${key} is missing`);
  for (const key of ['os', 'cpu', 'private']) if (key in manifest) errors.push(`metadata: package.json must not set ${key}`);
  for (const key of ['preinstall', 'install', 'postinstall', 'prepare']) if (manifest.scripts?.[key]) errors.push(`metadata: scripts.${key} would run on install (DSH refuses install scripts)`);
  const repository = typeof manifest.repository === 'string' ? manifest.repository : manifest.repository?.url;
  if (!repository) soft('metadata: repository is missing (set OFFICESDK_PUBLIC_REPOSITORY)');
  else for (const [kind, pattern] of INTERNAL_PATTERNS) if (new RegExp(pattern.source, 'i').test(repository)) soft(`metadata: repository ${repository} is a company address (${kind})`);
  const companyEmail = hosts.length ? new RegExp(`@(?:[a-z0-9-]+\\.)*(?:${hosts.map(escapeRegExp).join('|')})\\b`, 'i') : undefined;
  for (const person of [...manifest.maintainers ?? [], manifest.author].filter(Boolean)) if (companyEmail?.test(JSON.stringify(person))) errors.push('metadata: a company email address in maintainers/author');

  // audit
  const deps = { ...manifest.dependencies, ...manifest.optionalDependencies };
  if (Object.keys(deps).length) notes.push(`audit: ${Object.keys(deps).length} dependencies (run by gate())`);
  else notes.push('audit: no dependencies');

  // version
  const version = manifest.version ?? '';
  // The dist-tag is the version's stage: latest X.Y.Z, alpha|beta|rc X.Y.Z-<stage>.N, next X.Y.Z-test.N.
  if (!publishableVersion(version)) errors.push(`version: ${version} is not X.Y.Z, X.Y.Z-alpha|beta|rc.N or X.Y.Z-test.N`);
  else if (tag && DIST_TAGS.includes(tag) && distTagOf(version) !== tag) errors.push(`version: ${tag} takes ${{ latest: 'X.Y.Z', next: 'X.Y.Z-test.N' }[tag] ?? `X.Y.Z-${tag}.N`}, not ${version}`);
  if (tag && !DIST_TAGS.includes(tag)) errors.push(`version: ${tag} is not a dist-tag (${DIST_TAGS.join(', ')})`);
  let alreadyPublished = false; let firstPublish = false;
  if (registryInfo) {
    const published = registryInfo.packument?.versions?.[version]?.dist?.integrity;
    if (published) {
      if (published === tgzIntegrity) { alreadyPublished = true; notes.push(`version: ${version} is already published with these bytes`); }
      else errors.push(`version: ${version} is already published with other bytes (${published.slice(0, 20)}…); a version is never published twice`);
    }
    if (!registryInfo.packument?.['dist-tags']?.latest) firstPublish = true;
    if (firstPublish && tag === 'next') errors.push('version: the registry has no latest yet; the first publish must be a release (npm would tag a test build latest)');
  }

  // the universal package's own checks
  let layer = 'light';
  try { layer = JSON.parse(await fs.readFile(path.join(root, 'skills-catalog.json'), 'utf8')).layer ?? 'light'; } catch {}
  // A rollback republishes an older release as it was (npm-release.mjs rollback): its own kit-manifest.json stands.
  const own = !rollback && kitManifest !== undefined ? { kitManifest } : {};
  for (const problem of await universalProblems(root, { version, layer, ...own })) errors.push(`package: ${problem}`);

  return { errors: [...new Set(errors)], warnings: [...new Set(warnings)], notes, files, alreadyPublished, firstPublish };
}

/** `npm audit` of the package's dependencies in a scratch directory (only when it has any). */
export async function audit(manifest, { registry } = {}) {
  const deps = { ...manifest.dependencies, ...manifest.optionalDependencies };
  if (!Object.keys(deps).length) return [];
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-gate-audit-'));
  try {
    await fs.writeFile(path.join(dir, 'package.json'), JSON.stringify({ name: 'gate-audit', private: true, dependencies: deps }));
    const args = registry ? ['--registry', registry] : [];
    const lock = spawnSync('npm', ['install', '--package-lock-only', '--ignore-scripts', ...args], { cwd: dir, encoding: 'utf8' });
    if (lock.status !== 0) return [`audit: cannot resolve dependencies: ${lock.stderr.trim().split('\n').at(-1)}`];
    const result = spawnSync('npm', ['audit', '--audit-level=high', '--json', ...args], { cwd: dir, encoding: 'utf8' });
    let report = {};
    try { report = JSON.parse(result.stdout); } catch {}
    const counts = report.metadata?.vulnerabilities ?? {};
    return (counts.high || counts.critical) ? [`audit: ${counts.high ?? 0} high and ${counts.critical ?? 0} critical vulnerabilities`] : [];
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
}

/** The registry's packument of the package (undefined when it has none yet). */
export async function readRegistry(registry, name, fetchImpl = fetch) {
  const response = await fetchImpl(`${registry.replace(/\/?$/, '/')}${name.replace('/', '%2F')}`, { headers: { Accept: 'application/json' }, signal: AbortSignal.timeout(30_000) });
  if (response.status === 404) return { packument: undefined };
  if (!response.ok) throw new Error(`${registry}: HTTP ${response.status}`);
  return { packument: await response.json() };
}

/** The whole gate on a tarball. */
export async function gate(tgz, { target = 'ci', registry, tag, allowFile = DEFAULT_ALLOW, kitManifestFile = DEFAULT_KIT_MANIFEST, internalHostsFile, internalHosts, rollback = false, fetchImpl } = {}) {
  const bytes = await fs.readFile(tgz);
  const tgzIntegrity = sri512(bytes);
  const work = await fs.mkdtemp(path.join(os.tmpdir(), 'npm-gate-'));
  try {
    const listing = spawnSync('tar', ['-tzf', tgz], { encoding: 'utf8' });
    if (listing.status !== 0) throw new Error(`${tgz} is not a tarball`);
    const outside = listing.stdout.split('\n').filter(Boolean).filter(entry => !entry.startsWith('package/'));
    spawnSync('tar', ['-xzf', tgz, '-C', work]);
    const root = path.join(work, 'package');
    const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
    const allow = JSON.parse((await fs.readFile(allowFile, 'utf8').catch(() => '')).trim() || '[]');
    const registryInfo = registry ? await readRegistry(registry, manifest.name, fetchImpl) : undefined;
    const kitManifest = await fs.readFile(kitManifestFile, 'utf8');
    const hosts = internalHosts ?? await readInternalHosts(internalHostsFile ? { file: internalHostsFile } : {});
    const report = await gateChecks(root, { target, allow, registryInfo, tag, tgzIntegrity, kitManifest, rollback, internalHosts: hosts });
    for (const entry of outside) report.errors.push(`files: ${entry} is outside package/`);
    report.errors.push(...await audit(manifest, { registry }));
    return { package: `${manifest.name}@${manifest.version}`, target, integrity: tgzIntegrity, bytes: bytes.length, ...report, ok: report.errors.length === 0 };
  } finally {
    await fs.rm(work, { recursive: true, force: true });
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { values, positionals } = parseArgs({ allowPositionals: true, options: { target: { type: 'string' }, registry: { type: 'string' }, tag: { type: 'string' }, allow: { type: 'string' }, 'kit-manifest': { type: 'string' }, 'internal-hosts': { type: 'string' }, quiet: { type: 'boolean' } } });
  if (!positionals[0]) { console.error('usage: node scripts/npm-gate.mjs <package.tgz> [--target ci|local|npmjs] [--registry <url>] [--tag next|latest]'); process.exit(2); }
  const report = await gate(positionals[0], { target: values.target ?? 'ci', registry: values.registry, tag: values.tag, allowFile: values.allow ?? DEFAULT_ALLOW, kitManifestFile: values['kit-manifest'] ?? DEFAULT_KIT_MANIFEST, internalHostsFile: values['internal-hosts'] });
  const { files, ...summary } = report;
  // The file list is the `npm pack --dry-run` check of the CI log (§2).
  if (!values.quiet) console.error(`files (${files.length}):\n${files.map(file => `  ${file}`).join('\n')}`);
  console.log(JSON.stringify(summary, null, 2));
  process.exit(report.ok ? 0 : 1);
}
