#!/usr/bin/env node
/**
 * Builds the universal package (spec 261110-universal-package §3.1, §6.1) from this repository alone, on any machine:
 *
 *   npm ci --prefix tooling --ignore-scripts
 *   node tooling/kit/vendor.mjs fetch          office-kit's editor bundle, host fonts and kit-manifest.json (components/office-kit.json)
 *   node tooling/suite/pack-universal.mjs --version X.Y.Z[-alpha|beta|rc.N|-test.N] [--out <dir>] [--sources public]
 *     [--repository <url>] [--obfuscate none|light] [--registry <url>]
 *
 * The package is the npm build (spec 261111-npm-publish-prep §2): download-sources.json names the `public` profile (the
 * only one), the public package.json metadata, and its skills shipped inside it, trimmed to what goes public
 * (scripts/public-skills.mjs, spec 261113-bundle-public-skills): no skill downloads. --repository (or
 * OFFICESDK_PUBLIC_REPOSITORY) is the public repository named in package.json. --obfuscate light (§8, §13.4) runs
 * javascript-obfuscator over lib/ and the Kit page at low strength without encoding strings, so the gate still reads them.
 * --registry <url> makes the build for that registry (publishConfig.registry, capabilities.registry); npmjs otherwise.
 *
 *   lib/*, resources/kit-page        rebuilt from source with tooling's esbuild
 *   resources/editor/bundle, host-fonts   components/vendor (fetched by tooling/kit/vendor.mjs, checked against components/office-kit.json)
 *   kit-manifest.json                components/kit-manifest.json (the same)
 *   skills/<name>/…, skills-catalog.json   the public skills complete (layer full, no remote catalog)
 * No Kit, no office-kit editor resources (they download on first use from office-kit's release), no os/cpu.
 * Prints one JSON line: { version, sources, universal: { file, sha256, bytes } }.
 */
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { bundleAdapter, bundleKitPage, remoteMethods, remoteParameterNames, remoteSignatureProblems } from './bundle-adapter.mjs';
import { paths as officeKitPaths, vendorProblems } from '../kit/vendor.mjs';
import { releasePreview } from '../../scripts/whats-new.mjs';
import { registryAddress, stampCapabilities } from '../../scripts/capabilities.mjs';
import { GITATTRIBUTES } from '../../scripts/platform-package.mjs';
import { PLATFORM_DIRS, renameClient, renamePlugin, thirdPartyNotices, universalManifest, universalProblems } from '../../scripts/universal-package.mjs';
import { sourcesFor } from '../../packages/dsh-adapter/src/host/fetch/download-sources.mjs';
import { obfuscateTree } from './obfuscate.mjs';
import { replaceFile } from './link-tree.mjs';
import { catalogOf, catalogText, readTree, skillEntry } from '../../packages/dsh-adapter/src/skills/catalog.mjs';
import { PUBLIC_SKILLS_LOCALE, prunePublicSkills, publicCapabilities } from '../../scripts/public-skills.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const adapter = path.join(repo, 'packages/dsh-adapter');
const { values } = parseArgs({ options: { version: { type: 'string' }, out: { type: 'string' }, 'skip-build': { type: 'boolean' }, sources: { type: 'string' }, repository: { type: 'string' }, obfuscate: { type: 'string' }, registry: { type: 'string' } } });
const profile = values.sources ?? 'public';
if (profile !== 'public') throw new Error(`--sources ${profile}: only the public profile exists`);
const sourcesDoc = sourcesFor(profile);
const repository = values.repository ?? process.env.OFFICESDK_PUBLIC_REPOSITORY ?? undefined;
const obfuscate = values.obfuscate ?? 'none';
if (!['none', 'light'].includes(obfuscate)) throw new Error(`--obfuscate takes none or light, not ${obfuscate}`);
// How capabilities.json tells people to install this release (scripts/capabilities.mjs installFields): from npm, from
// --registry when given (its package.json publishConfig names it too), else npmjs.
const registry = values.registry ? registryAddress(values.registry) : undefined;
if (!values.version) throw new Error('usage: node tooling/suite/pack-universal.mjs --version X.Y.Z[-test.N] [--out <dir>]');
const out = path.resolve(values.out || path.join(repo, 'artifacts'));
const run = (command, args, options = {}) => {
  const result = spawnSync(command, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 64 << 20, ...options });
  if (result.status !== 0) throw new Error(`${command} ${args.join(' ')} failed:\n${result.stderr || result.stdout}`);
  return result.stdout.trim();
};
const sha256 = async file => createHash('sha256').update(await fs.readFile(file)).digest('hex');
// npm pack keeps what it packs in its cache: a throwaway one inside the work directory, emptied after every pack.
let npmCache;
const pack = async (dir, dest) => {
  try {
    const name = run('npm', ['pack', '--silent', '--pack-destination', dest, ...npmCache ? ['--cache', npmCache] : []], { cwd: dir }).split(/\r?\n/).at(-1);
    return path.join(dest, name);
  } finally { if (npmCache) await fs.rm(npmCache, { recursive: true, force: true }); }
};

const vendored = vendorProblems(repo);
if (vendored.length) throw new Error(`office-kit's files are not the pinned ones (node tooling/kit/vendor.mjs fetch):\n${vendored.join('\n')}`);
const kitManifestText = await fs.readFile(path.join(repo, 'components/kit-manifest.json'), 'utf8');
const kitManifest = JSON.parse(kitManifestText);

if (!values['skip-build']) {
  let esbuild;
  try { ({ build: esbuild } = createRequire(path.join(repo, 'tooling', 'package.json'))('esbuild')); }
  catch { throw new Error('esbuild is missing: run npm ci --prefix tooling --ignore-scripts'); }
  const bundled = [...await bundleAdapter({ esbuild, repo }), ...await bundleKitPage({ esbuild, repo })];
  const notices = [];
  for (const name of [...new Set(bundled)].sort()) {
    let meta = {};
    try { meta = createRequire(path.join(repo, 'tooling', 'package.json'))(`${name}/package.json`); } catch {}
    notices.push({ name, version: meta.version, license: typeof meta.license === 'string' ? meta.license : meta.license?.type });
  }
  await fs.writeFile(path.join(adapter, 'THIRD-PARTY-NOTICES.md'), thirdPartyNotices({ bundled: notices, officeKit: kitManifest.officeKit?.version }));
}
if (!await fs.access(path.join(adapter, 'THIRD-PARTY-NOTICES.md')).then(() => true, () => false)) await fs.writeFile(path.join(adapter, 'THIRD-PARTY-NOTICES.md'), thirdPartyNotices({ officeKit: kitManifest.officeKit?.version }));
// A previous platform build's directories must not decide what ships.
for (const dir of PLATFORM_DIRS) await fs.rm(path.join(adapter, dir), { recursive: true, force: true });
const vendor = officeKitPaths(repo).vendor;
await fs.rm(path.join(adapter, 'resources/editor'), { recursive: true, force: true });
await fs.cp(vendor.editorBundle, path.join(adapter, 'resources/editor/bundle'), { recursive: true });
await fs.rm(path.join(adapter, 'resources/host-fonts'), { recursive: true, force: true });
await fs.cp(vendor.hostFonts, path.join(adapter, 'resources/host-fonts'), { recursive: true });
await fs.writeFile(path.join(adapter, 'kit-manifest.json'), kitManifestText);
await fs.writeFile(path.join(adapter, '.gitattributes'), GITATTRIBUTES);

await fs.mkdir(out, { recursive: true });
// Next to --out, outside it (artifact globs never see a leftover): the same mount as the packages and, in CI, the cache.
const temporary = await fs.mkdtemp(path.join(path.dirname(out), '.pack-universal-'));
npmCache = path.join(temporary, 'npm-cache');
try {
  // The source package.json does not list kit-manifest.json yet; npm pack a copy whose manifest does.
  const source = JSON.parse(await fs.readFile(path.join(adapter, 'package.json'), 'utf8'));
  const raw = path.join(temporary, 'raw');
  await fs.mkdir(raw);
  await fs.writeFile(path.join(adapter, 'package.json'), JSON.stringify(universalManifest(source, values.version, { repository, registry }), null, 2) + '\n');
  let rawTarball;
  try { rawTarball = await pack(adapter, raw); }
  finally { await fs.writeFile(path.join(adapter, 'package.json'), JSON.stringify(source, null, 2) + '\n'); }
  const unpacked = path.join(temporary, 'universal');
  await fs.mkdir(unpacked);
  run('tar', ['-xzf', rawTarball, '-C', unpacked]);
  await fs.rm(raw, { recursive: true, force: true });
  const root = path.join(unpacked, 'package');
  await fs.writeFile(path.join(root, 'plugin.yml'), renamePlugin(await fs.readFile(path.join(root, 'plugin.yml'), 'utf8')));
  await fs.writeFile(path.join(root, 'lib/client.js'), renameClient(await fs.readFile(path.join(root, 'lib/client.js'), 'utf8')));
  stampCapabilities(root, { version: values.version, registry });
  await fs.writeFile(path.join(root, 'download-sources.json'), `${JSON.stringify(sourcesDoc, null, 2)}\n`);
  if (obfuscate === 'light') {
    const obfuscator = createRequire(path.join(repo, 'tooling', 'package.json'))('javascript-obfuscator');
    const hostSource = await fs.readFile(path.join(adapter, 'src/host/index.mjs'), 'utf8');
    const sizes = await obfuscateTree(root, { obfuscator, reserved: remoteParameterNames(hostSource), methods: remoteMethods(hostSource) });
    process.stderr.write(`obfuscated: ${JSON.stringify(sizes)}\n`);
    const problems = remoteSignatureProblems(hostSource, await fs.readFile(path.join(root, 'lib/host/index.mjs'), 'utf8'));
    if (problems.length) throw new Error(`obfuscation broke the Remote signatures:\n${problems.join('\n')}`);
  }
  // Skills (spec 261113-bundle-public-skills): the public skills complete in place, listed alone in the catalog and
  // capabilities, no remote catalog, no download source.
  const commit = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: repo, encoding: 'utf8' }).stdout.trim() || null;
  const kept = await prunePublicSkills(path.join(root, 'skills'));
  const entries = Object.fromEntries(Object.keys(kept).map(name => [name, { ...skillEntry(name, readTree(path.join(root, 'skills', name)), { commit }), bundled: 'public' }]));
  const layer = 'full';
  await replaceFile(path.join(root, 'skills-catalog.json'), catalogText(catalogOf(entries, { layer, remoteCatalog: false })));
  const capabilitiesFile = path.join(root, 'capabilities.json');
  await replaceFile(capabilitiesFile, `${JSON.stringify(publicCapabilities(JSON.parse(await fs.readFile(capabilitiesFile, 'utf8')), entries), null, 2)}\n`);
  for (const [lang, text] of Object.entries(PUBLIC_SKILLS_LOCALE)) await replaceFile(path.join(root, 'locale/skills', `${lang}.json`), JSON.stringify(text));
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  manifest.officeSdkReleaseNotes = releasePreview(JSON.parse(await fs.readFile(path.join(root, 'whats-new.json'), 'utf8')), manifest.version);
  manifest.files = [...new Set([...manifest.files, 'skills-catalog.json'])];
  await replaceFile(path.join(root, 'package.json'), JSON.stringify(manifest, null, 2) + '\n');
  const problems = await universalProblems(root, { kitManifest: kitManifestText, version: values.version, layer });
  if (problems.length) throw new Error(`not a universal package:\n${problems.join('\n')}`);
  const universalFile = await pack(root, out);
  await fs.rm(unpacked, { recursive: true, force: true });
  const result = { version: values.version, sources: profile, ...registry && { registry }, obfuscate, universal: { file: path.basename(universalFile), sha256: await sha256(universalFile), bytes: (await fs.stat(universalFile)).size } };
  console.log(JSON.stringify(result));
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}
