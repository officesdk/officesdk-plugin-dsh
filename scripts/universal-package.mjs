/**
 * The universal package (spec 261110-universal-package §3.1): one @officesdk/officesdk-plugin-dsh for every platform, pure
 * JS, carrying kit-manifest.json instead of office-kit's Kit and editor resources (both download on first use; spec
 * 261113-bundle-public-skills §6 undid 261112-bundle-editor).
 * No dependencies: CI runs this on Linux before anything is installed.
 */
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { publishableVersion } from '../packages/dsh-adapter/src/host/version.mjs';
import { PUBLIC_SKILL_NAMES } from './public-skills.mjs';

export const PACKAGE_NAME = '@officesdk/officesdk-plugin-dsh';
/** office-kit's editor resources: downloaded on first use, never in the package. */
export const EDITOR_DIR = 'resources/office-kit';
/** What the universal package never ships (a Kit, the editor resources, a converter runtime). */
export const PLATFORM_DIRS = Object.freeze(['resources/kit', EDITOR_DIR, 'resources/runtime']);
/** Unpacked bytes: about 4 MB with its skills complete; the editor resources (about 122 MB) never count. */
export const MAX_UNIVERSAL_BYTES = 20 * 1024 * 1024;
/** Files the plugin cannot start without (and skills/<name>/SKILL.md of each skill its download profile requires). */
export const REQUIRED = Object.freeze(['package.json', '.gitattributes', 'plugin.yml', 'kit-manifest.json', 'skills-catalog.json', 'capabilities.json', 'whats-new.json', 'CHANGELOG.md', 'lib/host/index.mjs', 'lib/tools/index.mjs',
  'lib/skills/index.mjs', 'lib/client.js', 'resources/kit-page/index.html', 'resources/kit-page/kit-page.js', 'resources/editor/bundle/offline-editor.js', 'resources/host-fonts/sheet-host-fonts.css',
  'locale/skills/zh.json', 'download-sources.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md']);
export const NPMJS = 'https://registry.npmjs.org/';
/** Files every build adds to package.json `files` (spec 261111-npm-publish-prep §2). */
export const EXTRA_FILES = Object.freeze(['kit-manifest.json', '.gitattributes', 'download-sources.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md']);
/** The web address of a repository URL (`git+https://host/a/b.git` → `https://host/a/b`). */
export const webAddress = url => url.replace(/^git\+/, '').replace(/\.git$/, '');

/** X.Y.Z (a release), X.Y.Z-alpha|beta|rc.N (a preview stage) or X.Y.Z-test.N (a test build). */
export const releaseVersion = publishableVersion;

/**
 * package.json of the universal package, the npm build (spec 261111-npm-publish-prep §2): the published name and version,
 * no os/cpu, no platform directories, published to npmjs (or `registry`) with public access, no maintainers (npm records
 * the publishing account), and `repository`/`homepage`/`bugs` when a public repository is given.
 */
export function universalManifest(manifest, version, { repository, registry } = {}) {
  if (!releaseVersion(version)) throw new Error(`invalid release version ${version} (X.Y.Z, X.Y.Z-alpha|beta|rc.N or X.Y.Z-test.N)`);
  const next = { ...manifest, name: PACKAGE_NAME, version };
  delete next.private; delete next.os; delete next.cpu; delete next.officesdkSource; delete next.maintainers; delete next.scripts;
  // .gitattributes `* -text`: keeps line endings of lib/*.mjs and the manifests' hashed files as packed.
  next.files = [...new Set([...(manifest.files ?? []).filter(file => !PLATFORM_DIRS.includes(file) && !EXTRA_FILES.includes(file)), ...EXTRA_FILES])];
  next.publishConfig = { registry: registry ?? NPMJS, access: 'public' };
  if (repository) {
    const url = repository.startsWith('git+') ? repository : `git+${repository.replace(/\/?$/, '').replace(/(\.git)?$/, '.git')}`;
    next.repository = { type: 'git', url };
    next.homepage = `${webAddress(url)}#readme`;
    next.bugs = { url: `${webAddress(url)}/issues` };
  }
  return next;
}

/** THIRD-PARTY-NOTICES.md: what the package ships that is not this repository's own code. */
export function thirdPartyNotices({ bundled = [], officeKit } = {}) {
  const lines = ['# Third-party notices', '', 'Components this package ships that are not its own source code. The office-kit files listed below are not covered by this package\'s LICENSE (MIT).', '', '## Bundled open-source packages', ''];
  if (bundled.length) for (const item of bundled) lines.push(`- ${item.name}${item.version ? ` ${item.version}` : ''} — ${item.license ?? 'UNKNOWN'}`);
  else lines.push('None: lib/ and resources/kit-page/ bundle only this package\'s own code; DeepSeek Harness packages are peer dependencies it provides.');
  lines.push('', '## office-kit', '', `- resources/editor/bundle/offline-editor.js and resources/host-fonts/: the office-kit editor page${officeKit ? ` (office-kit ${officeKit})` : ''} as office-kit publishes it, and the Sheet icon fonts; covered by office-kit's terms.`,
    '- The document engine and editor resources download on first use from office-kit\'s release and are not part of this package.', '');
  return lines.join('\n');
}

/** DSH imports plugin.yml entries and the client module by the package name. */
export function renamePlugin(text) { return text.replace(/^([ \t]*name:[ \t]*)"?officesdk(\/[\w-]+)?"?[ \t]*$/gm, `$1"${PACKAGE_NAME}$2"`); }
export function renameClient(text) {
  if (!text.startsWith("window.__ModuleLoader__.load({id:'officesdk',")) throw new Error('unexpected client module header');
  return text.replace("{id:'officesdk',", `{id:${JSON.stringify(PACKAGE_NAME)},`);
}

const MAGIC = [[0x7f, 0x45, 0x4c, 0x46], [0x4d, 0x5a], [0xcf, 0xfa, 0xed, 0xfe], [0xce, 0xfa, 0xed, 0xfe], [0xca, 0xfe, 0xba, 0xbe], [0xfe, 0xed, 0xfa, 0xcf]];
/** Whether `bytes` starts like an ELF, PE (MZ) or Mach-O executable. */
export const isExecutableImage = bytes => MAGIC.some(magic => magic.every((value, index) => bytes[index] === value));

async function walk(root, prefix = '') {
  const out = [];
  for (const entry of await fs.readdir(path.join(root, prefix), { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) out.push(...await walk(root, rel));
    else out.push(rel);
  }
  return out;
}
const sri = (algorithm, bytes) => `${algorithm}-${createHash(algorithm).update(bytes).digest('base64')}`;

/** `dir` under `root` holds exactly `files` (kit-manifest.json's `{ [path]: { sha256, executable? } }`): nothing missing, extra or changed. */
export async function treeProblems(root, dir, files) {
  const problems = [];
  const present = new Set(await walk(path.join(root, dir)).catch(() => []));
  for (const [name, entry] of Object.entries(files)) {
    if (!present.has(name)) { problems.push(`${dir}/${name} is missing`); continue; }
    const bytes = await fs.readFile(path.join(root, dir, name));
    if (sri('sha256', bytes) !== entry.sha256) problems.push(`${dir}/${name} differs from kit-manifest.json`);
    if (entry.executable && !((await fs.stat(path.join(root, dir, name))).mode & 0o100)) problems.push(`${dir}/${name} is not executable`);
    present.delete(name);
  }
  for (const extra of present) problems.push(`${dir}/${extra} is not in kit-manifest.json`);
  return problems;
}

/**
 * What is wrong with the unpacked universal package at `root` (empty when nothing). `kitManifest`: the committed
 * components/kit-manifest.json it must carry byte for byte (else the packaged one is read). The editor resources are
 * checked file by file against that manifest's `editor.files`, on the final package: npm pack honours `files` of the
 * nested package.json files office-kit ships (ppt/, writer/), which must not drop anything.
 */
export async function universalProblems(root, { kitManifest, version, layer = 'light' } = {}) {
  const problems = [];
  // Skills (§3.4): the light layer is SKILL.md alone; the public build (`full`) carries its skills complete. Either way
  // the packaged catalog must describe exactly what ships.
  const catalog = JSON.parse(await fs.readFile(path.join(root, 'skills-catalog.json'), 'utf8').catch(() => 'null'));
  if (catalog?.layer !== layer) problems.push(`skills-catalog.json layer is ${catalog?.layer}, expected ${layer}`);
  for (const [name, entries] of Object.entries(catalog?.skills ?? {})) {
    const entry = entries[0];
    const present = await walk(path.join(root, 'skills', name)).catch(() => []);
    const expected = layer === 'full' ? Object.keys(entry.files) : ['SKILL.md'];
    if (JSON.stringify([...present].sort()) !== JSON.stringify([...expected].sort())) problems.push(`skills/${name} ships ${present.length} files, the ${layer} layer has ${expected.length}`);
    for (const file of present) {
      const bytes = await fs.readFile(path.join(root, 'skills', name, file));
      if (entry.files[file] && `sha512-${createHash('sha512').update(bytes).digest('base64')}` !== entry.files[file].sha512) problems.push(`skills/${name}/${file} differs from skills-catalog.json`);
    }
  }
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  if (manifest.name !== PACKAGE_NAME) problems.push(`name is ${manifest.name}`);
  if (version && manifest.version !== version) problems.push(`version is ${manifest.version}, expected ${version}`);
  if (manifest.os || manifest.cpu) problems.push('package.json must not set os or cpu: the package serves every platform');
  for (const file of REQUIRED) if (!await fs.access(path.join(root, file)).then(() => true, () => false)) problems.push(`${file} is missing`);
  // Skills: the public build's skills are `bundled`, exactly the public ones (spec 261113-bundle-public-skills); the
  // catalog lists exactly what skills/ holds.
  let sources;
  try { sources = JSON.parse(await fs.readFile(path.join(root, 'download-sources.json'), 'utf8')); } catch {}
  const skillIds = [...sources?.profiles?.[sources?.profile]?.skills?.sources ?? [], ...sources?.profiles?.[sources?.profile]?.skills?.catalog ?? []];
  const bundledPublic = sources?.profile === 'public' && skillIds.length > 0 && skillIds.every(id => id === 'bundled');
  if (!bundledPublic) problems.push('download-sources.json must name the public profile with bundled skills');
  const profile = 'public';
  const wanted = PUBLIC_SKILL_NAMES;
  for (const name of wanted) if (!await fs.access(path.join(root, 'skills', name, 'SKILL.md')).then(() => true, () => false)) problems.push(`skills/${name}/SKILL.md is missing`);
  const skillDirs = (await fs.readdir(path.join(root, 'skills'), { withFileTypes: true }).catch(() => [])).filter(entry => entry.isDirectory()).map(entry => entry.name);
  for (const name of skillDirs) if (!wanted.includes(name)) problems.push(`skills/${name} must not ship in the ${profile} build`);
  if (JSON.stringify(Object.keys(catalog?.skills ?? {}).sort()) !== JSON.stringify([...wanted].sort())) problems.push(`skills-catalog.json lists ${Object.keys(catalog?.skills ?? {}).sort().join(', ') || 'no skill'}, the ${profile} build carries ${[...wanted].sort().join(', ')}`);
  if (catalog?.layer !== 'full' || catalog?.remoteCatalog !== false) problems.push('the public build carries its skills: skills-catalog.json must be layer full with remoteCatalog false');
  for (const dir of PLATFORM_DIRS) if (await fs.access(path.join(root, dir)).then(() => true, () => false)) problems.push(`${dir} must not ship in the universal package`);
  const shipped = await fs.readFile(path.join(root, 'kit-manifest.json'), 'utf8').catch(() => undefined);
  if (kitManifest !== undefined && shipped !== kitManifest) problems.push('kit-manifest.json differs from components/kit-manifest.json');
  let total = 0;
  for (const rel of await walk(root)) {
    const file = path.join(root, rel);
    const stat = await fs.stat(file);
    total += stat.size;
    const head = Buffer.alloc(4);
    const handle = await fs.open(file, 'r');
    try { await handle.read(head, 0, 4, 0); } finally { await handle.close(); }
    if (isExecutableImage(head)) problems.push(`${rel} is an executable image`);
  }
  if (total > MAX_UNIVERSAL_BYTES) problems.push(`the package is ${(total / 1048576).toFixed(1)} MB, over ${MAX_UNIVERSAL_BYTES / 1048576} MB`);
  return problems;
}
