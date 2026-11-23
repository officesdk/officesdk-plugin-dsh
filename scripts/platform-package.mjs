/**
 * One complete officesdk package per platform (spec 20.05 §4.6, B′): the platform table, the transform
 * that turns an unpacked package into a platform's package, and the structural check of the result.
 * No dependencies: CI runs this on Linux before anything is installed.
 */
import fs from 'node:fs/promises';
import path from 'node:path';

/**
 * `triple`: office-kit's Rust target (its release names the Kit archive by it). `gitattributes`: Git for
 * Windows defaults to core.autocrlf=true and would rewrite modoc.js, bos-pin.txt and lib/*.mjs in a git checkout of
 * the package, so the Windows package carries `* -text`.
 */
export const PLATFORMS = Object.freeze({
  'darwin-arm64': Object.freeze({ os: 'darwin', cpu: 'arm64', triple: 'aarch64-apple-darwin', exe: '', gitattributes: false }),
  'win32-x64': Object.freeze({ os: 'win32', cpu: 'x64', triple: 'x86_64-pc-windows-msvc', exe: '.exe', gitattributes: true }),
});
export const DEFAULT_PLATFORM = 'darwin-arm64';
/**
 * Every Rust target office-kit releases a Kit archive for, by the platform key the plugin uses at run time
 * (`${process.platform}-${process.arch}`). kit-manifest.json lists each one that is not `skip`; PLATFORMS above is the
 * subset the lock pins file by file and the offline packages are built for. A release with a triple missing here fails
 * the manifest build, so a new office-kit platform is a decision, not a silent addition.
 */
export const KIT_TRIPLES = Object.freeze({
  'aarch64-apple-darwin': Object.freeze({ platform: 'darwin-arm64', exe: '' }),
  'x86_64-apple-darwin': Object.freeze({ platform: 'darwin-x64', exe: '' }),
  'x86_64-pc-windows-msvc': Object.freeze({ platform: 'win32-x64', exe: '.exe' }),
  'aarch64-pc-windows-msvc': Object.freeze({ platform: 'win32-arm64', exe: '.exe' }),
  'x86_64-unknown-linux-gnu': Object.freeze({ platform: 'linux-x64', exe: '' }),
  'aarch64-unknown-linux-gnu': Object.freeze({ platform: 'linux-arm64', exe: '' }),
  // OpenHarmony gets the embedded library (liboffline_kit.so + index.d.ts), not the offline-kit-runtime executable.
  'aarch64-unknown-linux-ohos': Object.freeze({ platform: 'openharmony-arm64', exe: '', skip: 'embedded library, no offline-kit-runtime executable' }),
});
/** Kit's executables (with the platform suffix) and the platform-independent files that must sit beside them. */
export const KIT_EXECUTABLES = Object.freeze(['offline-kit-runtime', 'compose-worker']);
export const KIT_SHARED = Object.freeze(['bos-pin.txt', 'modoc.js', 'mow_wasm.d.ts', 'mow_wasm.js', 'mow_wasm_bg.wasm']);
export const GITATTRIBUTES = '* -text\n';

export function platformOf(name = DEFAULT_PLATFORM) {
  const platform = PLATFORMS[name];
  if (!platform) throw new Error(`unknown platform ${name} (known: ${Object.keys(PLATFORMS).join(', ')})`);
  return { name, ...platform };
}

export const kitExecutables = name => KIT_EXECUTABLES.map(file => file + platformOf(name).exe);

/**
 * Versions a package for `name` may carry: X.Y.Z (a release) or X.Y.Z-test.N (a local test build); a Windows package
 * also X.Y.Z-win32-test.N, the Windows test builds made before Windows is validated.
 */
export const versionAllowed = (version, name) => /^\d+\.\d+\.\d+(-test\.\d+)?$/.test(version) || (platformOf(name).os === 'win32' && /^\d+\.\d+\.\d+-win32-test\.\d+$/.test(version));
/** The default version of a local test build for `name`. */
export const testVersion = (base, build, name) => `${base}${platformOf(name).os === 'win32' ? '-win32' : ''}-test.${build}`;

/** The manifest for `name`: only that os/cpu, and `.gitattributes` in `files` exactly when the platform ships one (npm pack drops it otherwise). */
export function platformManifest(manifest, name) {
  const platform = platformOf(name);
  const next = { ...manifest, os: [platform.os], cpu: [platform.cpu] };
  if (Array.isArray(manifest.files)) {
    const files = manifest.files.filter(file => file !== '.gitattributes');
    next.files = platform.gitattributes ? [...files, '.gitattributes'] : files;
  }
  return next;
}

const exists = file => fs.access(file).then(() => true, () => false);

/**
 * Makes the unpacked package at `root` the package for `name`: `kit` (a staged Kit directory) replaces
 * resources/kit; .gitattributes and the
 * manifest follow the platform. Key order of package.json is kept so darwin-arm64 output stays byte-identical.
 */
export async function applyPlatform(root, name, { kit } = {}) {
  const platform = platformOf(name);
  if (kit) {
    await fs.rm(path.join(root, 'resources', 'kit'), { recursive: true, force: true });
    await fs.cp(kit, path.join(root, 'resources', 'kit'), { recursive: true });
  }
  const attributes = path.join(root, '.gitattributes');
  if (platform.gitattributes) await fs.writeFile(attributes, GITATTRIBUTES);
  else await fs.rm(attributes, { force: true });
  const manifestPath = path.join(root, 'package.json');
  const manifest = JSON.parse(await fs.readFile(manifestPath, 'utf8'));
  await fs.writeFile(manifestPath, JSON.stringify(platformManifest(manifest, name), null, 2) + '\n');
}

/** What is wrong with the package at `root` as the package for `name` (empty when it is one). */
export async function checkPackage(root, name) {
  const platform = platformOf(name);
  const problems = [];
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  if (JSON.stringify(manifest.os) !== JSON.stringify([platform.os])) problems.push(`package.json os is ${JSON.stringify(manifest.os)}, expected ["${platform.os}"]`);
  if (JSON.stringify(manifest.cpu) !== JSON.stringify([platform.cpu])) problems.push(`package.json cpu is ${JSON.stringify(manifest.cpu)}, expected ["${platform.cpu}"]`);

  const kit = path.join(root, 'resources', 'kit');
  const entries = await fs.readdir(kit, { withFileTypes: true }).catch(() => null);
  if (!entries) problems.push('resources/kit is missing');
  else {
    const names = entries.map(entry => entry.name);
    for (const file of kitExecutables(name)) {
      if (!names.includes(file)) { problems.push(`resources/kit/${file} is missing`); continue; }
      // npm packs any file with an executable bit as 0755; DSH spawns Kit directly.
      if (!((await fs.stat(path.join(kit, file))).mode & 0o100)) problems.push(`resources/kit/${file} is not executable`);
    }
    for (const file of KIT_SHARED) if (!names.includes(file)) problems.push(`resources/kit/${file} is missing`);
    // Another platform's executables must not ride along: each costs tens of MB and the host would not run them.
    const foreign = new Set(Object.keys(PLATFORMS).filter(other => other !== name).flatMap(kitExecutables));
    for (const file of names) if (foreign.has(file) && !kitExecutables(name).includes(file)) problems.push(`resources/kit/${file} belongs to another platform`);
    // Every shared file has an extension, so on Windows an extension-less file is a Unix executable.
    if (platform.exe) for (const file of names) if (!path.extname(file) && !foreign.has(file)) problems.push(`resources/kit/${file} has no extension`);
  }

  const attributes = path.join(root, '.gitattributes');
  if (platform.gitattributes) {
    const text = await fs.readFile(attributes, 'utf8').catch(() => null);
    if (text === null) problems.push('.gitattributes is missing (Git for Windows would convert line endings)');
    else if (text !== GITATTRIBUTES) problems.push(`.gitattributes is ${JSON.stringify(text)}, expected ${JSON.stringify(GITATTRIBUTES)}`);
    if (Array.isArray(manifest.files) && !manifest.files.includes('.gitattributes')) problems.push('package.json files does not list .gitattributes (npm pack would drop it)');
  } else if (await exists(attributes)) problems.push('.gitattributes must not ship on this platform');
  return problems;
}

/** Kit executables in a tarball listing (`tar -tvzf`) whose mode is not -rwxr-xr-x. */
export function tarballModeProblems(listing, name) {
  const problems = [];
  for (const file of kitExecutables(name)) {
    const line = listing.split('\n').find(row => row.trimEnd().endsWith(` package/resources/kit/${file}`));
    if (!line) problems.push(`tarball has no package/resources/kit/${file}`);
    else if (!line.startsWith('-rwxr-xr-x')) problems.push(`tarball mode of resources/kit/${file} is ${line.split(/\s+/)[0]}, expected -rwxr-xr-x`);
  }
  return problems;
}
