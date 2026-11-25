#!/usr/bin/env node
/**
 * Clean-install check of a universal package (spec 261110-universal-package §6.1 build_package):
 *   node scripts/verify-universal-package.mjs <package.tgz>
 * npm installs it into an empty directory, then: the universal rules (no Kit, no SDK resources, no executables, no
 * os/cpu, kit-manifest.json byte for byte the committed one), capabilities.json stamped for the version, whats-new.json shape, the
 * templates capabilities advertises, and tool errors keeping their codes across the bundles.
 */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { capabilityProblems } from './capabilities.mjs';
import { PACKAGE_NAME, releaseVersion, universalProblems } from './universal-package.mjs';

const { positionals } = parseArgs({ allowPositionals: true, options: {} });
const archive = positionals[0];
if (!archive) throw new Error('usage: node scripts/verify-universal-package.mjs <package.tgz>');
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const root = await fs.mkdtemp(path.join(os.tmpdir(), 'officesdk-verify-'));
const run = (command, args, options = {}) => {
  const result = spawnSync(command, args, { encoding: 'utf8', stdio: 'inherit', ...options });
  if (result.status !== 0) throw new Error(`${command} failed with ${result.status}`);
};
try {
  // Its peers are DSH's own and never installed here (--omit=peer --legacy-peer-deps). One retry covers a passing
  // network fault. Its own npm cache, removed with the directory: npm would otherwise keep a copy of the package in ~/.npm.
  const install = ['install', '--force', '--ignore-scripts', '--omit=peer', '--legacy-peer-deps', '--no-audit', '--no-fund', '--cache', path.join(root, '.npm-cache'), '--prefix', root, path.resolve(archive)];
  try { run('npm', install); } catch { console.warn('npm install failed; retrying once'); run('npm', install); }
  const packageRoot = path.join(root, 'node_modules', ...PACKAGE_NAME.split('/'));
  const manifest = JSON.parse(await fs.readFile(path.join(packageRoot, 'package.json'), 'utf8'));
  if (!releaseVersion(manifest.version)) throw new Error(`unexpected release version ${manifest.version}`);
  // The npm build names nobody (npm records the publishing account, spec 261111-npm-publish-prep §2).
  if (!manifest.publishConfig) throw new Error('package.json has no publishConfig: not the npm build');
  if (manifest.maintainers) throw new Error('the npm build must not list maintainers');
  {
    // The layer the package says it is (the public build: full), as npm-gate reads it.
    const layer = JSON.parse(await fs.readFile(path.join(packageRoot, 'skills-catalog.json'), 'utf8').catch(() => '{}')).layer ?? 'light';
    const problems = await universalProblems(packageRoot, { kitManifest: await fs.readFile(path.join(repo, 'components/kit-manifest.json'), 'utf8'), version: manifest.version, layer });
    if (problems.length) throw new Error(`not a universal package:\n${problems.join('\n')}`);
  }
  const entryNames = [...(await fs.readFile(path.join(packageRoot, 'plugin.yml'), 'utf8')).matchAll(/^\s*name:\s*"?([^"\s]+)"?\s*$/gm)].map(match => match[1]);
  if (!entryNames.length || entryNames.some(name => name !== PACKAGE_NAME && !name.startsWith(`${PACKAGE_NAME}/`))) throw new Error(`plugin.yml entries do not import ${PACKAGE_NAME}: ${entryNames.join(', ')}`);
  if (!(await fs.readFile(path.join(packageRoot, 'lib/client.js'), 'utf8')).startsWith(`window.__ModuleLoader__.load({id:${JSON.stringify(PACKAGE_NAME)},`)) throw new Error(`lib/client.js does not register module ${PACKAGE_NAME}`);
  const capabilities = JSON.parse(await fs.readFile(path.join(packageRoot, 'capabilities.json'), 'utf8'));
  // An npm build (publishConfig) is installed from npm, from the registry it is published to: its capabilities must say so.
  const npmBuild = { installVia: 'npm', registry: manifest.publishConfig.registry };
  const problems = capabilityProblems(capabilities, { version: manifest.version, ...npmBuild });
  if (problems.length) throw new Error(`capabilities.json: ${problems.join('; ')}`);
  // Skill resources download on first use (§4.5) once the package ships the light layer only; then the packaged catalog
  // lists them. A full layer (the public build) must carry every template and style pack it advertises.
  const catalog = JSON.parse(await fs.readFile(path.join(packageRoot, 'skills-catalog.json'), 'utf8').catch(() => 'null'));
  const full = !catalog || catalog.layer === 'full';
  const claims = [
    ...['document', 'spreadsheet'].flatMap(product => (capabilities.products?.find(item => item?.id === product)?.templates ?? []).map(name => product === 'document' ? ['shimo-docx', `templates/${name}.docx`] : ['shimo-xlsx', `templates/${name}.xlsx`])),
    ...(capabilities.products?.find(item => item?.id === 'presentation')?.stylePacks ?? []).map(name => ['shimo-pptx', `styles/${name}/style.json`]),
  ];
  for (const [skill, rel] of claims) {
    // The blank workbook is opened directly by the host, even with no spreadsheet skill installed.
    if (skill === 'shimo-xlsx' && rel === 'templates/blank.xlsx') {
      await fs.access(path.join(packageRoot, rel));
      continue;
    }
    await fs.access(path.join(packageRoot, 'skills', skill, ...rel.split('/'))).catch(() => {
      if (full || !manifest.files?.includes('skills-catalog.json')) throw new Error(`capabilities advertises ${skill} ${rel}, which the package lacks`);
    });
  }
  if (full && catalog) {
    const listed = (capabilities.skills ?? []).map(item => item.name).sort();
    if (JSON.stringify(listed) !== JSON.stringify(Object.keys(catalog.skills ?? {}).sort())) throw new Error(`capabilities.json skills (${listed.join(', ')}) differ from skills-catalog.json`);
  }
  const notes = JSON.parse(await fs.readFile(path.join(packageRoot, 'whats-new.json'), 'utf8'));
  if (notes?.v !== 1 || !Array.isArray(notes.entries)) throw new Error('whats-new.json is not { v: 1, entries }');
  run(process.execPath, [path.join(repo, 'scripts/verify-tool-errors.mjs'), packageRoot]);
  console.log(`Verified clean install: ${PACKAGE_NAME}@${manifest.version} (universal)`);
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
