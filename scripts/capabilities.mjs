/**
 * capabilities.json: what this plugin version is and can do, for machines (it ships at the package root).
 * Schema: docs/capabilities.md. Made by `node scripts/whats-new.mjs` together with whats-new.json from one source:
 *   packages/dsh-adapter/capabilities.source.json   plugin, platforms and their status, per-product capabilities, AI tools
 *   packages/dsh-adapter/package.json               engines.dsh / engines.node
 *   components/office-kit.json                      office-kit release (and commit, when it names one) and Kit platforms
 *   packages/dsh-adapter/skills/<name>/SKILL.md     bundled skills and their metadata.version
 *   skills/shimo-pptx/styles, shimo-docx/templates   style packs and templates
 * The committed file leaves `version`, `channel` and `highlights` null: the package gets them when it is versioned
 * (tooling/suite/pack-universal.mjs calls stampCapabilities).
 *
 * v2 (spec 261110-universal-package §7.3): one package serves every platform, so there is no `platform`; `channel`
 * (test | stable) follows the version; `platforms[]` only says how mature each platform is; where to install and how
 * tags look is `install` / `tagPattern`; skill fields describe the skills this plugin version ships (`skillsBaseline`),
 * (the public build ships its skills complete).
 */
import fs from 'node:fs';
import path from 'node:path';
import { notesBetween, readNotes } from '../packages/dsh-adapter/src/host/whats-new.mjs';
import { stageOf } from '../packages/dsh-adapter/src/host/version.mjs';

export const STATUSES = ['stable', 'test'];
export const FORMAT = 2;
// `install` says how the release is installed (`dsh plugin --profile <p> add <value>`): `npm`, one dist-tag per channel,
// from `registry` — npmjs unless the build names another (pack-universal --registry); no address is committed.
// One install line per channel; the preview stages are dist-tags of their own (release 0.0.0-alpha.1 on).
export const NPM_INSTALL = Object.freeze({ stable: '@officesdk/officesdk-plugin-dsh@latest', test: '@officesdk/officesdk-plugin-dsh@next', alpha: '@officesdk/officesdk-plugin-dsh@alpha', beta: '@officesdk/officesdk-plugin-dsh@beta', rc: '@officesdk/officesdk-plugin-dsh@rc' });
export const REGISTRY = 'https://registry.npmjs.org/';
export const INSTALL_VIA = Object.freeze(['npm']);
const httpUrl = value => { try { return ['http:', 'https:'].includes(new URL(value).protocol); } catch { return false; } };
/** A registry address with exactly one trailing slash. */
export const registryAddress = value => { if (!httpUrl(value)) fail(`registry ${value} is not an http(s) URL`); return value.replace(/\/*$/, '/'); };
/** The install fields for `via` (npm); `registry` is where npm installs come from (default npmjs). */
export function installFields(via, { registry = REGISTRY } = {}) {
  if (!INSTALL_VIA.includes(via)) fail(`installVia must be ${INSTALL_VIA.join(' or ')}, not ${via}`);
  const address = registryAddress(registry);
  return { installVia: via, install: NPM_INSTALL, registry: address, npm: { status: 'enabled', install: NPM_INSTALL, registry: address } };
}
/** Literal shapes of the release tags, not values. */
export const TAG_PATTERN = Object.freeze({ stable: 'vX.Y.Z', test: 'vX.Y.Z-test.N' });
/** stable for X.Y.Z, test for X.Y.Z-test.N, the stage for a preview (alpha | beta | rc). */
export const channelOf = version => { const stage = stageOf(version); return stage === 'release' || !stage ? 'stable' : stage; };
const fail = message => { throw new Error(`capabilities: ${message}`); };
const need = (value, what) => { if (value === undefined || value === null || value === '' || (Array.isArray(value) && !value.length)) fail(`${what} is missing`); return value; };
const readJson = file => { try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch (error) { return fail(`${file} unreadable: ${error.message}`); } };
const dirNames = dir => fs.readdirSync(dir, { withFileTypes: true }).filter(entry => entry.isDirectory()).map(entry => entry.name).sort();
const fileStems = (dir, ext) => fs.readdirSync(dir).filter(name => name.endsWith(ext)).map(name => name.slice(0, -ext.length)).sort();

/** Bundled skills: `name` and `metadata.version` from each SKILL.md front matter. */
export function readSkills(skillsRoot) {
  return dirNames(skillsRoot).filter(name => fs.existsSync(path.join(skillsRoot, name, 'SKILL.md'))).map(dir => {
    const front = /^---\r?\n([\s\S]*?)\r?\n---/.exec(fs.readFileSync(path.join(skillsRoot, dir, 'SKILL.md'), 'utf8'))?.[1] ?? '';
    return { name: need(/^name:\s*(\S+)/m.exec(front)?.[1], `skills/${dir}/SKILL.md name`), version: need(/^\s+version:\s*(\S+)/m.exec(front)?.[1], `skills/${dir}/SKILL.md metadata.version`) };
  });
}

/** AI tool names the plugin registers (`name: '<tool>'` in src/tools). */
export function readToolNames(toolsRoot) {
  return fs.readdirSync(toolsRoot).filter(name => name.endsWith('.mjs'))
    .flatMap(name => [...fs.readFileSync(path.join(toolsRoot, name), 'utf8').matchAll(/\bname: '([a-z]+_[a-z_]+)'/g)].map(match => match[1]));
}

function checkCapability(item, where) {
  need(item?.id, `${where}.id`);
  if (typeof item.supported !== 'boolean') fail(`${where}.supported must be true or false`);
  need(item.summary, `${where}.summary`);
  return { id: item.id, supported: item.supported, summary: item.summary };
}

/** capabilities.json content (unstamped) from the repository at `repo` (`source` replaces capabilities.source.json). Throws on any missing field. */
export function buildCapabilities(repo, { platforms: knownPlatforms, source: given, installVia = 'npm', registry } = {}) {
  const adapter = path.join(repo, 'packages/dsh-adapter');
  const source = given ?? readJson(path.join(adapter, 'capabilities.source.json'));
  const manifest = readJson(path.join(adapter, 'package.json'));
  const pin = readJson(path.join(repo, 'components/office-kit.json'));
  const tools = new Set(readToolNames(path.join(adapter, 'src/tools')));
  const skillsRoot = path.join(adapter, 'skills');
  const plugin = need(source.plugin, 'plugin');
  const platforms = need(source.platforms, 'platforms').map((item, index) => {
    need(item?.name, `platforms[${index}].name`);
    if (knownPlatforms && !knownPlatforms.includes(item.name)) fail(`platforms[${index}].name ${item.name} is not a platform the packaging scripts build`);
    if (!STATUSES.includes(item.status)) fail(`platforms[${index}].status must be ${STATUSES.join(' or ')}`);
    return { name: item.name, status: item.status };
  });
  const extras = {
    presentation: { stylePacks: dirNames(path.join(skillsRoot, 'shimo-pptx/styles')) },
    document: { templates: fileStems(path.join(skillsRoot, 'shimo-docx/templates'), '.docx') },
    // No spreadsheet skill is bundled: spreadsheets are read and edited, never created from a template.
    spreadsheet: { templates: fs.existsSync(path.join(skillsRoot, 'shimo-xlsx/templates')) ? fileStems(path.join(skillsRoot, 'shimo-xlsx/templates'), '.xlsx') : [] },
  };
  const withoutSkill = { spreadsheet: fs.existsSync(path.join(skillsRoot, 'shimo-xlsx/SKILL.md')) ? [] : ['spreadsheet_create'] };
  const products = need(source.products, 'products').map((product, index) => {
    const where = `products[${index}]`;
    need(product?.id, `${where}.id`);
    const aiTools = need(product.aiTools, `${where}.aiTools`).filter(tool => !(withoutSkill[product.id] ?? []).includes(tool));
    for (const tool of aiTools) if (!tools.has(tool)) fail(`${where}.aiTools names ${tool}, which src/tools does not register`);
    return {
      id: product.id, name: need(product.name, `${where}.name`), extensions: need(product.extensions, `${where}.extensions`),
      capabilities: need(product.capabilities, `${where}.capabilities`).map((item, i) => checkCapability(item, `${where}.capabilities[${i}]`)),
      aiTools, ...extras[product.id],
    };
  });
  const officeKitVersion = need(pin.officeKit?.version, 'office-kit version in components/office-kit.json officeKit.version');
  const kitPlatforms = [...need(pin.officeKit?.platforms, 'office-kit platforms in components/office-kit.json')].sort();
  return {
    v: FORMAT,
    plugin: { id: need(plugin.id, 'plugin.id'), name: need(plugin.name, 'plugin.name') },
    version: null,
    channel: null,
    platforms,
    ...installFields(installVia, { registry }),
    tagPattern: TAG_PATTERN,
    requires: { dsh: need(manifest.engines?.dsh, 'package.json engines.dsh'), node: need(manifest.engines?.node, 'package.json engines.node') },
    // office-kit 0.3 releases name no commit; an older pin that names one keeps it.
    officeKit: { version: officeKitVersion, ...pin.officeKit?.commit && { commit: pin.officeKit.commit }, platforms: kitPlatforms },
    common: need(source.common, 'common').map((item, i) => checkCapability(item, `common[${i}]`)),
    products,
    skills: need(readSkills(skillsRoot), 'skills'),
    skillsBaseline: true,
    highlights: null,
    changelog: 'CHANGELOG.md',
    whatsNew: 'whats-new.json',
  };
}

export const renderCapabilities = capabilities => `${JSON.stringify(capabilities, null, 2)}\n`;

/** Problems with a stamped capabilities.json for a package at `version`; empty when fine. */
export function capabilityProblems(capabilities, { version, installVia, registry } = {}) {
  const problems = [];
  if (capabilities?.v !== FORMAT) problems.push(`v must be ${FORMAT}`);
  if (capabilities?.version !== version) problems.push(`version ${capabilities?.version} is not the package's ${version}`);
  if (capabilities?.channel !== channelOf(version)) problems.push(`channel ${capabilities?.channel} is not ${channelOf(version)}`);
  if ('platform' in (capabilities ?? {})) problems.push('platform must not be set: the package serves every platform');
  if (!Array.isArray(capabilities?.highlights)) problems.push('highlights missing');
  for (const key of ['plugin', 'platforms', 'install', 'tagPattern', 'requires', 'officeKit', 'common', 'products', 'skills']) if (!capabilities?.[key]) problems.push(`${key} missing`);
  if (!INSTALL_VIA.includes(capabilities?.installVia)) problems.push(`installVia ${capabilities?.installVia} is not ${INSTALL_VIA.join(' or ')}`);
  else if (!httpUrl(capabilities.npm?.registry)) problems.push(`npm.registry ${capabilities.npm?.registry} is not an http(s) URL`);
  else if (JSON.stringify(['installVia', 'install', 'registry', 'npm'].map(key => capabilities[key])) !== JSON.stringify(Object.values(installFields(capabilities.installVia, { registry: capabilities.npm.registry })))) problems.push(`install, registry and npm do not match installVia ${capabilities.installVia}`);
  if (installVia && capabilities?.installVia !== installVia) problems.push(`installVia ${capabilities?.installVia} is not this build's ${installVia}`);
  if (registry && capabilities?.registry !== registryAddress(registry)) problems.push(`registry ${capabilities?.registry} is not the package's ${registryAddress(registry)}`);
  return problems;
}


/**
 * Writes this package's version, channel (channelOf: stable, test, alpha, beta, rc) and this version's highlights (from the package's
 * own whats-new.json) into its capabilities.json, and how it is installed (`installVia` and the npm `registry`: given by
 * pack-universal, else kept).
 * A package without capabilities.json (before 0.0.47) is left alone unless its package.json lists the file. Promotion
 * re-stamps only version, channel and highlights (spec §6.1): the install way stays the one the build chose.
 */
export function stampCapabilities(packageRoot, { version, installVia, registry }) {
  const file = path.join(packageRoot, 'capabilities.json');
  const manifest = readJson(path.join(packageRoot, 'package.json'));
  if (!fs.existsSync(file)) {
    if (manifest.files?.includes('capabilities.json')) fail('package.json lists capabilities.json but the package has none');
    return undefined;
  }
  const capabilities = readJson(file);
  need(version, 'version');
  capabilities.version = version;
  capabilities.channel = channelOf(version);
  delete capabilities.platform;
  // Given (pack-universal), else what the build chose: npm.
  Object.assign(capabilities, installFields(installVia ?? 'npm', { registry: registry ?? (httpUrl(capabilities.npm?.registry) ? capabilities.npm.registry : undefined) }));
  capabilities.highlights = notesBetween(readNotes(path.join(packageRoot, 'whats-new.json')), undefined, version)[0]?.highlights ?? [];
  const problems = capabilityProblems(capabilities, { version });
  if (problems.length) fail(problems.join('; '));
  fs.writeFileSync(file, renderCapabilities(capabilities));
  return capabilities;
}
