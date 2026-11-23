// capabilities.json for partners' bots (docs/capabilities.md). No DSH packages: CI runs this file.
import test from 'node:test';
import assert from 'node:assert/strict';
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { NPM_INSTALL, REGISTRY, TAG_PATTERN, buildCapabilities, capabilityProblems, installFields, registryAddress, readSkills, readToolNames, renderCapabilities, stampCapabilities } from '../../../scripts/capabilities.mjs';
import { PLATFORMS } from '../../../scripts/platform-package.mjs';

const repo = fileURLToPath(new URL('../../..', import.meta.url));
const adapter = join(repo, 'packages/dsh-adapter');
const readJson = async file => JSON.parse(await readFile(file, 'utf8'));
const platforms = Object.keys(PLATFORMS);

test('the committed capabilities.json is what its sources make (run node scripts/whats-new.mjs after changing them)', async () => {
  assert.equal(await readFile(join(adapter, 'capabilities.json'), 'utf8'), renderCapabilities(buildCapabilities(repo, { platforms })),
    'capabilities.json is out of date: run node scripts/whats-new.mjs');
  assert.ok((await readJson(join(adapter, 'package.json'))).files.includes('capabilities.json'), 'it ships in every package, so it is at the dist root');
});

test('capabilities agree with package.json, the office-kit pin, the skills and the registered AI tools', async () => {
  const capabilities = await readJson(join(adapter, 'capabilities.json'));
  const manifest = await readJson(join(adapter, 'package.json'));
  const pin = await readJson(join(repo, 'components/office-kit.json'));
  assert.deepEqual(capabilities.requires, { dsh: manifest.engines.dsh, node: manifest.engines.node });
  assert.deepEqual(capabilities.officeKit, { version: pin.officeKit.version, ...pin.officeKit.commit && { commit: pin.officeKit.commit }, platforms: [...pin.officeKit.platforms].sort() });
  assert.deepEqual(capabilities.skills, readSkills(join(adapter, 'skills')));
  assert.deepEqual(capabilities.skills.map(skill => skill.name), ['shimo-docx', 'shimo-pptx']);
  const tools = readToolNames(join(adapter, 'src/tools'));
  // A tool may serve several products (selection_rewrite: Word and spreadsheets); every registered one belongs to one at least,
  // except spreadsheet_create, which needs a spreadsheet skill no build bundles.
  assert.deepEqual([...new Set(capabilities.products.flatMap(product => product.aiTools))].sort(), [...tools].filter(tool => tool !== 'spreadsheet_create').sort(), 'every registered tool belongs to a product');
  assert.deepEqual(capabilities.products.find(product => product.id === 'spreadsheet').templates, []);
  assert.deepEqual(capabilities.products.find(product => product.id === 'document').templates, ['gongwen']);
  assert.deepEqual(capabilities.products.find(product => product.id === 'presentation').stylePacks, ['default']);
  assert.deepEqual(capabilities.products.map(product => product.id), ['presentation', 'document', 'spreadsheet']);
  assert.ok(capabilities.products[0].stylePacks.includes('default'));
  assert.deepEqual(capabilities.platforms, [{ name: 'darwin-arm64', status: 'stable' }, { name: 'win32-x64', status: 'test' }], 'v2: platform maturity only, no repository or tag');
  assert.equal(capabilities.v, 2);
  // The plugin ships through npm only (0.0.0-alpha.1): the committed template names npmjs; a build for another registry
  // stamps that one, so no address is committed.
  assert.equal(capabilities.installVia, 'npm');
  assert.deepEqual(capabilities.install, NPM_INSTALL);
  assert.equal(capabilities.registry, 'https://registry.npmjs.org/');
  assert.deepEqual(capabilities.npm, { status: 'enabled', install: NPM_INSTALL, registry: 'https://registry.npmjs.org/' });
  assert.deepEqual(NPM_INSTALL, { stable: '@officesdk/officesdk-plugin-dsh@latest', test: '@officesdk/officesdk-plugin-dsh@next', alpha: '@officesdk/officesdk-plugin-dsh@alpha', beta: '@officesdk/officesdk-plugin-dsh@beta', rc: '@officesdk/officesdk-plugin-dsh@rc' });
  assert.deepEqual(capabilities.tagPattern, TAG_PATTERN);
  assert.equal(capabilities.skillsBaseline, true);
  assert.equal('platform' in capabilities, false);
  assert.deepEqual([capabilities.version, capabilities.channel, capabilities.highlights], [null, null, null], 'packaging fills these');
  const print = Object.fromEntries(capabilities.products.map(product => [product.id, product.capabilities.find(item => item.id === 'print')?.supported]));
  assert.deepEqual(print, { presentation: false, document: true, spreadsheet: false }, 'only Word prints (toolbar.mjs PRINT_FORMATS)');
  assert.ok(!capabilities.common.some(item => item.id === 'print'));
});

test('a missing or wrong field in the source stops the build', async () => {
  const source = await readJson(join(adapter, 'capabilities.source.json'));
  const broken = change => { const copy = structuredClone(source); change(copy); return () => buildCapabilities(repo, { platforms, source: copy }); };
  assert.throws(broken(s => { delete s.products[0].aiTools; }), /products\[0\]\.aiTools is missing/);
  assert.throws(broken(s => { delete s.products[1].capabilities[0].supported; }), /supported must be true or false/);
  assert.throws(broken(s => { s.products[2].capabilities[0].summary = ''; }), /summary is missing/);
  assert.throws(broken(s => { s.products[0].aiTools.push('presentation_fly'); }), /presentation_fly, which src\/tools does not register/);
  assert.throws(broken(s => { s.platforms[1].status = 'beta'; }), /status must be stable or test/);
  assert.throws(broken(s => { s.platforms.push({ name: 'linux-x64', status: 'test' }); }), /linux-x64 is not a platform/);
  assert.throws(broken(s => { delete s.plugin; }), /plugin is missing/);
});

/** A package directory with the committed capabilities.json and whats-new.json. */
async function packageDir({ listed = true, withFile = true } = {}) {
  const dir = await mkdtemp(join(tmpdir(), 'capabilities-'));
  await writeFile(join(dir, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version: '0', files: listed ? ['capabilities.json', 'whats-new.json'] : ['whats-new.json'] }));
  if (withFile) await copyFile(join(adapter, 'capabilities.json'), join(dir, 'capabilities.json'));
  await copyFile(join(adapter, 'whats-new.json'), join(dir, 'whats-new.json'));
  return dir;
}

test('packaging stamps the version, the channel and this version\'s highlights; promotion re-stamps only those', async () => {
  const notes = await readJson(join(adapter, 'whats-new.json'));
  // The committed CHANGELOG's only section is 0.0.1-rc.1: that build takes its (Chinese) points.
  const released = notes.entries.find(entry => entry.version === '0.0.1-rc.1')?.highlights ?? [];
  assert.equal(released.length, 2, '0.0.1-rc.1 has two highlights');
  // A 未发布 section goes to whichever version is built from it (a test build takes its release's, 0.0.50-test.3 → 0.0.50).
  const unreleased = { v: 1, entries: [{ version: 'unreleased', highlights: ['coming'] }] };
  for (const [version, channel, points, synthetic] of [['0.0.1-rc.1', 'rc', released], ['0.0.50-test.3', 'test', ['coming'], unreleased], ['0.0.50', 'stable', ['coming'], unreleased]]) {
    const dir = await packageDir();
    try {
      if (synthetic) await writeFile(join(dir, 'whats-new.json'), JSON.stringify(synthetic));
      stampCapabilities(dir, { version });
      const stamped = await readJson(join(dir, 'capabilities.json'));
      assert.equal(stamped.version, version);
      assert.equal(stamped.channel, channel);
      assert.equal('platform' in stamped, false);
      assert.deepEqual(stamped.highlights, points, `${version} takes its section's points`);
      assert.deepEqual(capabilityProblems(stamped, { version }), []);
      assert.match(capabilityProblems(stamped, { version: '0.0.51' }).join(), /not the package's 0\.0\.51/);
    } finally { await rm(dir, { recursive: true, force: true }); }
  }
  // Test build then promotion of the same files: only version, channel and highlights differ.
  const dir = await packageDir();
  try {
    const test = stampCapabilities(dir, { version: '0.1.0-test.7' });
    const stable = stampCapabilities(dir, { version: '0.1.0' });
    const strip = c => ({ ...c, version: undefined, channel: undefined, highlights: undefined });
    assert.deepEqual(strip(test), strip(stable));
  } finally { await rm(dir, { recursive: true, force: true }); }
  // The build chooses the install way once; promotion (stamping without installVia) keeps it.
  const npmBuild = await packageDir();
  try {
    const test = stampCapabilities(npmBuild, { version: '0.1.1-test.2', installVia: 'npm' });
    assert.deepEqual([test.installVia, test.install, test.registry, test.npm.status], ['npm', NPM_INSTALL, REGISTRY, 'enabled']);
    const stable = stampCapabilities(npmBuild, { version: '0.1.1' });
    assert.deepEqual([stable.installVia, stable.install, stable.registry], ['npm', NPM_INSTALL, REGISTRY], 'promotion keeps npm');
    assert.deepEqual(capabilityProblems(stable, { version: '0.1.1', installVia: 'npm' }), []);
    // A package stamped before installVia existed is installed from npm too.
    const legacy = await readJson(join(npmBuild, 'capabilities.json'));
    for (const key of ['installVia', 'npm']) delete legacy[key];
    await writeFile(join(npmBuild, 'capabilities.json'), JSON.stringify(legacy));
    assert.equal(stampCapabilities(npmBuild, { version: '0.1.1' }).installVia, 'npm');
    assert.throws(() => stampCapabilities(npmBuild, { version: '0.1.1', installVia: 'git' }), /installVia must be npm, not git/);
  } finally { await rm(npmBuild, { recursive: true, force: true }); }
  const old = await packageDir();
  try {
    // A dated section is that version's own.
    await writeFile(join(old, 'whats-new.json'), JSON.stringify({ v: 1, latestReleased: '0.0.46', entries: [{ version: '0.0.46', highlights: ['old'] }] }));
    stampCapabilities(old, { version: '0.0.46' });
    assert.deepEqual((await readJson(join(old, 'capabilities.json'))).highlights, ['old']);
  } finally { await rm(old, { recursive: true, force: true }); }
});

test('a package that lists capabilities.json must have it; one from before it is left alone', async () => {
  const missing = await packageDir({ withFile: false });
  const before = await packageDir({ listed: false, withFile: false });
  try {
    assert.throws(() => stampCapabilities(missing, { version: '0.0.47' }), /lists capabilities\.json but the package has none/);
    assert.equal(stampCapabilities(before, { version: '0.0.46' }), undefined);
    assert.match(capabilityProblems({ v: 2, version: '0.0.47', channel: 'stable' }, { version: '0.0.47' }).join(), /highlights missing/);
    assert.match(capabilityProblems({ v: 1, version: '0.0.47', channel: 'stable', platform: { name: 'darwin-arm64' } }, { version: '0.0.47' }).join(), /v must be 2.*platform must not be set/);
  } finally { await rm(missing, { recursive: true, force: true }); await rm(before, { recursive: true, force: true }); }
});

test('an npm build for another registry names that registry, and promotion keeps it', async () => {
  const LAN = 'http://192.0.2.10:4873';
  assert.equal(registryAddress(LAN), `${LAN}/`);
  assert.equal(registryAddress(`${LAN}//`), `${LAN}/`);
  assert.throws(() => registryAddress('192.0.2.10:4873'), /not an http\(s\) URL/);
  const dir = await packageDir();
  try {
    const test = stampCapabilities(dir, { version: '0.1.0-test.1', installVia: 'npm', registry: LAN });
    assert.deepEqual([test.installVia, test.install, test.registry, test.npm], ['npm', NPM_INSTALL, `${LAN}/`, { status: 'enabled', install: NPM_INSTALL, registry: `${LAN}/` }]);
    const stable = stampCapabilities(dir, { version: '0.1.0' });
    assert.equal(stable.registry, `${LAN}/`, 'promotion keeps the registry the build was made for');
    assert.deepEqual(capabilityProblems(stable, { version: '0.1.0', installVia: 'npm', registry: LAN }), []);
    assert.match(capabilityProblems(stable, { version: '0.1.0', installVia: 'npm', registry: REGISTRY }).join(), /registry http:\/\/192\.0\.2\.10:4873\/ is not the package's https:\/\/registry\.npmjs\.org\//, 'verify: publishConfig.registry and capabilities agree');
    assert.match(capabilityProblems({ ...stable, npm: { ...stable.npm, registry: REGISTRY } }, { version: '0.1.0' }).join(), /do not match installVia npm/);
    const committed = await readJson(join(adapter, 'capabilities.json'));
    assert.equal(JSON.stringify(committed).includes('192.0.2.10') || /\d+\.\d+\.\d+\.\d+:4873/.test(JSON.stringify(committed)), false, 'no LAN address in the committed file');
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('install is npm: other ways are refused, and a broken install block is found', async () => {
  assert.throws(() => installFields('pnpm'), /installVia must be npm/);
  assert.throws(() => installFields('git'), /installVia must be npm/);
  const stamped = await packageDir();
  try {
    const good = stampCapabilities(stamped, { version: '0.1.0', installVia: 'npm' });
    assert.deepEqual(capabilityProblems(good, { version: '0.1.0' }), []);
    assert.match(capabilityProblems({ ...good, install: { stable: 'x' } }, { version: '0.1.0' }).join(), /do not match installVia npm/);
    assert.match(capabilityProblems({ ...good, installVia: 'git' }, { version: '0.1.0' }).join(), /installVia git is not npm/);
    assert.match(capabilityProblems({ ...good, installVia: undefined }, { version: '0.1.0' }).join(), /installVia undefined is not npm/);
    const committed = await readJson(join(adapter, 'capabilities.json'));
    assert.deepEqual(buildCapabilities(repo, { platforms, installVia: 'npm' }), { ...committed, ...installFields('npm') }, 'only the install fields differ');
  } finally { await rm(stamped, { recursive: true, force: true }); }
});
