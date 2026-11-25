// Plugin version numbers and channels (release 0.0.0-alpha.1 on): stages alpha < beta < rc < test < release.
import test from 'node:test';
import assert from 'node:assert/strict';
import { baseOf, channelTags, compareVersions, distTagOf, newer, parseVersion, publishableVersion, stageOf } from '../src/host/version.mjs';
import { ChannelStore, channelOf, channelVersion } from '../src/host/npm.mjs';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

test('versions: releases, preview stages, test builds and the old Windows forms parse; anything else does not', () => {
  assert.deepEqual(parseVersion('0.0.0-alpha.1'), { major: 0, minor: 0, patch: 0, stage: 'alpha', number: 1 });
  assert.deepEqual(parseVersion('0.1.2'), { major: 0, minor: 1, patch: 2, stage: 'release', number: 0 });
  assert.equal(parseVersion('0.0.47-win32-test.2').stage, 'test');
  assert.equal(parseVersion('0.0.47-win32').stage, 'release');
  for (const bad of ['0.0.0-alpha', '0.0.0-gamma.1', 'v0.0.1', '0.0', '0.0.0-alpha.1.2', undefined]) assert.equal(parseVersion(bad), undefined, String(bad));
  assert.ok(publishableVersion('0.0.0-rc.3') && publishableVersion('1.2.3') && publishableVersion('0.1.0-test.4'));
  assert.ok(!publishableVersion('0.0.47-win32') && !publishableVersion('0.0.47-win32-test.1'), 'the old Windows forms are not published any more');
  assert.equal(baseOf('0.0.0-beta.2'), '0.0.0');
});

test('ordering: alpha < beta < rc < test < release within one X.Y.Z, numbers within a stage; X.Y.Z first', () => {
  const sorted = ['0.0.1', '0.0.0', '0.0.0-rc.1', '0.0.0-alpha.10', '0.0.0-beta.1', '0.0.0-alpha.2', '0.0.0-test.3', '0.0.50'].sort(compareVersions);
  assert.deepEqual(sorted, ['0.0.0-alpha.2', '0.0.0-alpha.10', '0.0.0-beta.1', '0.0.0-rc.1', '0.0.0-test.3', '0.0.0', '0.0.1', '0.0.50']);
  assert.ok(newer('0.0.50', '0.0.0-alpha.1'), '0.0.0-alpha.1 sorts below 0.0.50 (the git releases): no update reaches it from there');
  assert.ok(newer('0.0.0-alpha.2', '0.0.0-alpha.1') && !newer('0.0.0-alpha.1', '0.0.0-alpha.1'));
  assert.ok(Number.isNaN(compareVersions('x', '0.0.1')));
});

test('dist-tags follow the stage; a channel takes its own tag and every more mature one', () => {
  assert.deepEqual(['0.0.0-alpha.1', '0.0.0-beta.1', '0.0.0-rc.1', '0.1.0-test.2', '0.1.0'].map(distTagOf), ['alpha', 'beta', 'rc', 'next', 'latest']);
  assert.equal(stageOf('0.0.0-alpha.1'), 'alpha');
  assert.deepEqual(channelTags('alpha'), ['alpha', 'beta', 'rc', 'latest']);
  assert.deepEqual(channelTags('rc'), ['rc', 'latest']);
  assert.deepEqual(channelTags('next'), ['next', 'latest']);
  assert.deepEqual(channelTags('latest'), ['latest']);
  assert.equal(channelOf('0.0.0-alpha.1'), 'alpha');
  assert.equal(channelOf('nonsense'), 'latest');
  const packument = tags => ({ 'dist-tags': tags });
  assert.equal(channelVersion(packument({ alpha: '0.0.0-alpha.2' }), 'alpha', newer), '0.0.0-alpha.2', 'no latest during the alpha stage');
  assert.equal(channelVersion(packument({ alpha: '0.0.0-alpha.3', beta: '0.0.0-beta.1' }), 'alpha', newer), '0.0.0-beta.1', 'an alpha user moves on to beta');
  assert.equal(channelVersion(packument({ alpha: '0.0.0-alpha.3', beta: '0.0.0-beta.1' }), 'beta', newer), '0.0.0-beta.1');
  assert.equal(channelVersion(packument({ alpha: '0.0.0-alpha.3', latest: '0.0.0-alpha.3' }), 'latest', newer), '0.0.0-alpha.3');
  assert.equal(channelVersion(packument({ beta: '0.0.0-beta.1' }), 'latest', newer), undefined, 'a release user is not offered a preview');
  assert.equal(channelVersion(packument({ next: '0.1.0-test.2', latest: '0.1.0' }), 'next', newer), '0.1.0');
});

test('the profile channel of a preview install is its stage, kept after it receives a more mature version', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'channels-'));
  try {
    const store = new ChannelStore(dir);
    assert.equal(await store.resolve('/p', '0.0.0-alpha.1', {}), 'alpha');
    assert.equal(await store.resolve('/p', '0.0.0-beta.1', {}), 'alpha', 'stored from then on');
    assert.equal(await store.resolve('/p', '0.0.0-beta.1', { OFFICESDK_CHANNEL: 'beta' }), 'beta', 'OFFICESDK_CHANNEL changes it');
    assert.equal(await store.resolve('/q', '0.0.0-beta.1', { OFFICESDK_CHANNEL: 'gamma' }), 'beta', 'an unknown channel is ignored');
  } finally { await rm(dir, { recursive: true, force: true }); }
});
