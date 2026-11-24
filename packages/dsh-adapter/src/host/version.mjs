/**
 * Plugin version numbers (release 0.0.0-alpha.1 on): `X.Y.Z` is a release; `X.Y.Z-alpha.N`, `-beta.N`, `-rc.N` are the
 * preview stages, in that order, before X.Y.Z; `X.Y.Z-test.N` (CI test builds of the git era) still parses and sorts
 * after rc. The old Windows forms `X.Y.Z-win32[-test.N]` sort as their release or test build.
 *
 * npm dist-tags follow the stage: `alpha`, `beta`, `rc`; `latest` for a release; `next` for a test build.
 */
export const STAGES = Object.freeze(['alpha', 'beta', 'rc', 'test']);
/** dist-tag of each stage, and of a release. */
export const STAGE_TAGS = Object.freeze({ alpha: 'alpha', beta: 'beta', rc: 'rc', test: 'next', release: 'latest' });
export const DIST_TAGS = Object.freeze(['alpha', 'beta', 'rc', 'next', 'latest']);

const VERSION = /^(\d+)\.(\d+)\.(\d+)(?:-(?:(alpha|beta|rc|test)\.(\d+)|win32(?:-(test)\.(\d+))?))?$/;

/** `{ major, minor, patch, stage, number }` (stage `release` for X.Y.Z), or undefined for anything else. */
export function parseVersion(version) {
  const match = VERSION.exec(typeof version === 'string' ? version : '');
  if (!match) return undefined;
  const stage = match[4] ?? match[6] ?? 'release';
  const number = match[5] ?? match[7];
  return { major: Number(match[1]), minor: Number(match[2]), patch: Number(match[3]), stage, number: number === undefined ? 0 : Number(number) };
}

/** A version this package may be published as: X.Y.Z or X.Y.Z-alpha|beta|rc|test.N (no -win32 forms). */
export const publishableVersion = version => /^\d+\.\d+\.\d+(?:-(?:alpha|beta|rc|test)\.\d+)?$/.test(version ?? '') && Boolean(parseVersion(version));

/** The stage of a version: alpha | beta | rc | test | release (undefined when it is not a version). */
export const stageOf = version => parseVersion(version)?.stage;
/** The npm dist-tag a version is published under. */
export const distTagOf = version => STAGE_TAGS[stageOf(version)];
/** X.Y.Z of a version. */
export const baseOf = version => { const v = parseVersion(version); return v && `${v.major}.${v.minor}.${v.patch}`; };

const rank = stage => (stage === 'release' ? STAGES.length : STAGES.indexOf(stage));
/** Negative, zero or positive like a sort comparator; NaN when either side is not a version. */
export function compareVersions(left, right) {
  const [a, b] = [parseVersion(left), parseVersion(right)];
  if (!a || !b) return NaN;
  return a.major - b.major || a.minor - b.minor || a.patch - b.patch || rank(a.stage) - rank(b.stage) || a.number - b.number;
}
export const newer = (left, right) => compareVersions(left, right) > 0;

/**
 * The dist-tags a channel takes updates from: its own and every more mature one (an alpha user also gets beta, rc and
 * the release; a test user (next) gets releases too).
 */
export function channelTags(channel) {
  if (channel === 'next') return ['next', 'latest'];
  const at = ['alpha', 'beta', 'rc', 'latest'].indexOf(channel);
  return at < 0 ? ['latest'] : ['alpha', 'beta', 'rc', 'latest'].slice(at);
}
