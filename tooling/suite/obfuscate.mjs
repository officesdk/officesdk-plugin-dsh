/**
 * Light obfuscation over the minified bundles (spec 261111-npm-publish-prep §8, §13.4). It renames local identifiers and
 * moves string literals into a rotated, shuffled array — without encoding them, so the release gate's secret and
 * company-address scans still read every string. No control-flow flattening, dead code, self-defending or debugger
 * traps: those cost start-up time and make field errors unreadable, and determined readers get past them anyway.
 * Object keys and globals are never renamed (DSH and the SDK address them by name); the host keeps every Remote method's
 * parameter names (Typert reads them from the source text).
 */
import fs from 'node:fs/promises';
import path from 'node:path';

export const LIGHT = Object.freeze({
  compact: true,
  identifierNamesGenerator: 'mangled',
  renameGlobals: false,
  renameProperties: false,
  transformObjectKeys: false,
  stringArray: true,
  stringArrayThreshold: 0.75,
  stringArrayEncoding: [],
  stringArrayRotate: true,
  stringArrayShuffle: true,
  stringArrayWrappersCount: 1,
  splitStrings: false,
  controlFlowFlattening: false,
  deadCodeInjection: false,
  selfDefending: false,
  debugProtection: false,
  disableConsoleOutput: false,
  unicodeEscapeSequence: false,
  numbersToExpressions: false,
  simplify: true,
  sourceMap: false,
  seed: 20261111,
});

/**
 * The bundles rewritten: [relative path, sourceType]. Not lib/host/index.mjs: its Remote service class (decorated
 * methods lowered by esbuild) comes out with computed method keys, and Typert finds Remote methods and their parameters
 * by their source text — measured 2026-11-11 (spec §8 fallback). The host stays minified (whitespace and syntax).
 */
export const TARGETS = Object.freeze([
  ['lib/client.js', 'script'],
  ['lib/tools/index.mjs', 'module'],
  ['lib/skills/index.mjs', 'module'],
  ['resources/kit-page/kit-page.js', 'module'],
]);

/** Legal comments (`/*! … *\/`, `@license`, `@preserve`) the bundles carry, kept at the top of the output. */
const legalComments = text => [...text.matchAll(/\/\*![\s\S]*?\*\/|\/\*[^*]*?@(?:license|preserve)[\s\S]*?\*\//g)].map(match => match[0]);

/** One bundle's text, obfuscated. `reserved`: identifier names kept as they are. */
export function obfuscateText(obfuscator, text, { sourceType = 'module', reserved = [], reservedStrings = [] } = {}) {
  const head = legalComments(text);
  const code = obfuscator.obfuscate(text, { ...LIGHT, target: sourceType === 'script' ? 'browser' : 'node', sourceType, reservedNames: reserved.map(name => `^${name}$`), reservedStrings: reservedStrings.map(name => `^${name}$`) }).getObfuscatedCode();
  return `${head.length ? `${head.join('\n')}\n` : ''}${code}\n`;
}

/**
 * Obfuscates every bundle of the unpacked package at `root`; returns `{ file: { before, after } }` sizes. `reserved`:
 * the host's Remote parameter names; `methods`: its Remote method names, kept as literal method keys.
 */
export async function obfuscateTree(root, { obfuscator, reserved = [], methods = [] }) {
  const sizes = {};
  for (const [rel, sourceType] of TARGETS) {
    const file = path.join(root, rel);
    let text;
    try { text = await fs.readFile(file, 'utf8'); } catch (error) { if (error?.code === 'ENOENT') continue; throw error; }
    let out;
    if (rel === 'lib/client.js') {
      // DSH's module loader wrapper stays literal (its id names the plugin); only the bundle inside it is rewritten.
      const open = text.indexOf('\n');
      const close = text.lastIndexOf('\nreturn module.exports;}});');
      if (!text.startsWith('window.__ModuleLoader__.load({id:') || open < 0 || close < open) throw new Error('lib/client.js does not have the module loader wrapper');
      out = `${text.slice(0, open + 1)}${obfuscateText(obfuscator, text.slice(open + 1, close), { sourceType: 'script' }).trimEnd()}${text.slice(close)}`;
    } else out = obfuscateText(obfuscator, text, { sourceType, ...rel === 'lib/host/index.mjs' && { reserved, reservedStrings: methods } });
    await fs.writeFile(file, out);
    sizes[rel] = { before: Buffer.byteLength(text), after: Buffer.byteLength(out) };
  }
  return sizes;
}
