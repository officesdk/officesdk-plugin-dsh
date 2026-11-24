/**
 * The package steps that need only this repository and esbuild: bundled adapter code and the Kit editor page.
 * tooling/suite/pack-universal.mjs runs bundleAdapter and bundleKitPage to build the universal package on any machine.
 */
import fs from 'node:fs/promises';
import path from 'node:path';

/** The third-party packages an esbuild metafile bundled (`node_modules/<name>` inputs): their names, sorted. */
export const bundledPackages = metafile => [...new Set(Object.keys(metafile?.inputs ?? {}).map(input => /(?:^|\/)node_modules\/((?:@[^/]+\/)?[^/]+)\//.exec(input)?.[1]).filter(Boolean))].sort();

/**
 * Remote signature check: Typert reads a Remote method's parameter names from its source text, so every `async
 * method(params)` of src/host/index.mjs must appear verbatim in the bundled ShimoPreview class. Returns the problems.
 */
export function remoteSignatureProblems(source, bundled) {
  const methods = /const METHODS = (\[[^\]]*\]);/.exec(source);
  if (!methods) return ['src/host/index.mjs has no one-line `const METHODS = [...]`; the Remote signature check cannot run'];
  // Only the Remote service's own class body counts: another class with the same signature must not hide a rename.
  // esbuild with keepNames writes `ShimoPreview=class extends …`; the next class declaration ends it.
  const head = /(?:class\s+ShimoPreview|ShimoPreview\s*=\s*class)\s+extends\b/.exec(bundled);
  if (!head) return ['the host bundle has no class ShimoPreview'];
  const body = head.index + head[0].length;
  const next = bundled.slice(body).search(/\bclass(?:\s+[$\w]+)?\s*(?:extends\b|\{)/);
  const service = bundled.slice(head.index, next < 0 ? undefined : body + next).replace(/\s+/g, '');
  const problems = [];
  for (const method of JSON.parse(methods[1].replaceAll("'", '"'))) {
    const signature = new RegExp(`^\\s*async ${method}\\(([^)]*)\\)`, 'm').exec(source)?.[0].replace(/\s+/g, '');
    if (!signature) { problems.push(`Remote method ${method} is not an \`async ${method}(...)\` in src/host/index.mjs`); continue; }
    if (service.split(signature).length !== 2) problems.push(`Remote method ${method} lost its parameter names in the bundle; rename the clashing top-level identifier`);
  }
  return problems;
}

/** The Remote method names of src/host/index.mjs (`const METHODS = [...]`). */
export const remoteMethods = source => JSON.parse((/const METHODS = (\[[^\]]*\]);/.exec(source)?.[1] ?? '[]').replaceAll("'", '"'));

/** The parameter names of every Remote method (kept verbatim by any later rewriting of the host bundle). */
export function remoteParameterNames(source) {
  const methods = JSON.parse((/const METHODS = (\[[^\]]*\]);/.exec(source)?.[1] ?? '[]').replaceAll("'", '"'));
  const names = new Set();
  for (const method of methods) for (const part of (new RegExp(`^\\s*async ${method}\\(([^)]*)\\)`, 'm').exec(source)?.[1] ?? '').split(',')) { const name = part.trim().split(/[\s=]/)[0]; if (name) names.add(name); }
  return [...names].sort();
}

/**
 * lib/host, lib/tools, lib/skills (ESM, every @deepseek-ai package external, names kept) and lib/client.js. Returns the
 * third-party packages bundled into them (THIRD-PARTY-NOTICES.md).
 */
export async function bundleAdapter({ esbuild, repo }) {
  const adapter = path.join(repo, 'packages/dsh-adapter');
  const lib = path.join(adapter, 'lib');
  await fs.rm(lib, { recursive: true, force: true });
  const bundled = new Set();
  for (const [entry, out] of [['src/host/index.mjs', 'host/index.mjs'], ['src/tools/index.mjs', 'tools/index.mjs'], ['src/skills/index.mjs', 'skills/index.mjs']]) {
    const result = await esbuild({ metafile: true, entryPoints: [path.join(adapter, entry)], outfile: path.join(lib, out), bundle: true, platform: 'node', format: 'esm', target: 'node22', packages: 'external', legalComments: 'inline', sourcemap: false, minifyWhitespace: true, minifySyntax: true, keepNames: true });
    for (const name of bundledPackages(result.metafile)) bundled.add(name);
  }
  // Typert reads a Remote method's parameter names from its source text, so the bundle must keep them
  // verbatim: esbuild renames a parameter that shadows a top-level name from another module (path → path4).
  const problems = remoteSignatureProblems(await fs.readFile(path.join(adapter, 'src/host/index.mjs'), 'utf8'), await fs.readFile(path.join(lib, 'host/index.mjs'), 'utf8'));
  if (problems.length) {
    await fs.rm(lib, { recursive: true, force: true });
    throw new Error(problems.join('\n'));
  }
  const client = await esbuild({ metafile: true, entryPoints: [path.join(adapter, 'src/client/index.mjs')], bundle: true, format: 'cjs', platform: 'browser', external: ['react'], write: false, target: 'es2022', minify: true, keepNames: true, legalComments: 'inline', sourcemap: false });
  await fs.writeFile(path.join(lib, 'client.js'), `window.__ModuleLoader__.load({id:'officesdk',factory(require){const module={exports:{}};const exports=module.exports;\n${client.outputFiles[0].text}\nreturn module.exports;}});\n`);
  for (const name of bundledPackages(client.metafile)) bundled.add(name);
  return [...bundled].sort();
}

/** The Kit editor page: our code only; the editor bundle is loaded by URL at run time. Returns the bundled third-party packages. */
export async function bundleKitPage({ esbuild, repo }) {
  const kitPage = path.join(repo, 'packages/dsh-adapter/resources/kit-page');
  await fs.rm(kitPage, { recursive: true, force: true });
  await fs.mkdir(kitPage, { recursive: true });
  const page = await esbuild({ metafile: true, entryPoints: [path.join(repo, 'packages/core/src/browser/kit-page-entry.mjs')], bundle: true, format: 'esm', platform: 'browser', target: 'es2022', minify: true, write: false, legalComments: 'inline', sourcemap: false });
  await fs.writeFile(path.join(kitPage, 'kit-page.js'), page.outputFiles[0].text);
  await fs.copyFile(path.join(repo, 'packages/core/src/browser/kit-page.html'), path.join(kitPage, 'index.html'));
  return bundledPackages(page.metafile);
}
