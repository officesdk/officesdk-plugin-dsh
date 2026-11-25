#!/usr/bin/env node
/**
 * Engine-level smoke of the pinned office-kit release on this machine, through the plugin's own downloader and the
 * public download profile (no DeepSeek Harness needed): the Kit for this platform from the manifest's release sources,
 * the editor resources from npm, every hash checked by the downloader against components/kit-manifest.json; the
 * manifest's archive sha256 values against the release's SHA256SUMS; then the Kit's self-test (`--version --json`).
 *
 *   node tooling/kit/kit-smoke.mjs --data <dir> [--manifest components/kit-manifest.json] [--github-env <file>]
 *
 * Prints a JSON report; exits 1 on any failure. `--github-env` appends OFFICESDK_KIT_DIR=<the unpacked Kit> to that file
 * (GitHub Actions' $GITHUB_ENV), for packages/kit-host/test/real-kit.test.mjs in a later step.
 */
import { execFile } from 'node:child_process';
import { appendFileSync, readFileSync } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs, promisify } from 'node:util';
import { OfficeKitRuntime } from '../../packages/dsh-adapter/src/host/fetch/runtime.mjs';
import { DEFAULT_SOURCES, selectProfile } from '../../packages/dsh-adapter/src/host/fetch/download-sources.mjs';
import { parseSums } from './vendor.mjs';

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
/** An SRI `sha256-<base64>` as hex. */
export const sriHex = sri => (/^sha256-/.test(sri ?? '') ? Buffer.from(sri.slice(7), 'base64').toString('hex') : undefined);

/** The manifest's archives (each Kit and the editor tarball) that SHA256SUMS lists otherwise or not at all. */
export function sumsProblems(manifest, sums) {
  const problems = [];
  const rows = [...Object.entries(manifest.kit ?? {}).map(([platform, kit]) => [`kit ${platform}`, kit.file, kit.sha256]), ['editor', manifest.editor?.file, manifest.editor?.sha256]];
  for (const [what, file, sri] of rows) {
    const listed = sums.get(file);
    if (!listed) problems.push(`SHA256SUMS does not list ${file} (${what})`);
    else if (listed !== sriHex(sri)) problems.push(`SHA256SUMS lists ${file} as ${listed}, kit-manifest.json says ${sriHex(sri)} (${what})`);
  }
  return problems;
}

async function main() {
  const { values } = parseArgs({ options: { data: { type: 'string' }, manifest: { type: 'string', default: path.join(REPO, 'components/kit-manifest.json') }, 'github-env': { type: 'string' } } });
  if (!values.data) { console.error('usage: node tooling/kit/kit-smoke.mjs --data <dir> [--manifest <kit-manifest.json>] [--github-env <file>]'); process.exit(2); }
  const manifest = JSON.parse(readFileSync(values.manifest, 'utf8'));
  const report = { officeKit: manifest.officeKit.version, problems: [] };
  // The release's SHA256SUMS, from the first release source (the manifest's own `sources`).
  const base = manifest.sources?.[0]?.base;
  report.sums = 'unavailable';
  try {
    if (!base) throw new Error('kit-manifest.json names no release source');
    const response = await fetch(`${base}SHA256SUMS`);
    if (!response.ok) throw new Error(`${base}SHA256SUMS: HTTP ${response.status}`);
    report.problems.push(...sumsProblems(manifest, parseSums(await response.text())));
    report.sums = 'checked';
  } catch (error) { report.problems.push(error.message); }

  const dataRoot = path.resolve(values.data);
  await fs.mkdir(dataRoot, { recursive: true });
  const packageRoot = await fs.mkdtemp(path.join(dataRoot, 'empty-package-'));
  const logs = [];
  const runtime = new OfficeKitRuntime({ manifest, dataRoot, packageRoot, profile: selectProfile(DEFAULT_SOURCES, {}), log: (level, message) => logs.push(`${level}: ${message}`) });
  report.platform = runtime.platform;
  if (runtime.unsupported) { report.problems.push(`no Kit for ${runtime.platform}`); return finish(report, logs); }
  report.sources = { kit: runtime.kitItem.sources.map(source => source.base ?? source.id), editor: runtime.editorItem.sources.map(source => (source.kind === 'npm' ? `npm ${source.registry} ${source.package}@${source.version}` : source.base ?? source.id)) };
  const started = Date.now();
  try {
    const binary = await runtime.binary();
    report.seconds = Math.round((Date.now() - started) / 100) / 10;
    report.from = Object.fromEntries(Object.entries(runtime.progress).map(([key, progress]) => [key, progress.source]));
    const { stdout } = await promisify(execFile)(binary, ['--version', '--json'], { timeout: 120_000, windowsHide: true });
    const info = JSON.parse(stdout);
    report.selfTest = { version: info.version, protocol: info.compat?.protocolVersion };
    if (info.version !== manifest.officeKit.version || info.compat?.protocolVersion !== manifest.officeKit.protocol) report.problems.push(`self-test reports ${info.version}/${info.compat?.protocolVersion}`);
    report.kitDir = runtime.kitDir;
    report.editorDir = runtime.editorDir;
    if (values['github-env']) appendFileSync(values['github-env'], `OFFICESDK_KIT_DIR=${runtime.kitDir}\n`);
  } catch (error) { report.problems.push(error.message); }
  return finish(report, logs);
}

function finish(report, logs) {
  report.logs = logs;
  console.log(JSON.stringify(report, null, 2));
  process.exit(report.problems.length ? 1 : 0);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main();
