import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, statSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { KitDocuments, KitProcess } from '../src/index.mjs';

const FAKE = fileURLToPath(new URL('./fake-kit.mjs', import.meta.url));
export const owner = (id = 'tab-1') => ({ id, principal: 'owner', signal: new AbortController().signal });
export const guest = (id = 'tab-guest') => ({ id, principal: 'guest', signal: new AbortController().signal });

export function setup(t) {
  const base = realpathSync(mkdtempSync(path.join(tmpdir(), 'kit-docs-')));
  t.after(() => rmSync(base, { recursive: true, force: true }));
  const workspace = path.join(base, 'workspace'), outside = path.join(base, 'outside');
  mkdirSync(workspace); mkdirSync(outside);
  for (const name of ['deck.pptx', 'changed.pptx', 'locked.pptx', 'disk.pptx', 'slow.pptx', 'report.docx', 'book.xlsx', 'mismatch.docx', 'mismatch.xlsx', 'protected.docx']) writeFileSync(path.join(workspace, name), 'pptx');
  writeFileSync(path.join(outside, 'far.pptx'), 'pptx');
  symlinkSync(path.join(outside, 'far.pptx'), path.join(workspace, 'link.pptx'));
  const kit = new KitProcess({ binary: process.execPath, work: path.join(base, 'kit'), restartMs: [50], killGraceMs: 300,
    spawn: (bin, args, options) => spawn(process.execPath, [FAKE, ...args], options) });
  const authorizations = [];
  const docs = new KitDocuments({
    kit, exportRoot: path.join(base, 'exports'),
    async authorize(peer, scope, signal, { action, path: file }) {
      authorizations.push([peer.principal, action, file]);
      if (scope.sessionId === 'gone') throw Object.assign(new Error('The Session is no longer available.'), { code: 'forbidden' });
      if (peer.principal === 'guest' && action !== 'open') throw Object.assign(new Error('Access to the document was denied.'), { code: 'forbidden' });
    },
    async stat(scope, file) {
      const absolutePath = path.resolve(scope.workspaceRoot, file);
      if (!absolutePath.startsWith(scope.workspaceRoot + path.sep)) throw Object.assign(new Error('outside the workspace'), { code: 'forbidden' });
      return { absolutePath, version: existsSync(absolutePath) ? `v${statSync(absolutePath).mtimeMs}` : undefined };
    },
    async writeAccess(scope, real) {
      if (scope.readOnly) return 'denied';
      return real.startsWith(workspace + path.sep) ? 'allowed' : 'outside';
    },
  });
  t.after(() => docs.dispose());
  const scope = { sessionId: 's1', workspaceRoot: workspace };
  const commands = () => existsSync(path.join(base, 'kit', 'commands.log')) ? readFileSync(path.join(base, 'kit', 'commands.log'), 'utf8').trim().split('\n').filter(Boolean).map(line => JSON.parse(line)) : [];
  return { docs, kit, scope, workspace, outside, base, commands, authorizations };
}
export const code = (expected, details) => error => error.code === expected && Object.entries(details ?? {}).every(([key, value]) => error.details?.[key] === value);

