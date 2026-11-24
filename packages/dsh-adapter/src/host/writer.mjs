import { createHash } from 'node:crypto';
import { lstat, mkdir, realpath, stat } from 'node:fs/promises';
import { dirname, isAbsolute, join, resolve, sep } from 'node:path';
import { SuiteError, fail } from '../../../core/src/index.mjs';
import { fromWorkspaceFile } from './errors.mjs';

const LOCAL_FILESYSTEMS = new Set(['LocalFileSystem', 'SandboxedFileSystem']);

/** Absolute paths from WorkspaceFiles are host paths only for the local providers. */
export function isLocalFileSystem(fs) {
  for (let proto = fs && Object.getPrototypeOf(fs); proto && proto !== Object.prototype; proto = Object.getPrototypeOf(proto)) {
    if (LOCAL_FILESYSTEMS.has(proto.constructor?.name)) return true;
  }
  return false;
}

/** realpath of the deepest existing ancestor, as fs-sandbox's resolve does, so new files are checked too. */
export async function canonicalTarget(path) {
  let current = path, suffix = [];
  for (;;) {
    try { return join(await realpath(current), ...suffix.reverse()); }
    catch (error) {
      if (error.code !== 'ENOENT' || dirname(current) === current) throw error;
      suffix.push(current.slice(dirname(current).length + 1));
      current = dirname(current);
    }
  }
}
const under = (target, root) => target === root || target.startsWith(root.endsWith(sep) ? root : root + sep);

/**
 * Binary write-back under the same rules as dsh-fs-sandbox's checkedTarget: a live Session's
 * policy, writable roots from dsh-sandbox, a version check right before the atomic replace.
 * Locks live in the plugin data directory so the workspace never sees a .lock file.
 */
export function createWriter({ ctx, workspaceFiles, dataRoot, writableRoots, atomic }) {
  async function policyFor(scope) {
    const sandboxPolicy = ctx.get('sandboxPolicy'), sessions = ctx.get('sessions');
    if (!sandboxPolicy || !sessions) fail('unsupported-operation', 'This host cannot write files.');
    const session = sessions.get(scope.sessionId);
    if (!session) fail('sandbox-denied', 'Reopen the Session to save: its write policy is not loaded.');
    return sandboxPolicy.resolve({ session });
  }
  /** `approved` is one user consent to write this target although it lies outside the writable roots. */
  async function checked(scope, path, { approved = false } = {}) {
    if (!isLocalFileSystem(ctx.get('fs'))) fail('unsupported-operation', 'This filesystem does not support saving here.');
    const policy = await policyFor(scope);
    const absolute = isAbsolute(path) ? path : resolve(scope.workspaceRoot, path);
    const target = await canonicalTarget(absolute);
    if (policy.mode === 'read-only') fail('sandbox-denied', 'The Session is read only.');
    if (policy.mode === 'workspace-write') {
      const roots = await Promise.all(writableRoots(policy).map(root => canonicalTarget(root).catch(() => root)));
      if (!approved && !roots.some(root => under(target, root))) fail('sandbox-denied', 'The file is outside the workspace, so saving it needs the user\'s approval. The changes stay in the editor until then.', { outside: true });
    } else if (policy.mode !== 'danger-full-access') fail('sandbox-denied', 'Unknown sandbox mode.');
    return target;
  }
  async function locked(target, operation) {
    const lock = join(dataRoot, 'locks', createHash('sha256').update(target).digest('hex'));
    await mkdir(dirname(lock), { recursive: true, mode: 0o700 });
    return atomic.withFileLock(lock, operation);
  }
  const current = async (scope, path, signal) => { try { return await workspaceFiles.stat(scope, path, signal); } catch (error) { throw fromWorkspaceFile(error); } };
  return {
    async replace(scope, path, data, { expectedVersion, approved = false }, signal) {
      const before = await current(scope, path, signal);
      const target = await checked(scope, before.absolutePath, { approved });
      const probe = await lstat(target);
      if (!probe.isFile()) fail('invalid-path', 'Only regular files can be saved.');
      return locked(target, async () => {
        const latest = await current(scope, path, signal);
        if (latest.version !== expectedVersion || latest.absolutePath !== before.absolutePath) fail('source-changed', 'The document changed before it could be saved.');
        signal.throwIfAborted();
        await atomic.writeFileAtomic(target, data, { mode: (await stat(target)).mode & 0o777 });
        const written = await current(scope, path, signal);
        return { version: written.version, absolutePath: written.absolutePath };
      });
    },
    async create(scope, path, data, signal, { approved = false } = {}) {
      const target = await checked(scope, path, { approved });
      return locked(target, async () => {
        if (await lstat(target).then(() => true, error => { if (error.code === 'ENOENT') return false; throw error; })) fail('exists', 'A file with that name already exists.');
        signal.throwIfAborted();
        await atomic.writeFileAtomic(target, data, { mode: 0o644 });
        const written = await current(scope, target, signal).catch(() => undefined);
        return { absolutePath: written?.absolutePath ?? target, version: written?.version };
      });
    },
    /**
     * 'allowed' saves without asking; 'outside' needs one user approval (outside the writable roots of a
     * workspace-write Session); 'denied' cannot be saved from this Session (read only, policy not loaded).
     * Any other failure surfaces on the save itself.
     */
    async writeAccess(scope, path) {
      try { await checked(scope, (await current(scope, path, AbortSignal.timeout(10_000))).absolutePath); return 'allowed'; }
      catch (error) { if (error?.code === 'sandbox-denied') return error.details?.outside ? 'outside' : 'denied'; throw error; }
    },
    /**
     * The same answer for a host path Kit will write (already a real path): checked against the Session's
     * policy directly, without WorkspaceFiles, so a file outside the workspace reads as 'outside'. A
     * filesystem or host that cannot write here is 'denied'.
     */
    async accessFor(scope, absolutePath) {
      if (!isAbsolute(absolutePath)) return 'denied';
      try { await checked(scope, absolutePath); return 'allowed'; }
      catch (error) {
        if (error?.code === 'sandbox-denied') return error.details?.outside ? 'outside' : 'denied';
        if (error?.code === 'unsupported-operation') return 'denied';
        throw error;
      }
    },
    canWrite: () => isLocalFileSystem(ctx.get('fs')) && !!ctx.get('sandboxPolicy') && !!ctx.get('sessions'),
  };
}

export { SuiteError };
