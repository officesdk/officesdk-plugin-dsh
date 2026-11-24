/**
 * Disk-light copies for the package builds (261112-buildoffline-slim): the offline packages are ~100 MB each and CI's
 * Runner disk is small, so whole-tree copies are hard links where the file system allows them.
 *
 * Invariant for every tree made with linkTree: its files share their bytes with the source, so nothing may rewrite a
 * file in place afterwards (fs.writeFile, fs.copyFile, fs.cp onto it); replace it with replaceFile, or fs.rm it.
 */
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';

/** Errors after which a hard link cannot work here (another mount, a file system without links, too many links). */
const NO_LINK = new Set(['EXDEV', 'EPERM', 'ENOTSUP', 'EOPNOTSUPP', 'EMLINK', 'EACCES', 'ENOSYS']);

/**
 * Recreates `from` at `to` (which must not exist): directories with their modes, symbolic links as links, files as hard
 * links, or copies when linking fails. `filter(source)` false skips an entry, as fs.cp's filter does. An inode met twice
 * in one call is copied the second time, so a packed tree never holds two names for one inode (tar would write a Link).
 * Returns { linked, copied }.
 */
export async function linkTree(from, to, { filter = () => true } = {}) {
  const counts = { linked: 0, copied: 0 };
  const seen = new Set();
  let canLink = true;
  const walk = async (source, target) => {
    const stat = await fs.lstat(source);
    if (stat.isDirectory()) {
      await fs.mkdir(target);
      for (const name of (await fs.readdir(source)).sort()) {
        const child = path.join(source, name);
        if (filter(child)) await walk(child, path.join(target, name));
      }
      await fs.chmod(target, stat.mode & 0o7777);
      return;
    }
    if (stat.isSymbolicLink()) { await fs.symlink(await fs.readlink(source), target); return; }
    const inode = `${stat.dev}:${stat.ino}`;
    if (canLink && !seen.has(inode)) {
      try { await fs.link(source, target); seen.add(inode); counts.linked++; return; }
      catch (error) { if (!NO_LINK.has(error?.code)) throw error; canLink = false; }
    }
    await fs.copyFile(source, target);
    await fs.chmod(target, stat.mode & 0o7777);
    counts.copied++;
  };
  await fs.mkdir(path.dirname(to), { recursive: true });
  await walk(from, to);
  return counts;
}

/** Writes `file` as a new inode (temporary sibling, then rename), keeping an existing file's mode: safe on linked trees. */
export async function replaceFile(file, data) {
  const mode = await fs.stat(file).then(stat => stat.mode & 0o7777, () => undefined);
  const temporary = path.join(path.dirname(file), `.${path.basename(file)}.${randomUUID()}.tmp`);
  try {
    await fs.writeFile(temporary, data);
    if (mode !== undefined) await fs.chmod(temporary, mode);
    await fs.rename(temporary, file);
  } catch (error) {
    await fs.rm(temporary, { force: true });
    throw error;
  }
}
