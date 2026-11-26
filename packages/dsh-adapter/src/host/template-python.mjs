import { execFile } from 'node:child_process';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { SuiteError } from '../../../core/src/index.mjs';

const execute = promisify(execFile);
const bundledSkills = fileURLToPath(new URL('../../skills/', import.meta.url));

/** Resolve the interpreter once per invocation; only bundled helpers run in the host. */
export async function runTemplateScript(skill, script, args, { cwd, signal }) {
  const runtime = join(homedir(), '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python');
  const interpreters = process.platform === 'win32'
    ? [[join(runtime, 'python.exe')], ['py', '-3'], ['python']]
    : [[join(runtime, 'bin', 'python3')], ['python3']];
  const file = join(bundledSkills, skill, 'scripts', script);
  for (const [binary, ...prefix] of interpreters) {
    signal?.throwIfAborted();
    try {
      return await execute(binary, [...prefix, '-I', '-X', 'utf8', file, ...args], {
        cwd, signal, timeout: 30000, maxBuffer: 1024 * 1024, windowsHide: true,
      });
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  throw new SuiteError('unavailable', 'The template needs Python 3. Install the DSH Python runtime or make Python 3 available, then retry.');
}
