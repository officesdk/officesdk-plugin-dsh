import { crc32 } from 'node:zlib';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { SuiteError, SuiteService } from '../src/index.mjs';

export const scope = { sessionId: 'session-a', workspaceRoot: '/authorized' };
export const signal = () => new AbortController().signal;
export const peerOf = (id = 'peer-a', principal = 'owner') => ({ id, principal, signal: new AbortController().signal });
export const MAIN = { docx: 'word/document.xml', xlsx: 'xl/workbook.xml', pptx: 'ppt/presentation.xml' };

/** Stored (uncompressed) ZIP: enough for Office part checks and MOW packages. */
export function zip(files) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const data = Buffer.from(content), nameBytes = Buffer.from(name), crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x800, 6);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(data.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(nameBytes.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0x800, 8);
    central.writeUInt32LE(crc, 16); central.writeUInt32LE(data.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(nameBytes.length, 28); central.writeUInt32LE(offset, 42);
    locals.push(local, nameBytes, data); centrals.push(central, nameBytes);
    offset += 30 + nameBytes.length + data.length;
  }
  const directory = Buffer.concat(centrals), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(Object.keys(files).length, 8); end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directory, end]);
}

/** Import layouts mirror what a converter writes; export packs the snapshot's content into an Office zip. */
export function fakeConverter(overrides = {}) {
  const calls = { import: 0, export: 0, exports: [] };
  const converter = {
    calls,
    capabilities: () => Object.fromEntries(['docx', 'xlsx', 'pptx'].map(f => [f, { import: true, preview: true, edit: true, save: true, export: true }])),
    async import({ format, output }, sgn) {
      calls.import++;
      if (overrides.import) return overrides.import({ format, output }, sgn);
      if (format === 'docx') {
        await writeFile(join(output, 'content.json'), '{"doc":1}');
        await mkdir(join(output, 'image')); await writeFile(join(output, 'image', 'a.png'), 'png');
        return { format, content: 'content.json', assets: ['image'] };
      }
      if (format === 'xlsx') {
        await mkdir(join(output, 'workbook.modoc', 'media'), { recursive: true });
        await writeFile(join(output, 'workbook.modoc', 'content'), 'modoc-1');
        await writeFile(join(output, 'workbook.modoc', 'media', 'a.png'), 'png');
        return { format, content: 'workbook.modoc/content', assets: ['workbook.modoc/media'] };
      }
      await writeFile(join(output, 'content.json'), '{"slides":[]}');
      await mkdir(join(output, 'media')); await writeFile(join(output, 'media', 'a.png'), 'png');
      return { format, content: 'content.json', assets: ['media'] };
    },
    async export({ format, snapshot, content, output }, sgn) {
      calls.export++;
      if (overrides.export) return overrides.export({ format, snapshot, content, output }, sgn);
      const text = await readFile(join(snapshot, content), 'utf8');
      calls.exports.push(text);
      await writeFile(output, zip({ '[Content_Types].xml': '<Types/>', [MAIN[format]]: text, ...(overrides.extra ?? {}) }));
      return { output };
    },
  };
  return converter;
}

export async function fixture(t, options = {}) {
  const root = await mkdtemp(join(tmpdir(), 'suite-core-test-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const runtimeRoot = join(root, 'runtime'), tempRoot = join(root, 'temp'), dataRoot = join(root, 'data');
  await mkdir(runtimeRoot); await writeFile(join(runtimeRoot, 'preview.html'), '<html>preview</html>');
  const state = { version: 1, data: Buffer.from('original'), allowed: true, denyActions: new Set(), writes: [], authorizations: [], outside: new Set(), denied: new Set() };
  const source = {
    async stat(s, path, sgn) {
      sgn.throwIfAborted();
      if (s.sessionId !== scope.sessionId) throw new SuiteError('forbidden', 'denied');
      return { absolutePath: '/authorized/' + path, version: `v${state.version}`, bytes: state.data.length };
    },
    async readBytes(s, path, opts, sgn) { const info = await this.stat(s, path, sgn); return { ...info, data: opts.range ? state.data.subarray(0, 1) : state.data, eof: true }; },
  };
  const writer = {
    async replace(s, path, data, { expectedVersion, approved }, sgn) {
      sgn.throwIfAborted();
      if (state.writeError) throw state.writeError;
      if (state.outside.has(path) && !approved) throw new SuiteError('sandbox-denied', 'The target is outside the writable roots.');
      if (expectedVersion !== `v${state.version}`) throw new SuiteError('source-changed', 'changed');
      state.version++; state.data = Buffer.from(data); state.writes.push({ path, data });
      return { version: `v${state.version}`, absolutePath: '/authorized/' + path };
    },
    async writeAccess(s, path) { return state.denied?.has(path) ? 'denied' : state.outside.has(path) ? 'outside' : 'allowed'; },
    async create(s, path, data) {
      if (state.writes.some(w => w.path === path)) throw new SuiteError('exists', 'exists');
      state.writes.push({ path, data, created: true });
      return { absolutePath: '/authorized/' + path, version: 'new' };
    },
  };
  const authorizer = {
    async authorize(peer, s, sgn, args) {
      state.authorizations.push(args.action);
      if (!state.allowed || state.denyActions.has(args.action)) throw new SuiteError('forbidden', 'denied');
    },
  };
  const converter = options.converter ?? fakeConverter();
  const service = new SuiteService({ source, writer: options.noWriter ? undefined : writer, authorizer, converter, runtimeRoot, tempRoot, dataRoot, limits: options.limits, now: options.now, hostCapabilities: options.hostCapabilities });
  t.after(() => service.dispose());
  return { service, state, root, tempRoot, dataRoot, converter, source, writer };
}

export const taskDirs = async dir => (await import('node:fs/promises')).readdir(dir).then(n => n.filter(x => x.startsWith('officesdk-')), () => []);
