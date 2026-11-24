import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync, spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { PassThrough } from 'node:stream';
import { fileURLToPath } from 'node:url';
import { FrameDecoder, KitProcess, MAX_FRAME, UNKNOWN, createSideStream, encodeFrame, lockWorkRoot } from '../src/index.mjs';

const FAKE = fileURLToPath(new URL('./fake-kit.mjs', import.meta.url));
const work = t => { const dir = mkdtempSync(path.join(tmpdir(), 'kit-host-')); t.after(() => rmSync(dir, { recursive: true, force: true })); return dir; };
const fakeKit = (t, env = {}) => {
  const kit = new KitProcess({ binary: process.execPath, work: work(t), restartMs: [50], killGraceMs: 300, spawn: (bin, args, options) => spawnFake(args, options, env) });
  t.after(() => kit.stop());
  return kit;
};
const spawnFake = (args, options, env) => spawn(process.execPath, [FAKE, ...args], { ...options, env: { ...options.env, ...env } });
const sideBody = (stream, text) => { const bytes = Buffer.from(text); const head = Buffer.alloc(16); head.writeBigUInt64BE(BigInt(stream)); head.writeBigUInt64BE(BigInt(bytes.length), 8); return Buffer.concat([head, bytes]); };
const read = async readable => { const chunks = []; for await (const chunk of readable) chunks.push(chunk); return Buffer.concat(chunks).toString(); };

test('frames: 4-byte big-endian length plus JSON, split across chunks; oversize and non-JSON break the channel', () => {
  const bytes = encodeFrame({ id: '1', op: 'open', path: '/a.pptx' });
  assert.equal(bytes.readUInt32BE(0), bytes.length - 4);
  const decoder = new FrameDecoder();
  assert.deepEqual(decoder.push(bytes.subarray(0, 3)), []);
  assert.deepEqual(decoder.push(Buffer.concat([bytes.subarray(3), bytes])), [{ id: '1', op: 'open', path: '/a.pptx' }, { id: '1', op: 'open', path: '/a.pptx' }]);
  const huge = Buffer.alloc(4); huge.writeUInt32BE(MAX_FRAME + 1);
  assert.throws(() => new FrameDecoder().push(huge), /exceeds 64 MiB/);
  const junk = Buffer.from('xx'); const head = Buffer.alloc(4); head.writeUInt32BE(2);
  assert.throws(() => new FrameDecoder().push(Buffer.concat([head, junk])), /not JSON/);
  assert.throws(() => encodeFrame({ body: 'x'.repeat(MAX_FRAME) }), /exceeds 64 MiB/);
});

test('side stream: bytes that arrive before the JSON wait for their stream; split headers and bodies reassemble', async () => {
  const duplex = new PassThrough();
  const side = createSideStream(duplex);
  const early = sideBody(5, 'early bytes');
  duplex.write(early.subarray(0, 10)); duplex.write(early.subarray(10));
  duplex.write(sideBody(6, ''));
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(await read(side.openBody('5', 11)), 'early bytes');
  assert.equal(await read(side.openBody('6', 0)), '');
  const late = side.openBody('8', 4);
  duplex.write(sideBody(8, 'late'));
  assert.equal(await read(late), 'late');
  const short = side.openBody('9', 10);
  duplex.write(sideBody(9, 'short'));
  await assert.rejects(read(short), /length mismatch/);
});

test('ready: protocol "5" is required; nothing is sent to a Kit that speaks another version', async t => {
  const kit = fakeKit(t, { FAKE_KIT_PROTOCOL: '6' });
  await assert.rejects(kit.start(), /protocol 6, expected 5/);
  assert.equal(kit.client, null);
});

test('requests pair by id, explicit failures stay explicit, at most 8 commands are in flight', async t => {
  const kit = fakeKit(t);
  const client = await kit.start();
  const [opened, failed] = await Promise.all([client.request('open', { path: 'a.pptx' }), client.request('fail', { documentId: 'x' })]);
  assert.deepEqual([opened.ok, opened.documentId, opened.nextChangeSetRevision], [true, 'doc-a.pptx', 0]);
  assert.deepEqual([failed.ok, failed.code, failed.unknown], [false, 'file-changed', undefined]);
  const held = await Promise.all(Array.from({ length: 12 }, () => client.request('hold')));
  assert.ok(held.every(frame => frame.ok));
  assert.equal((await client.request('stats')).maxOutstanding, 8);
  const cancelled = new AbortController(); cancelled.abort();
  assert.equal((await client.request('open', { path: 'b' }, { signal: cancelled.signal })).code, 'cancelled');
});

test('media bytes: read before the JSON names the stream, and a write is framed on fd 3', async t => {
  const kit = fakeKit(t);
  const client = await kit.start();
  const response = await client.request('media-read', { text: 'png bytes' });
  assert.equal(await read(client.side.openBody(response.stream, response.length)), 'png bytes');
  const begin = await client.request('media-write');
  await client.side.writeBody(begin.stream, Buffer.from('uploaded'));
  assert.equal((await client.request('media-written')).received, 'uploaded');
});

for (const [op, reason] of [['no-id', 'a frame without id'], ['garbage', 'a frame that is not JSON'], ['crash', 'the process exiting']]) {
  test(`${reason} makes in-flight commands unknown, reports the generation down and restarts Kit`, async t => {
    const kit = fakeKit(t);
    const client = await kit.start();
    const downs = [], readies = [];
    kit.on('down', event => downs.push(event.generation));
    kit.on('ready', event => readies.push(event.generation));
    const inFlight = client.request('never');
    await new Promise(resolve => setTimeout(resolve, 30));
    await client.request(op);
    assert.deepEqual(await inFlight, UNKNOWN);
    assert.equal((await client.request('open', { path: 'a' })).code, 'kit-unavailable', 'the broken client never sends again');
    for (let i = 0; i < 100 && !readies.length; i++) await new Promise(resolve => setTimeout(resolve, 20));
    assert.deepEqual([downs, readies], [[1], [2]]);
    assert.equal((await kit.client.request('open', { path: 'a' })).ok, true, 'the new generation works');
  });
}

test('Windows side stream: two one-way pipes — the Host writes media on index 3 and reads Kit\'s bytes on index 4', async t => {
  const launches = [];
  const kit = new KitProcess({ binary: 'offline-kit-runtime.exe', work: work(t), restartMs: [50], killGraceMs: 300, platform: 'win32',
    spawn: (bin, args, options) => { launches.push(options); return spawnFake(args, options, { FAKE_KIT_SPLIT: '1' }); } });
  t.after(() => kit.stop());
  const client = await kit.start();
  assert.deepEqual(launches[0].stdio, ['pipe', 'pipe', 'pipe', 'pipe', 'pipe']);
  assert.equal(launches[0].windowsHide, true, 'no console window flashes up');
  // The fake Kit writes only on 4 and reads only on 3: a Host using one duplex pipe would never see these bytes.
  const response = await client.request('media-read', { text: 'png bytes' });
  assert.equal(await read(client.side.openBody(response.stream, response.length)), 'png bytes');
  const begin = await client.request('media-write');
  await client.side.writeBody(begin.stream, Buffer.from('uploaded'));
  assert.equal((await client.request('media-written')).received, 'uploaded');
  // Unix keeps its single duplex index 3.
  const unix = [];
  const mac = new KitProcess({ binary: 'offline-kit-runtime', work: work(t), killGraceMs: 300, platform: 'darwin', spawn: (bin, args, options) => { unix.push(options); return spawnFake(args, options); } });
  t.after(() => mac.stop());
  await mac.start();
  assert.deepEqual([unix[0].stdio, unix[0].windowsHide], [['pipe', 'pipe', 'pipe', 'pipe'], true]);
});

test('a split side stream reads one pipe and writes the other; an error on either breaks it', async () => {
  const input = new PassThrough(), output = new PassThrough();
  const side = createSideStream(input, output);
  input.write(sideBody(5, 'from kit'));
  assert.equal(await read(side.openBody('5', 8)), 'from kit');
  const written = [];
  output.on('data', chunk => written.push(chunk));
  await side.writeBody('9', Buffer.from('to kit'));
  assert.deepEqual(Buffer.concat(written), sideBody(9, 'to kit'));
  output.destroy(new Error('pipe gone'));
  await new Promise(resolve => setImmediate(resolve));
  assert.match(side.failed?.message ?? '', /pipe gone/);
  await assert.rejects(side.writeBody('10', Buffer.from('x')), /pipe gone/);
});

test('stopping on Windows: stdin closes first, Kit gets the grace period to leave, only then kill() — never a POSIX signal', async t => {
  const stopped = async env => {
    const kills = [];
    const kit = new KitProcess({ binary: 'offline-kit-runtime.exe', work: work(t), killGraceMs: 300, platform: 'win32',
      spawn: (bin, args, options) => { const child = spawnFake(args, options, { FAKE_KIT_SPLIT: '1', ...env }); const kill = child.kill.bind(child); child.kill = signal => { kills.push([signal, Date.now()]); return kill(signal); }; return child; } });
    t.after(() => kit.stop());
    await kit.start();
    const { pid } = kit;
    await kit.stop();
    assert.throws(() => process.kill(pid, 0), /ESRCH/, 'no process is left behind');
    const closed = Number(readFileSync(path.join(kit.work, 'fake-kit-stdin.log'), 'utf8').trim());
    return { kills, closed };
  };
  // A Kit that leaves when stdin closes is never killed.
  assert.deepEqual((await stopped({})).kills, []);
  // One that stays is killed with kill() — no SIGTERM first — after the grace period.
  const { kills, closed } = await stopped({ FAKE_KIT_IGNORE_STDIN_END: '1' });
  assert.deepEqual(kills.map(([signal]) => signal), [undefined]);
  assert.ok(kills[0][1] - closed >= 250, `killed ${kills[0][1] - closed}ms after stdin closed`);
});

test('suspend stops Kit and keeps it stopped for an update: no restart, requests fail as kit-suspended, resume starts again', async t => {
  const kit = fakeKit(t);
  await kit.start();
  const { pid } = kit;
  await kit.suspend();
  assert.equal(kit.suspended, true);
  assert.throws(() => process.kill(pid, 0), /ESRCH/);
  assert.equal(kit.client, null);
  assert.equal(existsSync(path.join(kit.work, '.officesdk-kit.lock')), false, 'the work root is free for whatever runs next');
  await assert.rejects(kit.ready(), error => error.code === 'kit-suspended');
  await new Promise(resolve => setTimeout(resolve, 150));
  assert.equal(readFileSync(path.join(kit.work, 'fake-kit.log'), 'utf8').trim().split('\n').length, 1, 'no restart timer ran');
  kit.resume();
  assert.equal(kit.suspended, false);
  const client = await kit.ready();
  assert.equal((await client.request('open', { path: 'a' })).ok, true);
  assert.notEqual(kit.pid, pid);
});

test('the work root admits one Kit; a lock left by a dead process is taken over; stop leaves no process', async t => {
  const dir = work(t);
  const release = lockWorkRoot(dir);
  release();
  const dead = execFileSync(process.execPath, ['-e', 'console.log(process.pid)'], { encoding: 'utf8' }).trim();
  writeFileSync(path.join(dir, '.officesdk-kit.lock'), dead);
  lockWorkRoot(dir)();
  writeFileSync(path.join(dir, '.officesdk-kit.lock'), String(process.ppid));
  assert.throws(() => lockWorkRoot(dir), /in use by process/);
  rmSync(path.join(dir, '.officesdk-kit.lock'));

  const kit = fakeKit(t, { FAKE_KIT_IGNORE_TERM: '1' });
  await kit.start();
  const pid = Number(readFileSync(path.join(kit.work, 'fake-kit.log'), 'utf8').trim().split(' ').at(-1));
  await kit.stop();
  assert.throws(() => process.kill(pid, 0), /ESRCH/, 'a Kit that ignores SIGTERM is killed after the grace period');
  assert.equal(existsSync(path.join(kit.work, '.officesdk-kit.lock')), false);
  assert.equal(kit.client, null);
});

test('a restart timer and a request starting Kit at once share one process; stop leaves none behind', async t => {
  const kit = fakeKit(t);
  const client = await kit.start();
  const pids = new Set();
  kit.on('ready', () => pids.add(kit.pid));
  await client.request('crash');
  // A caller asks for Kit while the restart timer is also due.
  const [a, b] = await Promise.all([kit.ready(), new Promise(resolve => setTimeout(() => resolve(kit.ready()), 60))]);
  await new Promise(resolve => setTimeout(resolve, 200));
  assert.equal(a, b, 'one client');
  assert.equal(pids.size, 1, 'one restart, one process');
  const started = readFileSync(path.join(kit.work, 'fake-kit.log'), 'utf8').trim().split('\n').length;
  assert.equal(started, 2, 'the original Kit and exactly one restart');
  const { pid } = kit;
  await kit.stop();
  assert.throws(() => process.kill(pid, 0), /ESRCH/);
});

test('two holders in one process cannot share a work root; a Kit that never starts is not respawned forever', async t => {
  const dir = work(t);
  const release = lockWorkRoot(dir);
  assert.throws(() => lockWorkRoot(dir), /in use by process/, 'a second instance in the same process is refused');
  release();
  lockWorkRoot(dir)();

  const kit = new KitProcess({ binary: process.execPath, work: work(t), restartMs: [10], maxRestarts: 2, killGraceMs: 300,
    spawn: (bin, args, options) => spawn(process.execPath, [FAKE, ...args], { ...options, env: { ...options.env, FAKE_KIT_PROTOCOL: '6' } }) });
  t.after(() => kit.stop());
  let gaveUp = false;
  kit.on('gave-up', () => { gaveUp = true; });
  await assert.rejects(kit.start(), /protocol 6/);
  for (let i = 0; i < 100 && !gaveUp; i++) await new Promise(resolve => setTimeout(resolve, 20));
  const starts = readFileSync(path.join(kit.work, 'fake-kit.log'), 'utf8').trim().split('\n').length;
  await new Promise(resolve => setTimeout(resolve, 200));
  assert.equal(gaveUp, true);
  assert.equal(readFileSync(path.join(kit.work, 'fake-kit.log'), 'utf8').trim().split('\n').length, starts, 'no more launches after giving up');
  assert.ok(starts <= 3, `${starts} launches`);
});

test('the work root lock is taken once and released; a live holder refuses another', () => {
  const dir = mkdtempSync(path.join(tmpdir(), 'kit-lock-'));
  const release = lockWorkRoot(dir);
  assert.equal(existsSync(path.join(dir, '.officesdk-kit.lock')), true);
  release();
  assert.equal(existsSync(path.join(dir, '.officesdk-kit.lock')), false);
  writeFileSync(path.join(dir, '.officesdk-kit.lock'), `${process.ppid}:other`);
  assert.throws(() => lockWorkRoot(dir), /in use by process/);
  rmSync(dir, { recursive: true, force: true });
});
