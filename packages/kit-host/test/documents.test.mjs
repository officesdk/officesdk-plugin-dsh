import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { slashed } from '../src/documents.mjs';
import { code, guest, owner, setup } from './helpers.mjs';

test('open: the first edit view is the writer; another tab or conversation gets a read-only view; content passes through unchanged', async t => {
  const { docs, scope, workspace } = setup(t);
  const first = await docs.open(scope, 'deck.pptx', owner());
  assert.deepEqual([first.mode, first.suite, first.content, first.nextChangeSetRevision, first.unsaved], ['edit', 'presentation', '{"slides":[]}', 0, false]);
  const second = await docs.open({ ...scope, sessionId: 's2' }, 'deck.pptx', owner('tab-2'));
  assert.equal(second.mode, 'view');
  await assert.rejects(docs.change(second.id, owner('tab-2'), { seq: 0, format: 'mop-v1', body: 'x' }), code('read-only'));
  await assert.rejects(docs.open(scope, 'locked.pptx', owner()), code('password-protected'));
  await docs.close(first.id, owner());
  const third = await docs.open(scope, 'deck.pptx', owner('tab-3'));
  assert.equal(third.mode, 'edit', 'a closed writer frees the document for the next editor');
  assert.equal(docs.views.get(third.id).realPath, path.join(workspace, 'deck.pptx'));
});

test('the native preparation view reads only its bundled template and never acquires a writer', async t => {
  const {docs,scope,workspace} = setup(t);
  const bundledPreview = path.join(workspace,'deck.pptx');
  const view = await docs.open(scope,'Untitled.pptx',owner(),{mode:'edit',bundledPreview});
  assert.equal(view.mode,'view');
  assert.equal(docs.views.get(view.id).preparingPresentation,true);
  assert.equal(existsSync(path.join(workspace,'Untitled.pptx')),false);
  await assert.rejects(docs.change(view.id,owner(),{seq:0,format:'mop-v1',body:'x'}),code('read-only'));
  await docs.close(view.id,owner());
});

test('authorization comes first: refused requests never reach Kit', async t => {
  const { docs, scope, commands } = setup(t);
  await assert.rejects(docs.open({ ...scope, sessionId: 'gone' }, 'deck.pptx', owner()), code('forbidden'));
  await assert.rejects(docs.open(scope, '../outside/far.pptx', owner()), code('forbidden'));
  await assert.rejects(docs.open(scope, 'deck.pptx', guest(), { mode: 'edit' }), code('forbidden'));
  assert.deepEqual(commands(), [], 'no command was sent');
  const view = await docs.open(scope, 'deck.pptx', guest(), { mode: 'view' });
  assert.equal(view.mode, 'view');
  const readOnly = await docs.open({ ...scope, sessionId: 's-ro', readOnly: true }, 'changed.pptx', owner('tab-ro'));
  await docs.change(readOnly.id, owner('tab-ro'), { seq: 0, format: 'mop-v1', body: 'edit' });
  await assert.rejects(docs.save(readOnly.id, owner('tab-ro')), code('sandbox-denied'));
  assert.equal(commands().filter(command => command.op === 'save').length, 0, 'a read-only Session never asks Kit to save');
});

test('save always names the authorized real path; a link to a file outside the workspace needs this approval', async t => {
  const { docs, scope, outside, workspace, commands } = setup(t);
  const deck = await docs.open(scope, 'deck.pptx', owner());
  await docs.change(deck.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  const saved = await docs.save(deck.id, owner());
  assert.deepEqual([saved.saved, saved.path, typeof saved.version], [true, path.join(workspace, 'deck.pptx'), 'string']);
  const linked = await docs.open(scope, 'link.pptx', owner('tab-2'));
  await docs.change(linked.id, owner('tab-2'), { seq: 0, format: 'mop-v1', body: 'edit' });
  await assert.rejects(docs.save(linked.id, owner('tab-2')), code('sandbox-denied', { outside: true }));
  await docs.save(linked.id, owner('tab-2'), { approved: true });
  const saves = commands().filter(command => command.op === 'save');
  assert.deepEqual(saves.map(command => command.target), [path.join(workspace, 'deck.pptx'), path.join(outside, 'far.pptx')]);
});

test('changes: numbered without gaps; a gap or a refused change stops the view until the page rebuilds', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  const results = await Promise.all([0, 1, 2].map(seq => docs.change(view.id, owner(), { seq, format: 'mop-v1', body: `change ${seq}` })));
  assert.deepEqual(results.map(result => result.accepted), [0, 1, 2]);
  assert.deepEqual(commands().filter(command => command.op === 'change').map(command => command.body), ['change 0', 'change 1', 'change 2']);
  await assert.rejects(docs.change(view.id, owner(), { seq: 5, format: 'mop-v1', body: 'late' }), code('edit-conflict', { rebuild: true }));
  await assert.rejects(docs.change(view.id, owner(), { seq: 3, format: 'mop-v1', body: 'next' }), code('edit-conflict', { rebuild: true }), 'the view stays stopped');
  await assert.rejects(docs.save(view.id, owner()), code('edit-conflict'));

  await docs.close(view.id, owner(), { autosave: false });
  const rebuilt = await docs.open(scope, 'deck.pptx', owner());
  await assert.rejects(docs.change(rebuilt.id, owner(), { seq: 0, format: 'mop-v1', body: 'reject' }), code('invalid-body', { rebuild: true }));
  await assert.rejects(docs.change(rebuilt.id, owner(), { seq: 1, format: 'mop-v1', body: 'more' }), code('edit-conflict', { rebuild: true }));
});

test('save waits for queued changes; AI edits and unplaced pictures hold it off; a changed file needs overwrite', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'changed.pptx', owner());
  void docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'one' });
  void docs.change(view.id, owner(), { seq: 1, format: 'mop-v1', body: 'two' });
  const done = docs.aiBegin(view.id, owner());
  await assert.rejects(docs.save(view.id, owner()), code('busy'));
  done();
  const { reference } = await docs.writeMedia(view.id, owner(), { bytes: Buffer.from('png'), extension: 'png', kind: 'media' });
  await assert.rejects(docs.save(view.id, owner()), code('busy'));
  await docs.change(view.id, owner(), { seq: 2, format: 'mop-v1', body: `{"src":"${reference}"}` });
  await assert.rejects(docs.save(view.id, owner()), code('source-changed', { overwrite: true }));
  const ops = commands().map(command => command.op);
  assert.ok(ops.lastIndexOf('change') < ops.lastIndexOf('save'), 'every queued change reached Kit before the save');
  await docs.save(view.id, owner(), { overwrite: true });
  assert.equal(commands().at(-1).overwrite, true);
  assert.equal((await new Response((await docs.readMedia(view.id, owner(), { reference })).body).text()), `media ${reference}`);
});

test('a Kit restart makes every view stale; closing autosaves only for the owner who may write', async t => {
  const { docs, kit, scope, commands } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  const before = kit.generation;
  (await kit.ready()).request('crash');
  for (let i = 0; i < 100 && kit.generation === before; i++) await new Promise(resolve => setTimeout(resolve, 20));
  assert.throws(() => docs.view(view.id, owner()), code('edit-conflict', { rebuild: true }));
  const reopened = await docs.open(scope, 'deck.pptx', owner());
  assert.equal(reopened.mode, 'edit', 'the stale writer does not block the reopened view');
  await docs.change(reopened.id, owner(), { seq: 0, format: 'mop-v1', body: 'again' });
  const closed = await docs.close(reopened.id, owner());
  assert.deepEqual(closed, { closed: true, saved: true });
  const viewer = await docs.open(scope, 'deck.pptx', guest(), { mode: 'view' });
  await docs.close(viewer.id, guest());
  assert.equal(commands().filter(command => command.op === 'save').length, 1, 'only the owner\'s close saved');
  assert.equal(commands().filter(command => command.op === 'discard').length, 0, 'closing never discards');
});

test('export: downloads go to a Host-private file; a Session copy must be new and writable', async t => {
  const { docs, scope, workspace, base } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  const download = await docs.export(view.id, owner(), { target: 'download' });
  assert.ok(download.file.startsWith(path.join(base, 'exports')));
  assert.equal(readFileSync(download.file, 'utf8'), 'pptx bytes');
  await assert.rejects(docs.export(view.id, owner(), { target: { path: 'changed.pptx' } }), code('exists'));
  await assert.rejects(docs.export(view.id, owner(), { target: { path: 'copy.docx' } }), code('invalid-path'));
  const copy = await docs.export(view.id, owner(), { target: { path: 'copy.pptx' } });
  assert.equal(copy.file, path.join(workspace, 'copy.pptx'));
});

test('a disk change is the user\'s choice: keep editing sends nothing, use the disk imports and reopens', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'disk.pptx', owner());
  assert.equal(view.diskChanged, true);
  assert.deepEqual(await docs.resolveDiskChange(view.id, owner(), { useDisk: false }), { kept: true });
  assert.equal(commands().filter(command => command.op === 'import').length, 0);
  await docs.resolveDiskChange(view.id, owner(), { useDisk: true });
  assert.equal(commands().at(-1).op, 'import');
  assert.throws(() => docs.view(view.id, owner()), code('edit-conflict', { rebuild: true }), 'the page reopens with the imported content');
});

test('a change Kit takes while a save runs keeps the document unsaved, so closing still saves it', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'slow.pptx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'before' });
  const saving = docs.save(view.id, owner());
  await new Promise(resolve => setTimeout(resolve, 40));
  await docs.change(view.id, owner(), { seq: 1, format: 'mop-v1', body: 'during' });
  await saving;
  assert.equal(docs.views.get(view.id).unsaved, true, 'the edit typed during the save is not saved');
  await docs.close(view.id, owner());
  assert.equal(commands().filter(command => command.op === 'save').length, 2, 'closing saves it');
});

test('pictures: held from the first byte, for every view of the document; only the editing tab cancels', async t => {
  const { docs, scope } = setup(t);
  const writer = await docs.open(scope, 'deck.pptx', owner());
  const reader = await docs.open({ ...scope, sessionId: 's2' }, 'deck.pptx', owner('tab-2'));
  const storing = docs.writeMedia(writer.id, owner(), { bytes: Buffer.from('png'), extension: 'png', kind: 'media' });
  await new Promise(resolve => setImmediate(resolve));
  await assert.rejects(docs.save(writer.id, owner()), code('busy'), 'held while the upload runs');
  const { reference } = await storing;
  await assert.rejects(docs.export(reader.id, owner('tab-2'), { target: 'download' }), code('busy'), 'another view of the document waits too');
  assert.throws(() => docs.cancelMedia(writer.id, owner('tab-2'), reference), code('lease-unavailable'), 'another tab cannot release the hold');
  assert.throws(() => docs.cancelMedia(reader.id, owner('tab-2'), reference), code('read-only'));
  docs.cancelMedia(writer.id, owner(), reference);
  await docs.export(reader.id, owner('tab-2'), { target: 'download' });
});

test('closing autosaves only while the Session still authorizes the save', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  docs.views.get(view.id).scope = { ...scope, sessionId: 'gone' };
  assert.deepEqual(await docs.close(view.id, owner()), { closed: true, saved: false });
  assert.equal(commands().filter(command => command.op === 'save').length, 0);
});

test('DOCX and XLSX open as their own suites; a record Kit opened as another suite gets no view and is never discarded', async t => {
  const { docs, scope, commands } = setup(t);
  const doc = await docs.open(scope, 'report.docx', owner());
  const book = await docs.open(scope, 'book.xlsx', owner());
  assert.deepEqual([doc.suite, doc.mode, book.suite, book.mode], ['document', 'edit', 'sheet', 'edit']);
  const views = docs.views.size;
  await assert.rejects(docs.open(scope, 'mismatch.docx', owner()), code('unsupported-format'));
  await assert.rejects(docs.open(scope, 'mismatch.xlsx', owner()), code('unsupported-format'));
  assert.equal(docs.views.size, views);
  assert.equal(commands().filter(command => command.op === 'discard').length, 0);
});

test('each suite takes only its own change format; a wrong one never reaches Kit', async t => {
  const { docs, scope, commands } = setup(t);
  const doc = await docs.open(scope, 'report.docx', owner());
  const book = await docs.open(scope, 'book.xlsx', owner());
  await assert.rejects(docs.change(doc.id, owner(), { seq: 0, format: 'mop-v1', body: 'x' }), code('invalid-body'));
  await assert.rejects(docs.change(book.id, owner(), { seq: 0, body: 'x' }), code('invalid-body'));
  assert.equal(commands().filter(command => command.op === 'change').length, 0);
  await docs.change(doc.id, owner(), { seq: 0, format: 'mow-v1', body: 'word' });
  await docs.change(book.id, owner(), { seq: 0, format: 'modoc-v1', body: 'cell' });
  assert.deepEqual(commands().filter(command => command.op === 'change').map(command => command.body), ['word', 'cell']);
});

test('media kinds per suite: document image|embedding, sheet none, presentation media|embeddings — refused in the Host', async t => {
  const { docs, scope, commands } = setup(t);
  const doc = await docs.open(scope, 'report.docx', owner());
  const book = await docs.open(scope, 'book.xlsx', owner());
  const deck = await docs.open(scope, 'deck.pptx', owner());
  const write = (view, kind) => docs.writeMedia(view.id, owner(), { bytes: Buffer.from('png'), extension: 'png', kind });
  for (const [view, kind] of [[doc, undefined], [doc, 'media'], [book, 'image'], [book, ''], [deck, 'image'], [deck, undefined]]) await assert.rejects(write(view, kind), code('invalid-body'), `${view.suite} ${kind}`);
  assert.equal(commands().filter(command => command.op === 'begin-media-write').length, 0);
  await write(doc, 'image'); await write(doc, 'embedding'); await write(book, undefined); await write(book, null); await write(deck, 'embeddings');
  assert.deepEqual(commands().filter(command => command.op === 'begin-media-write').map(command => command.kind ?? null), ['image', 'embedding', null, null, 'embeddings']);
});

test('copies keep the document\'s own extension; downloads carry its file type; restriction failures keep the edits', async t => {
  const { docs, scope, workspace } = setup(t);
  const book = await docs.open(scope, 'book.xlsx', owner());
  await assert.rejects(docs.export(book.id, owner(), { target: { path: 'copy.pptx' } }), code('invalid-path'));
  assert.equal((await docs.export(book.id, owner(), { target: { path: 'copy.xlsx' } })).file, path.join(workspace, 'copy.xlsx'));
  const download = await docs.exportDownload(book.id, owner());
  const taken = docs.takeDownload(book.id, owner(), download.token);
  assert.deepEqual([path.extname(taken.file), taken.name, taken.contentType], ['.xlsx', 'book.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']);
  const doc = await docs.open(scope, 'report.docx', owner());
  assert.equal(docs.takeDownload(doc.id, owner(), (await docs.exportDownload(doc.id, owner())).token).contentType, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  const guarded = await docs.open(scope, 'protected.docx', owner());
  await docs.change(guarded.id, owner(), { seq: 0, format: 'mow-v1', body: 'edit' });
  await assert.rejects(docs.save(guarded.id, owner()), code('unsupported-operation'));
  assert.equal(docs.views.get(guarded.id).unsaved, true, 'the edits stay in the record');
});

test('the file version travels with the view: as opened, then as each save wrote it', async t => {
  const { docs, scope, workspace } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  assert.match(view.version, /^v/);
  await docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  await new Promise(resolve => setTimeout(resolve, 20));
  writeFileSync(path.join(workspace, 'deck.pptx'), 'pptx written by Kit');
  const saved = await docs.save(view.id, owner());
  assert.equal(saved.version, docs.views.get(view.id).version);
  assert.notEqual(saved.version, view.version, 'the tab learns the version its own save produced');
  assert.deepEqual(await docs.checkDisk(view.id, owner()), { version: saved.version }, 'its own write is no change');
  await new Promise(resolve => setTimeout(resolve, 20));
  writeFileSync(path.join(workspace, 'deck.pptx'), 'pptx written by another program');
  await assert.rejects(docs.checkDisk(view.id, owner()), code('source-changed'));
});

test('abandon: a view whose editor could not mount closes without saving or discarding; only its own tab may abandon it', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  assert.throws(() => docs.abandon(view.id, owner('tab-2')), code('lease-unavailable'));
  assert.deepEqual(docs.abandon(view.id, owner()), { closed: true, unsaved: true });
  assert.equal(docs.views.has(view.id), false);
  assert.deepEqual(commands().filter(command => ['save', 'discard'].includes(command.op)), []);
  const again = await docs.open(scope, 'deck.pptx', owner());
  assert.deepEqual([again.mode, again.unsaved], ['edit', true], 'the edits stay in Kit and the document has a writer again');
});

test('a disk check while a save is writing waits for it: the save\'s own write is not another program\'s', async t => {
  const { docs, scope, workspace } = setup(t);
  const view = await docs.open(scope, 'slow.pptx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  const saving = docs.save(view.id, owner());
  await new Promise(resolve => setTimeout(resolve, 40));
  writeFileSync(path.join(workspace, 'slow.pptx'), 'pptx written by Kit'); // what Kit's save does to the file
  const checked = await docs.checkDisk(view.id, owner());
  const saved = await saving;
  assert.equal(checked.version, saved.version);
});

test('the same tab reopening takes over its own writer view; another tab still gets a read-only view', async t => {
  const { docs, scope, commands } = setup(t);
  const first = await docs.open(scope, 'deck.pptx', owner());
  await docs.change(first.id, owner(), { seq: 0, format: 'mop-v1', body: 'edit' });
  const again = await docs.open(scope, 'deck.pptx', owner());
  assert.equal(again.mode, 'edit', 'its release has not arrived yet, the tab still edits');
  assert.equal(docs.views.has(first.id), false);
  assert.deepEqual(commands().filter(command => ['save', 'discard'].includes(command.op)), [], 'no autosave, no discard: the edits stay in Kit');
  assert.equal(again.unsaved, true);
  const other = await docs.open(scope, 'deck.pptx', owner('tab-2'));
  assert.equal(other.mode, 'view');
  assert.deepEqual(await docs.close(first.id, owner()), { closed: true }, 'the late release of the replaced view is harmless');
});

test('Save As: an existing name is replaced only when asked, through a temporary file; never this document itself; the panel\'s choice allows outside the workspace', async t => {
  const { docs, scope, workspace, outside } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  writeFileSync(path.join(workspace, 'copy.pptx'), 'old copy');
  await assert.rejects(docs.export(view.id, owner(), { target: { path: 'copy.pptx' } }), code('exists'));
  const replaced = await docs.export(view.id, owner(), { target: { path: 'copy.pptx', overwrite: true } });
  assert.equal(replaced.file, path.join(workspace, 'copy.pptx'));
  assert.equal(readFileSync(replaced.file, 'utf8'), 'pptx bytes');
  assert.deepEqual(readdirSync(workspace).filter(name => name.includes('.tmp')), [], 'no temporary file is left behind');
  await assert.rejects(docs.export(view.id, owner(), { target: { path: 'deck.pptx', overwrite: true } }), code('same-file'));
  await assert.rejects(docs.export(view.id, owner(), { target: { path: 'deck.pptx' } }), code('same-file'), 'its own name is refused before asking to replace');
  // Outside the workspace: refused without the owner's choice, written with it.
  const far = path.join(outside, 'far-copy.pptx');
  await assert.rejects(docs.export(view.id, owner(), { target: { path: far } }), code('sandbox-denied', { outside: true }));
  assert.equal((await docs.export(view.id, owner(), { target: { path: far, approved: true } })).file, far);
  // Only the owner's choice counts as approval.
  const guestView = await docs.open(scope, 'changed.pptx', guest(), { mode: 'view' });
  await assert.rejects(docs.export(guestView.id, guest(), { target: { path: path.join(outside, 'guest-copy.pptx'), approved: true } }), error => ['sandbox-denied', 'forbidden'].includes(error.code));
});

test('Save As on Windows: replacing a file another program holds open (EPERM/EBUSY) is target-busy; the old file and no temporary stay', async t => {
  const { docs, scope, workspace } = setup(t);
  const view = await docs.open(scope, 'deck.pptx', owner());
  writeFileSync(path.join(workspace, 'copy.pptx'), 'old copy');
  const held = errno => async () => { throw Object.assign(new Error(`${errno}: operation not permitted, rename`), { code: errno }); };
  docs.platform = 'win32';
  for (const errno of ['EPERM', 'EBUSY']) {
    docs.renameFile = held(errno);
    await assert.rejects(docs.export(view.id, owner(), { target: { path: 'copy.pptx', overwrite: true } }), error => error.code === 'target-busy' && /close it/.test(error.message));
    assert.equal(readFileSync(path.join(workspace, 'copy.pptx'), 'utf8'), 'old copy');
    assert.deepEqual(readdirSync(workspace).filter(name => name.includes('.tmp')), [], 'no temporary file is left behind');
  }
  // Elsewhere the same errno is no program holding the file: it stays a plain failure.
  docs.platform = 'darwin';
  docs.renameFile = held('EPERM');
  await assert.rejects(docs.export(view.id, owner(), { target: { path: 'copy.pptx', overwrite: true } }), code('unavailable'));
});

test('paths handed to the Client are \'/\'-separated on every platform', () => {
  assert.equal(slashed('sub\\deep\\a.docx', '\\'), 'sub/deep/a.docx');
  assert.equal(slashed('', '\\'), '');
  assert.equal(slashed('sub/a\\b.docx', '/'), 'sub/a\\b.docx', 'on POSIX a backslash is part of a name');
});

test('probe: the canonical path and real path of a picked file or folder, with open\'s read rules; browse stays inside the workspace', async t => {
  const { docs, scope, workspace, outside } = setup(t);
  writeFileSync(path.join(workspace, 'real.docx'), Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.from('rest')]));
  writeFileSync(path.join(workspace, 'fake.docx'), 'not a zip');
  writeFileSync(path.join(workspace, 'notes.txt'), 'text');
  // The same file picked by its absolute path maps to the workspace-relative path.
  const absolute = await docs.probe(scope, path.join(workspace, 'real.docx'), owner());
  const relative = await docs.probe(scope, 'real.docx', owner());
  assert.deepEqual([absolute.path, absolute.realPath, absolute.inWorkspace], ['real.docx', relative.realPath, true]);
  await assert.rejects(docs.probe(scope, 'fake.docx', owner()), code('invalid-format'));
  await assert.rejects(docs.probe(scope, 'notes.txt', owner()), code('unsupported-format'));
  await assert.rejects(docs.probe(scope, 'missing.docx', owner()), error => typeof error.code === 'string');
  await assert.rejects(docs.probe({ ...scope, sessionId: 'gone' }, 'real.docx', owner()), code('forbidden'));
  const folder = await docs.probe(scope, outside, owner(), { kind: 'folder' });
  assert.deepEqual([folder.inWorkspace, folder.access], [false, 'outside']);
  mkdirSync(path.join(workspace, 'sub'));
  assert.deepEqual(await docs.probe(scope, 'sub', owner(), { kind: 'folder' }), { realPath: path.join(workspace, 'sub'), path: 'sub', inWorkspace: true, access: 'allowed' });
  mkdirSync(path.join(workspace, 'sub', 'deep'));
  assert.equal((await docs.browse(scope, 'sub/deep', owner())).path, 'sub/deep');
  const listing = await docs.browse(scope, '', owner());
  assert.ok(listing.entries.some(entry => entry.name === 'real.docx' && entry.kind === 'file'));
  assert.ok(!listing.entries.some(entry => entry.name === 'notes.txt'), 'only Office files are listed');
  await assert.rejects(docs.browse(scope, '..', owner()), error => ['forbidden'].includes(error.code));
});

test('inert changes: Kit takes the editor\'s own changes, but nothing turns unsaved, autosaves or survives a close', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'book.xlsx', owner());
  assert.deepEqual(await docs.change(view.id, owner(), { seq: 0, format: 'modoc-v1', body: 'calc-csr', inert: true }), { accepted: 0 });
  assert.equal(docs.views.get(view.id).unsaved, false);
  // The same tab reopening takes over its writer view: Kit now calls the record unsaved, but only housekeeping is in it.
  const again = await docs.open(scope, 'book.xlsx', owner());
  assert.equal(again.unsaved, false);
  assert.deepEqual(await docs.close(again.id, owner()), { closed: true, saved: false });
  const ops = commands().map(command => command.op);
  assert.equal(ops.includes('save'), false, 'never written back');
  assert.equal(ops.at(-1), 'import', 'the disk copy goes back into Kit');
  const later = await docs.open(scope, 'book.xlsx', owner('tab-2'));
  assert.equal(later.unsaved, false, 'Kit itself is clean after the close');
  await docs.close(later.id, owner('tab-2'));
});

test('a real change after inert ones is unsaved and autosaved on close as before; a reader keeps the record', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'book.xlsx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'modoc-v1', body: 'calc-csr', inert: true });
  await docs.change(view.id, owner(), { seq: 1, format: 'modoc-v1', body: 'typed' });
  assert.equal(docs.views.get(view.id).unsaved, true);
  await docs.change(view.id, owner(), { seq: 2, format: 'modoc-v1', body: 'late housekeeping', inert: true });
  assert.equal(docs.views.get(view.id).unsaved, true, 'an inert change never clears a real one');
  assert.deepEqual(await docs.close(view.id, owner()), { closed: true, saved: true });
  assert.deepEqual(commands().filter(command => ['save', 'import'].includes(command.op)).map(command => command.op), ['save']);

  const writer = await docs.open(scope, 'report.docx', owner('tab-3'));
  const reader = await docs.open({ ...scope, sessionId: 's2' }, 'report.docx', owner('tab-4'));
  await docs.change(writer.id, owner('tab-3'), { seq: 0, format: 'mow-v1', body: 'layout', inert: true });
  await docs.close(writer.id, owner('tab-3'));
  assert.equal(commands().filter(command => command.op === 'import').length, 0, 'another open view keeps the record as it is');
  await docs.close(reader.id, owner('tab-4'));
});

test('closing while a real change is still on its way keeps it in Kit: no import after an inert change', async t => {
  const { docs, scope, commands } = setup(t);
  const view = await docs.open(scope, 'book.xlsx', owner());
  await docs.change(view.id, owner(), { seq: 0, format: 'modoc-v1', body: 'calc-csr', inert: true });
  const typed = docs.change(view.id, owner(), { seq: 1, format: 'modoc-v1', body: 'typed' }); // Kit answers after 20 ms
  await docs.close(view.id, owner(), { autosave: false });
  await typed;
  assert.equal(commands().some(command => command.op === 'import'), false, 'the typed edit is not dropped');
  const again = await docs.open(scope, 'book.xlsx', owner());
  assert.equal(again.unsaved, true, 'the edit waits in Kit for the next open');
  await docs.close(again.id, owner(), { autosave: false });
});
