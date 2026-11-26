import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { withFileLock, writeFileAtomic } from '@deepseek-ai/dsh-atomic-write';
import ShimoPreview, { loadStyle, resolveStyle, styleList } from '../src/host/index.mjs';
import { SuiteError } from '../../core/src/index.mjs';
import { StyleMemory } from '../src/host/style-catalog.mjs';
import { StyleSuggestions } from '../src/host/style-suggest.mjs';
import { BUILTIN_TEMPLATES, PICK_CONTEXT_MS, TemplatePicks, pickContext } from '../src/host/template-pick.mjs';
import { withSecondPack } from './style-picker.test.mjs';

const signal = () => new AbortController().signal;

function host(t, { principal = 'owner', authorize, now } = {}) {
  const { root, styles } = withSecondPack(t);
  const ws = join(root, 'ws');
  mkdirSync(ws, { recursive: true });
  const authorized = [];
  const self = Object.create(ShimoPreview.prototype);
  Object.assign(self, {
    ctx: {}, styleRoot: styles, styleCache: undefined,
    styleMemory: new StyleMemory({ file: join(root, 'data', 'style-defaults.json'), atomic: { withFileLock, writeFileAtomic } }),
    styleSuggestions: new StyleSuggestions(), templatePicks: new TemplatePicks(now ? { now } : {}),
    remotePeer: () => ({ principal }),
    suite: { authorize: async (_peer, scope, _signal, action) => { authorized.push([scope.sessionId, action]); await authorize?.(); } },
  });
  return { self, styles, ws, scope: { sessionId: 's1', workspaceRoot: ws }, authorized };
}

test('a style is named by its pack name, display name (zh/en) or alias, ignoring case and spaces', t => {
  const { styles } = host(t);
  assert.equal(resolveStyle(styles, 'second'), 'second');
  assert.equal(resolveStyle(styles, ' 黑底画报 '), 'second');
  assert.equal(resolveStyle(styles, 'dark EDITORIAL'), 'second');
  assert.equal(resolveStyle(styles, '黑底'), 'second');
  assert.equal(resolveStyle(styles, 'Photo-Noir'), 'second');
  for (const bad of ['', '   ', 'nope', '../default', null, 3]) assert.equal(resolveStyle(styles, bad), undefined, String(bad));
  assert.match(styleList(styles), /second（黑底画报）/);
  assert.throws(() => loadStyle(styles, '黑底画报'), /available: .*second（黑底画报）/, 'loadStyle itself still wants the id; callers resolve first');
});

test('rememberStyle accepts a display name and stores the pack name', async t => {
  const { self, scope } = host(t);
  assert.equal((await self.rememberStyle(scope, '黑底画报', signal())).remembered, 'second');
  await assert.rejects(self.rememberStyle(scope, '没有这个', signal()), /available: .*second（黑底画报）/);
});

test('selectTemplate: authorized, per Session, validated by kind; null clears; any peer of the Session may set it', async t => {
  const { self, scope, authorized } = host(t);
  assert.deepEqual(await self.selectTemplate(scope, 'ppt', '黑底画报', signal()), { kind: 'ppt', id: 'second' });
  assert.deepEqual(await self.templatePick(scope, signal()), { kind: 'ppt', id: 'second' });
  assert.deepEqual(await self.templatePick({ ...scope, sessionId: 's2' }, signal()), null, 'another Session has none');
  assert.deepEqual(await self.selectTemplate(scope, 'word', 'gongwen', signal()), { kind: 'word', id: 'gongwen' });
  await assert.rejects(self.selectTemplate(scope, 'word', 'redhead', signal()), /Unknown word template "redhead"; available: meeting-minutes, gongwen\./);
  await assert.rejects(self.selectTemplate(scope, 'excel', 'blank', signal()), /Unknown excel template "blank"; available: none\./);
  await assert.rejects(self.selectTemplate(scope, 'word', 'second', signal()), /Unknown word template "second"; available: meeting-minutes, gongwen/);
  await assert.rejects(self.selectTemplate(scope, 'ppt', 'gongwen', signal()), /Unknown style/);
  await assert.rejects(self.selectTemplate(scope, 'pdf', 'x', signal()), /Unknown template kind/);
  assert.deepEqual(await self.templatePick(scope, signal()), { kind: 'word', id: 'gongwen' }, 'a failed pick leaves the last one');
  assert.equal(await self.selectTemplate(scope, 'excel', null, signal()), null);
  assert.equal(await self.templatePick(scope, signal()), null);
  assert.ok(authorized.length >= 9 && authorized.every(([, action]) => action === 'open'), 'every call is authorized for opening');
  const guest = host(t, { principal: 'share:g' });
  assert.deepEqual(await guest.self.selectTemplate(guest.scope, 'ppt', 'second', signal()), { kind: 'ppt', id: 'second' });
  const denied = host(t, { authorize: async () => { throw new SuiteError('forbidden', 'Access to the document was denied.'); } });
  await assert.rejects(denied.self.selectTemplate(denied.scope, 'ppt', 'second', signal()), /denied/);
});

test('a package that lists its skills: only its templates may be picked or shown, a skill it lacks has no root (261113)', async t => {
  const { self, scope } = host(t);
  // The public build: shimo-pptx and shimo-docx (gongwen) only.
  self.skills = { names: ['shimo-pptx', 'shimo-docx'], effective: name => ({ ready: true, entry: { templates: name === 'shimo-docx' ? ['gongwen'] : undefined } }) };
  assert.deepEqual(self.templateIds(), { word: ['gongwen'], excel: [] });
  assert.equal(self.skillIncluded('shimo-xlsx'), false);
  assert.deepEqual(await self.selectTemplate(scope, 'word', 'gongwen', signal()), { kind: 'word', id: 'gongwen' });
  await assert.rejects(self.selectTemplate(scope, 'word', 'redhead', signal()), /Unknown word template "redhead"; available: gongwen\./);
  await assert.rejects(self.selectTemplate(scope, 'excel', 'blank', signal()), /Unknown excel template "blank"; available: none\./);
  await assert.rejects(self.skillRoot('shimo-xlsx', 's1', 'templates'), error => error.code === 'unsupported-operation' && /does not include shimo-xlsx/.test(error.message));
  const listed = await self.listStyles(scope, signal());
  assert.deepEqual(listed.templates, { word: ['gongwen'], excel: [] });
  // The repository (no packaged catalog): every built-in, no list sent.
  delete self.skills;
  assert.equal(self.templateIds(), undefined);
  assert.equal((await self.listStyles(scope, signal())).templates, undefined);
  assert.deepEqual(await self.selectTemplate(scope, 'word', 'gongwen', signal()), { kind: 'word', id: 'gongwen' });
});

test('the pick reaches every step of the turn after the workspace default, says it counts as naming, and ends by itself', async t => {
  let clock = 1000;
  const { self, scope, ws } = host(t, { now: () => clock });
  await self.rememberStyle(scope, 'default', signal());
  await self.rememberStyle(scope, 'second', signal());
  await self.selectTemplate(scope, 'ppt', 'default', signal());
  const text = self.styleContext(ws, 's1');
  const lines = text.split('\n');
  assert.match(lines[0], /^Workspace default PPT style/);
  assert.match(lines.at(-1), /^Inspiration library pick: .*counts as the user naming it and overrides any workspace default\. PPT style pack "default" .*pass style "default" to presentation_create \(offline: --style default\)\./);
  assert.equal(self.styleContext(ws, 's2'), lines[0], 'other Sessions do not see it');
  clock += PICK_CONTEXT_MS;
  assert.match(self.styleContext(ws, 's1'), /Inspiration library pick/, 'still within the window of the first step that read it');
  clock += 1;
  assert.doesNotMatch(self.styleContext(ws, 's1'), /Inspiration library pick/, 'gone after the window, even if the Client never cleared it');
  assert.equal(await self.templatePick(scope, signal()), null);
});

test('a pick not yet read by any step does not expire; Word and Excel lines name their tool argument', () => {
  let clock = 0;
  const picks = new TemplatePicks({ now: () => clock });
  picks.set('s', { kind: 'word', id: 'gongwen' });
  clock += PICK_CONTEXT_MS * 3;
  assert.deepEqual(picks.get('s'), { kind: 'word', id: 'gongwen' });
  const name = pick => BUILTIN_TEMPLATES[pick.kind].find(item => item.id === pick.id)?.name.zh;
  assert.match(pickContext(picks.forStep('s'), name), /Word template "gongwen" \(公文版式\): pass template "gongwen" to document_create\./);
  assert.match(pickContext({ kind: 'excel', id: 'blank' }, name), /Excel template "blank": create the workbook with spreadsheet_create/);
  assert.equal(pickContext(undefined), '');
});

test('presentation_create resolves a display name before loading; recommendation candidates may be names too', async t => {
  const { self, scope, styles } = host(t);
  // aiSuggestStyles: a candidate given by display name is kept as its pack.
  self.scopeOf = async () => scope;
  const result = await self.aiSuggestStyles('s1', { candidates: [{ style: '黑底画报', reason: 'r' }, { style: 'default', reason: 'd' }] });
  assert.deepEqual(result.candidates.map(item => item.style), ['second', 'default']);
  assert.equal(resolveStyle(styles, 'Dark editorial'), 'second');
});
