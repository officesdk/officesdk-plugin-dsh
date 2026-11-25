import test from 'node:test';
import assert from 'node:assert/strict';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { withFileLock, writeFileAtomic } from '@deepseek-ai/dsh-atomic-write';
import ShimoPreview from '../src/host/index.mjs';
import { SuiteError } from '../../core/src/index.mjs';
import { PREVIEWS_MAX, StyleMemory, styleCard } from '../src/host/style-catalog.mjs';
import { ANSWERED_CONTEXT_MS, SUGGEST_MAX, StyleSuggestions, normalizeCandidates, suggestionContext } from '../src/host/style-suggest.mjs';
import { apply as applyTools } from '../src/tools/index.mjs';

const STYLES = fileURLToPath(new URL('../skills/shimo-pptx/styles/', import.meta.url));
const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const signal = () => new AbortController().signal;

/** The real packs plus test-only `second` and `third` (copies of default with their own meta); never shipped. */
function packs(t, { previews } = {}) {
  const root = realpathSync(mkdtempSync(join(tmpdir(), 'officesdk-suggest-')));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  const styles = join(root, 'styles');
  cpSync(STYLES, styles, { recursive: true });
  for (const [id, zh, en] of [['second', '黑底画报', 'Dark editorial'], ['third', '水墨', 'Ink']]) {
    cpSync(join(styles, 'default'), join(styles, id), { recursive: true });
    const style = JSON.parse(readFileSync(join(styles, id, 'style.json'), 'utf8'));
    style.name = id;
    // Copies of default carry its previews only when a test asks (second, `previews`).
    const { previews: _shipped, ...meta } = style.meta;
    style.meta = { ...meta, id, name: { zh, en }, aliases: [zh], ...(id === 'second' && previews ? { previews } : {}) };
    writeFileSync(join(styles, id, 'style.json'), JSON.stringify(style));
  }
  return { root, styles };
}

test('candidates: known packs only, each once, at most four, reasons clipped to one line', () => {
  const long = 'x'.repeat(300);
  const { candidates, dropped } = normalizeCandidates([
    { style: 'second', reason: ' 黑底大图，\n适合城市  旅行 ' }, { style: 'nope', reason: 'r' }, { style: 'second', reason: 'again' },
    { style: 'default', reason: long }, { style: 'third' }, { style: 'a' }, { style: 'b' }, null, { style: 7 },
  ], ['default', 'second', 'third', 'a', 'b']);
  assert.deepEqual(candidates.map(item => item.style), ['second', 'default', 'third', 'a']);
  assert.equal(candidates.length, SUGGEST_MAX);
  assert.equal(candidates[0].reason, '黑底大图， 适合城市 旅行');
  assert.equal(candidates[1].reason.length, 140);
  assert.match(candidates[1].reason, /…$/);
  assert.equal(candidates[2].reason, '', 'a missing reason is empty, not an error');
  assert.deepEqual(dropped, ['nope']);
  assert.deepEqual(normalizeCandidates(undefined, ['default']), { candidates: [], dropped: [] });
});

test('suggestions: one per Session, replaced by the next; none known fails and lists the packs', () => {
  let id = 0;
  const store = new StyleSuggestions({ newId: () => `g${++id}`, now: () => 1000 });
  const first = store.propose('s1', { candidates: [{ style: 'default', reason: 'a' }], topic: '城市出行' }, ['default', 'second']);
  assert.equal(first.suggestion.id, 'g1');
  assert.equal(first.suggestion.topic, '城市出行');
  store.propose('s1', { candidates: [{ style: 'second', reason: 'b' }] }, ['default', 'second']);
  assert.equal(store.show('s1').id, 'g2', 'the later one replaces the earlier');
  store.propose('s2', { candidates: [{ style: 'default' }] }, ['default']);
  assert.equal(store.show('s2').id, 'g3', 'Sessions are separate');
  assert.throws(() => store.propose('s1', { candidates: [{ style: 'fancy' }] }, ['default', 'second']), /No candidate names a style pack \("fancy" are not style packs\); available: default, second/);
  assert.equal(store.show('s1').id, 'g2', 'a failed recommendation keeps the earlier one');
});

test('suggestions: a step before the strip showed it changes nothing; the first step after it was shown answers it', () => {
  let now = 0;
  const store = new StyleSuggestions({ now: () => now });
  store.propose('s1', { candidates: [{ style: 'default' }, { style: 'second' }] }, ['default', 'second']);
  assert.equal(store.forStep('s1'), undefined, 'the model\'s own next step in the same turn');
  assert.equal(store.show('s1').state, 'shown');
  now = 5;
  const answered = store.forStep('s1');
  assert.equal(answered.state, 'answered', 'the user wrote in a new turn');
  assert.equal(store.show('s1'), undefined, 'the strip no longer shows it');
  now = 5 + ANSWERED_CONTEXT_MS;
  assert.ok(store.forStep('s1'), 'still in context for a while');
  now += 1;
  assert.equal(store.forStep('s1'), undefined, 'then gone');
  assert.equal(store.bySession.size, 0);
});

test('suggestions: dismiss ends only the current one; with no id, whatever is there', () => {
  const store = new StyleSuggestions({ newId: (() => { let n = 0; return () => `g${++n}`; })() });
  store.propose('s1', { candidates: [{ style: 'default' }] }, ['default']);
  assert.equal(store.dismiss('s1', 'old'), false);
  assert.ok(store.show('s1'));
  assert.equal(store.dismiss('s1', 'g1'), true);
  assert.equal(store.show('s1'), undefined);
  store.propose('s1', { candidates: [{ style: 'default' }] }, ['default']);
  assert.equal(store.dismiss('s1'), true);
  assert.equal(store.dismiss('s1'), false);
});

test('context: the answered recommendation is listed by number and name for the reply to be read', () => {
  const text = suggestionContext({ candidates: [{ style: 'second' }, { style: 'default' }] }, id => (id === 'second' ? '黑底画报' : id));
  assert.match(text, /1\) second \(黑底画报\), 2\) default\./);
  assert.match(text, /by number, name or alias/);
  assert.equal(suggestionContext(undefined), '');
});

test('cards: previews named in meta are read as data URIs, at most four, unsafe or missing names left out', t => {
  const names = ['p1.png', 'p2.png', '../x.png', 'p1.png', 'gone.png', 'p3.png', 'p4.png', 'p5.png', '.hidden.png'];
  const { styles } = packs(t, { previews: names });
  for (const file of ['p1.png', 'p2.png', 'p3.png', 'p4.png', 'p5.png']) writeFileSync(join(styles, 'second', file), PNG);
  writeFileSync(join(styles, 'second', 'p2.png'), Buffer.from('not a png'));
  const card = styleCard(styles, 'second');
  // p1, p2 (not a PNG → left out), gone (missing → left out), p3: the first four safe distinct names, then readable ones.
  assert.equal(PREVIEWS_MAX, 4);
  assert.deepEqual(card.previews, [`data:image/png;base64,${PNG.toString('base64')}`, `data:image/png;base64,${PNG.toString('base64')}`]);
  assert.deepEqual(styleCard(styles, 'second', { thumb: false }).previews, [], 'no image read on the context path');
  assert.deepEqual(styleCard(styles, 'third').previews, [], 'a pack without previews');
});

/** ShimoPreview's suggestion methods over a plain object; the Remote peer, authorization and scope are stand-ins. */
function host(t, { authorize } = {}) {
  const { root, styles } = packs(t);
  const ws = join(root, 'ws');
  mkdirSync(ws, { recursive: true });
  const authorized = [];
  const self = Object.create(ShimoPreview.prototype);
  Object.assign(self, {
    ctx: {}, styleRoot: styles, styleCache: undefined, styleSuggestions: new StyleSuggestions(),
    styleMemory: new StyleMemory({ file: join(root, 'data', 'style-defaults.json'), atomic: { withFileLock, writeFileAtomic } }),
    remotePeer: () => ({ principal: 'owner' }),
    suite: { authorize: async (_peer, scope, _signal, action) => { authorized.push([scope.sessionId, action]); await authorize?.(); } },
    scopeOf: async sessionId => { if (sessionId !== 's1') throw new SuiteError('forbidden', 'The Session is no longer available.'); return { sessionId, workspaceRoot: ws }; },
  });
  return { self, ws, scope: { sessionId: 's1', workspaceRoot: ws }, authorized };
}

test('Host: the tool records the recommendation, the strip gets cards with reasons and the default, dismiss ends it', async t => {
  const { self, scope, authorized, ws } = host(t);
  const result = await self.aiSuggestStyles('s1', { topic: '重庆旅行', candidates: [{ style: 'second', reason: '照片多的旅行主题' }, { style: 'ghost', reason: 'x' }, { style: 'default', reason: '稳妥' }] });
  assert.deepEqual(result.candidates, [{ index: 1, style: 'second', name: '黑底画报 / Dark editorial', reason: '照片多的旅行主题' }, { index: 2, style: 'default', name: '蓝色品牌 / Blue Brand', reason: '稳妥' }]);
  assert.deepEqual(result.dropped, ['ghost']);
  assert.equal(result.defaultStyle, 'default');
  assert.match(result.next, /End your turn/);
  const strip = await self.styleSuggestion(scope, signal());
  assert.equal(strip.id, result.suggestionId);
  assert.equal(strip.topic, '重庆旅行');
  assert.deepEqual(strip.candidates.map(card => [card.id, card.reason]), [['second', '照片多的旅行主题'], ['default', '稳妥']]);
  assert.match(strip.candidates[1].thumb, /^data:image\/png;base64,/, 'the same card reader as the picker');
  assert.equal(strip.defaultStyle.id, 'default');
  assert.deepEqual(authorized, [['s1', 'open']]);
  // The workspace default, when there is one, is what "use the default" means.
  await self.styleMemory.set(ws, 'third');
  assert.equal((await self.styleSuggestion(scope, signal())).defaultStyle.id, 'third');
  assert.deepEqual(await self.dismissStyleSuggestion(scope, 'other', signal()), { dismissed: false });
  assert.deepEqual(await self.dismissStyleSuggestion(scope, result.suggestionId, signal()), { dismissed: true });
  assert.equal(await self.styleSuggestion(scope, signal()), null);
  await assert.rejects(self.aiSuggestStyles('s1', { candidates: [{ style: 'ghost' }] }), { code: 'invalid-edit', message: /available: (?:[a-z0-9-]+, )*default, (?:[a-z0-9-]+, )*second, (?:[a-z0-9-]+, )*third\b/ });
  await assert.rejects(self.aiSuggestStyles('gone', { candidates: [{ style: 'default' }] }), /no longer available/);
  const denied = host(t, { authorize: async () => { throw new SuiteError('forbidden', 'Access to the document was denied.'); } });
  await assert.rejects(denied.self.styleSuggestion(denied.scope, signal()), /denied/);
});

test('Host: runtime context lists the recommendation on the first step after the strip showed it, and the strip is then empty', async t => {
  const { self, scope, ws } = host(t);
  await self.aiSuggestStyles('s1', { candidates: [{ style: 'second', reason: 'a' }, { style: 'default', reason: 'b' }] });
  assert.equal(self.styleContext(ws, 's1'), '', 'the model\'s step that ends the turn: nothing yet');
  await self.styleSuggestion(scope, signal());
  const text = self.styleContext(ws, 's1');
  assert.match(text, /1\) second \(黑底画报 \/ Dark editorial\), 2\) default \(蓝色品牌 \/ Blue Brand\)/);
  assert.equal(await self.styleSuggestion(scope, signal()), null, 'answered: no strip');
  assert.equal(self.styleContext(ws, 'other'), '', 'other Sessions see nothing');
  await self.styleMemory.set(ws, 'third');
  assert.match(self.styleContext(ws, 's1'), /^Workspace default PPT style \(shimo-pptx\): third .*\nPPT styles you recommended last turn/s, 'both, workspace first');
  assert.match(self.styleContext(ws), /^Workspace default PPT style/, 'no Session: the workspace part only');
});

test('tools: presentation_suggest_styles is registered and reaches the Host with the agent Session; the context gets the Session id', async () => {
  const tools = new Map(), calls = [], entries = [];
  const service = {
    aiSuggestStyles: async (sessionId, args) => { calls.push([sessionId, args]); return { suggestionId: 'g1' }; },
    styleContext: (cwd, sessionId) => `${cwd}|${sessionId}`,
  };
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service, get: () => undefined,
    inject: (_names, apply) => apply({ systemPrompt: { context: entry => entries.push(entry) } }) });
  const tool = tools.get('presentation_suggest_styles');
  assert.ok(tool, 'registered');
  assert.match(tool.description, /end your turn/i);
  const candidates = [{ style: 'second', reason: 'r' }];
  assert.deepEqual(await tool.execute({ candidates, topic: 't' }, { agent: { session: { id: 's9' } } }), { suggestionId: 'g1' });
  assert.deepEqual(calls, [['s9', { candidates, topic: 't' }]]);
  await assert.rejects(tool.execute({ candidates }, {}), /need an agent Session/);
  assert.equal(entries[0].text({ agent: { session: { id: 's9', header: { cwd: '/ws' } } } }), '/ws|s9');
});
