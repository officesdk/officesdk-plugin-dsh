import test from 'node:test';
import assert from 'node:assert/strict';
import { SuiteError, fromAbort, isSuiteError } from '../src/errors.mjs';

// The release bundles Core separately into lib/host and lib/tools; a second module instance is the same situation.
const other = await import('../src/errors.mjs?second-copy');

test('a SuiteError made by another copy of Core is recognised', () => {
  assert.notEqual(other.SuiteError, SuiteError, 'the two copies have different classes');
  const foreign = new other.SuiteError('busy', 'The editor is busy.');
  assert.equal(foreign instanceof SuiteError, false);
  assert.equal(isSuiteError(foreign), true);
  assert.equal(other.isSuiteError(new SuiteError('forbidden', 'x')), true, 'and the other way round');
  assert.equal(isSuiteError(new SuiteError('timeout', 'x')), true);
});

test('errors that are not SuiteErrors are not mistaken for one', () => {
  const named = (code, name = 'SuiteError') => Object.assign(new Error('x'), { name, code });
  const abort = new AbortController(); abort.abort();
  for (const [label, error] of [
    ['plain Error', new Error('x')],
    ['Error with a suite code but another name', named('busy', 'Error')],
    ['AbortError (DOMException)', abort.signal.reason],
    ['AbortError (Error)', named('busy', 'AbortError')],
    ['unknown code', named('nope')],
    ['inherited key as code', named('toString')],
    ['constructor as code', named('constructor')],
    ['non-string code', named(429)],
    ['missing code', named(undefined)],
    ['null', null],
    ['undefined', undefined],
    ['plain object', { name: 'SuiteError' }],
  ]) assert.equal(isSuiteError(error), false, label);
});

test('fromAbort keeps a SuiteError reason from either copy and turns anything else into cancelled', () => {
  const abortWith = reason => { const controller = new AbortController(); controller.abort(reason); return controller.signal; };
  const foreign = new other.SuiteError('timeout', 'Too slow.');
  assert.equal(fromAbort(abortWith(foreign)), foreign);
  const own = new SuiteError('busy', 'Busy.');
  assert.equal(fromAbort(abortWith(own)), own);
  for (const reason of [undefined, new Error('x'), Object.assign(new Error('x'), { name: 'SuiteError', code: 'nope' })]) {
    const error = fromAbort(abortWith(reason));
    assert.equal(error.code, 'cancelled');
    assert.ok(error instanceof SuiteError);
  }
  assert.equal(fromAbort(undefined).code, 'cancelled');
});
