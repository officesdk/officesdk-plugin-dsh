import test from 'node:test';
import assert from 'node:assert/strict';
import { checkBody, checkTitle } from './copy-check.mjs';

test('live titles must be two lines of 2–6 characters', () => {
  assert.deepEqual(checkTitle('厦门'), ['标题须两行']);
  // The feedback deck's one-line title fails; split too long it still fails on the line length.
  assert.deepEqual(checkTitle('城在绿意后面'), ['标题须两行']);
  assert.match(checkTitle('城在绿意\n后面的厦门岛屿').join('\n'), /「后面的厦门岛屿」7 字，每行须 2–6 字/);
  assert.match(checkTitle('城\n厦门').join('\n'), /「城」1 字/);
  assert.deepEqual(checkTitle('一\n二\n三').slice(0, 1), ['标题须两行']);
  assert.deepEqual(checkTitle('鹭岛\n厦门'), []);
  assert.deepEqual(checkTitle('厦门\n城市介绍'), []);
});

test('finished slides must not carry generation-process wording', () => {
  assert.deepEqual(checkBody('福建东南沿海，主岛四面环海。'), []);
  assert.deepEqual(checkBody('有两张图被限流，改用其他照片。'), ['成品页不要写过程说明']);
  assert.deepEqual(checkBody('裁切没问题，八页都已写入。'), ['成品页不要写过程说明']);
});
