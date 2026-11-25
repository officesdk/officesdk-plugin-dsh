import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import * as atomic from '@deepseek-ai/dsh-atomic-write';
import {ReviewStore} from '../src/host/review-store.mjs';
import {createPendingReviews} from '../../core/src/browser/pending-reviews.mjs';

test('unresolved changes survive cold opens, a lost reply and overlapping writes without leaking between documents',async t=>{
  const dir=await mkdtemp(join(tmpdir(),'pending-reviews-'));
  t.after(()=>rm(dir,{recursive:true,force:true}));
  const view={principal:'owner',realPath:'/a/report.docx'};
  const store=new ReviewStore(dir,atomic);
  const a={id:'a',kind:'docx',before:'原句',after:'第一版',revert:{anchor:{paragraphs:[1,1],start:0,end:3},text:'原句',original:'第一版'}};
  const b={...a,id:'b'}, updated={...a,after:'第二版',latestId:'c'};
  const notices=[];
  let first=true;
  const queue=createPendingReviews(await store.read(view),async patch=>{
    const receipt=await store.commit(view,patch);
    if (first) {first=false;throw new Error('Connection lost after commit');}
    return receipt;
  },failed=>notices.push(failed));
  queue.stage([a]);
  const firstWrite=queue.flush();
  queue.stage([updated,b]);
  await assert.rejects(firstWrite,/Connection lost/);
  await queue.flush();
  assert.deepEqual(notices,[true,false]);
  const reopened=new ReviewStore(dir,atomic);
  const saved=await reopened.read({...view,sessionId:'another-chat'});
  assert.deepEqual(saved.changes,[updated,b]);
  assert.deepEqual((await reopened.read({...view,realPath:'/b/report.docx'})).changes,[]);
  assert.deepEqual((await reopened.read({...view,principal:'guest'})).changes,[]);
  await assert.rejects(store.commit(view,{base:null,revision:'stale',put:[],remove:['a']}),/another view/);
  queue.stage([updated]);
  await queue.flush();
  assert.deepEqual((await reopened.read(view)).changes,[updated]);
});
