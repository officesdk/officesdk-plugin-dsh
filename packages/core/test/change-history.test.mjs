import test from 'node:test';
import assert from 'node:assert/strict';
import {overlapsPreviousChange, recordChange, restoreChangeHistory} from '../src/browser/change-history.mjs';
import {previousWriterChange} from '../src/browser/selection-edit.mjs';

function word(id, paragraph, before, after) {
  const anchor = {paragraphs:[paragraph,paragraph], start:0, end:Array.from(after).length};
  return {id,kind:'docx',before,after,revert:{anchor,original:after,text:before},
    located:{anchor,view:{first:paragraph,texts:[after]}}};
}

test('nearby changes retain their own text anchors, while overwritten cards share a verified location', () => {
  const first=word('first',0,'甲','乙');
  first.located.view.texts=['乙春日微风'];
  assert.equal(previousWriterChange([first],['乙春夜微风'],first.revert.anchor,'乙'),'first');
  const wider=word('wider',0,'乙春日','新春');
  wider.located.view.texts=['新春微风'];
  assert.equal(overlapsPreviousChange(first,wider),true);
  wider.before='别处内容';
  assert.equal(overlapsPreviousChange(first,wider),false);
});

test('repeated Word edits update the same card in place and undo to its first original', () => {
  const first = word('1',0,'甲','乙'), other = word('2',1,'甲','乙'), final = word('3',0,'乙','最后的内容');
  const prior = previousWriterChange([first,other],['乙','乙'],first.revert.anchor,'乙');
  const cards = recordChange([first,other],final,prior);
  assert.equal(prior,'1');
  assert.deepEqual(cards.map(r=>r.id),['1','2']);
  assert.deepEqual(cards[0],{...final,id:'1',latestId:'3',before:'甲',revert:{...final.revert,text:'甲'}});
  assert.deepEqual(restoreChangeHistory([first,other,final]),cards);
  assert.deepEqual(restoreChangeHistory(cards),cards);
});

test('a later user change at the same coordinates starts a new edit', () => {
  const first = word('1',0,'甲','乙');
  assert.equal(previousWriterChange([first],['丙'],first.revert.anchor,'丙'),undefined);
  assert.equal(restoreChangeHistory([first,word('2',0,'丙','甲')]).length,2);
});

test('repeated edits to the same sheet cells preserve the original formula for undo', () => {
  const anchor={sheet:'Sheet1',row:0,column:0,rowCount:1,columnCount:1};
  const first={id:'x',kind:'xlsx',before:'=1+1',after:'3',revert:{anchor,cells:[[{formula:'1+1'}]],written:[[3]]}};
  const latest={id:'y',kind:'xlsx',before:'3',after:'4',revert:{anchor,cells:[[{value:3}]],written:[[4]]}};
  const [card]=restoreChangeHistory([first,latest]);
  assert.deepEqual(card,{...latest,id:'x',latestId:'y',before:'=1+1',revert:{...latest.revert,cells:first.revert.cells}});
});
