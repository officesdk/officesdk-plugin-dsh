import test from 'node:test';
import assert from 'node:assert/strict';
import { editQuoteRule, fitFilename, registerEditQuote } from '../src/client/edit-quote.mjs';
import { parallelEditMessage, selectionMessage } from '../../core/src/agent/selection.mjs';

const parse = text => editQuoteRule.match(text, 0)?.data;
const quote = text => text.split('\n').map(line => '> ' + line).join('\n');
const request = {kind:'pptx', original:'武汉\n江城印象', anchor:{slideId:'slide-1',shapeId:'title'}, instruction:'丰富一下标题'};

test('local slide edits display the filename and instruction separately and expose only the selected text', () => {
  const message = selectionMessage({id:'sel-750d08db',name:'武汉介绍.pptx',request});
  const ref = quote('编辑目标：C:\\资料\\武汉介绍.pptx\n\n'+message);
  const expected = {name:'武汉介绍.pptx',original:request.original,instruction:request.instruction,en:false};
  for (const sent of [ref+'\n\n'+request.instruction, request.instruction+'\n\n'+ref, '编辑目标：武汉介绍.pptx\n\n'+message+'\n\n'+request.instruction, message]) {
    assert.deepEqual(parse(sent), expected);
    assert.equal(editQuoteRule.match(sent, 1), undefined);
  }
});

test('parallel edits retain paragraph text, line breaks and a literal quote character on replay', () => {
  const msg = parallelEditMessage({name:'报告.docx', path:'草稿/报告.docx', request:{
    kind:'docx',instruction:'润色',original:'甲\n\n> 乙',anchor:{paragraphs:[1,2]}
  }});
  assert.deepEqual(parse(msg.replaceAll('\n','\r\n')), {name:'报告.docx',original:'甲\n\n> 乙',instruction:'润色',en:false});
});

test('Excel quotes expose cell values without the editing protocol', () => {
  const req = {kind:'xlsx',anchor:{sheet:'Q4',row:0,column:0,rowCount:1,columnCount:2},cells:[[{text:'A|B'},{text:'100'}]],instruction:'Update these cells'};
  const selected = selectionMessage({id:'sel-750d08db',name:'Plan.xlsx',request:req,lang:'en'});
  assert.deepEqual(parse(selected), {name:'Plan.xlsx',original:'A|B\t100',instruction:req.instruction,en:true});
});

test('ordinary DSH references pass through and document-only messages have no fabricated selection', () => {
  assert.equal(parse('解释\n\n> 普通引用'), undefined);
  assert.equal(parse('解释\n\n> 编辑目标：任务.txt'), undefined);
  assert.deepEqual(parse('> 编辑目标：总结.docx\n\n写个小结'), {name:'总结.docx', original:'', instruction:'写个小结', en:false});
});

test('filename truncation keeps the extension, both ends and whole grapheme clusters', () => {
  const measure = s => [...new Intl.Segmenter().segment(s)].length;
  const name = '研👩‍💻究项目最终定稿.pptx';
  assert.equal(fitFilename(name,20,measure),name);
  assert.equal(fitFilename(name,12,measure),'研👩‍💻...定稿.pptx');
  assert.equal(fitFilename('预算.xlsx',16,measure),'预算.xlsx');
});

test('a quoted instruction renders inside the existing DSH user-text factory', () => {
  let EditQuote;
  const slots = {
    inject: (_name, install) => install(),
    register: (_definition, component) => { EditQuote = component; },
  };
  const React = {
    createElement: (_tag, _props, ...children) => children,
    useEffect() {}, useLayoutEffect() {}, useRef: () => ({current:null}),
    useId: () => 'quote', useState: value => [value, () => {}],
  };
  registerEditQuote({uiConversation:{userText:{register() {}}}, slots}, React);
  const node = editQuoteRule.match('改写第二段\n\n> 编辑目标：示例.docx\n> 原文', 0);
  const content = EditQuote({node, renderFactorySlot() {
    assert.fail('DSH rejects re-entry into the conversation.user.text factory');
  }});
  const text = content.flat(Infinity).filter(value => typeof value === 'string');
  assert.ok(text.includes('示例.docx'));
  assert.ok(text.includes('改写第二段'));
});
