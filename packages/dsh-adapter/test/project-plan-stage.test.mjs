import {test} from 'node:test';import assert from 'node:assert/strict';
import {finishProjectTable} from '../src/host/project-plan-stage.mjs';
function fixture({failSave=false,stopDuringFill=false}={}) {
 const events=[],generation={template:'project-plan',milestoneCount:1,stopped:false},view={writer:true,closed:false,stale:false,aiActive:0,unsaved:true,id:'view',peer:{},realPath:'/w/file.docx'};
 const service={scopeOf:async()=>({}),generationOf:async()=>generation,kitViewFor:()=>view,writer:{writeAccess:async()=>'allowed'},kitDocs:{whenIdle:async()=>{},save:async()=>{events.push('save');if(failSave)throw Error('save failed');view.unsaved=false;return {version:'v1'}}},agents:{request:async(_s,{kind})=>{events.push(kind);if(kind==='table-fill'){if(stopDuringFill)generation.stopped=true;return {revision:3}}return {}}}};
 return {service,events,generation,view};
}
test('fills and saves in the same editor without reloading',async()=>{const f=fixture();const r=await finishProjectTable(f.service,'s',{path:'file.docx',milestones:[{stage:'启动'}]});assert.equal(r.saved,true);assert.deepEqual(f.events,['table-freeze','table-fill','save','table-mark-saved','table-release']);assert.equal(f.view.tableLock,false)});
test('save failure preserves edits and releases the frame',async()=>{const f=fixture({failSave:true});await assert.rejects(finishProjectTable(f.service,'s',{path:'file.docx',milestones:[{}]}),/save failed/);assert.equal(f.view.unsaved,true);assert.ok(f.events.includes('table-release'));assert.equal(f.generation.tableFrozen,false)});
test('stopping after filling prevents a save',async()=>{const f=fixture({stopDuringFill:true});await assert.rejects(finishProjectTable(f.service,'s',{path:'file.docx',milestones:[{}]}),/changed/);assert.ok(!f.events.includes('save'));assert.ok(f.events.includes('table-release'))});
test('rejects a mismatched row count or invalid values before editing',async()=>{const f=fixture();await assert.rejects(finishProjectTable(f.service,'s',{path:'file.docx',milestones:[{},{}]}),/exactly/);await assert.rejects(finishProjectTable(f.service,'s',{path:'file.docx',milestones:[{stage:17}]}),/milestone/);assert.equal(f.events.length,0)});
