import { SuiteError } from '../../../core/src/index.mjs';

/** Fill the pre-existing table behind its Agent frame, then save the same live view. */
export async function finishProjectTable(service, sessionId, { path, milestones }, signal, kind='project') {
  const fields=kind==='meeting'?['decision','owner','deadline']:kind==='proposal'?['description','quantity','unitPrice','amount']:['stage','deliverable','owner','deadline'];
  if (!Array.isArray(milestones) || milestones.length < 1 || milestones.length > 12 || milestones.some(item=>!item || typeof item!=='object' || Array.isArray(item) || Object.keys(item).some(k=>!fields.includes(k)) || Object.values(item).some(v=>typeof v!=='string' || v.length>2000)))
    throw new SuiteError('invalid-edit','Provide 1–12 milestone objects with text fields.');
  const scope=await service.scopeOf(sessionId), generation=await service.generationOf(sessionId,scope,path);
  if(!generation || generation.stopped || generation.template!==(kind==='meeting'?'meeting-minutes':kind==='proposal'?'project-proposal':'project-plan') || generation.tableFrozen)
    throw new SuiteError('invalid-edit','The project plan is not an active generation.');
  const count=kind==='meeting'?generation.decisionCount:kind==='proposal'?generation.proposalCount:generation.milestoneCount;
  if(milestones.length!==count) throw new SuiteError('invalid-edit',`Provide exactly ${count} table rows.`);
  const view=service.kitViewFor(sessionId,path);
  if(!view?.writer || view.closed || view.stale || view.aiActive) throw new SuiteError('invalid-edit','The project plan editor is not ready.');
  if(await service.writer.writeAccess(scope,path)!=='allowed') throw new SuiteError('sandbox-denied','The generated file must be in a writable workspace.');
  generation.tableFrozen=true;let freezeAttempted=false;
  try {
    signal?.throwIfAborted();freezeAttempted=true;
    await service.agents.request(sessionId,{kind:'table-freeze',path},signal);
    if(generation.stopped) throw new SuiteError('cancelled','Stopped before filling the table.');
    const written=await service.agents.request(sessionId,{kind:'table-fill',path,milestones,tableKind:kind},signal);
    await service.kitDocs.whenIdle(view.id);
    if(generation.stopped || view.closed || view.stale || view.aiActive || await service.generationOf(sessionId,scope,path)!==generation)
      throw new SuiteError('cancelled','The editor changed while the table was being filled; its edits remain in the editor.');
    view.tableLock=true;
    const saved=await service.kitDocs.save(view.id,view.peer,{signal});
    if(view.unsaved || !saved.version) throw new SuiteError('unavailable','The table was filled but could not be safely saved; its edits remain in the editor.');
    await service.agents.request(sessionId,{kind:'table-mark-saved',path,version:saved.version,revision:written.revision},signal);
    return {applied:true,saved:true,path:view.realPath,version:saved.version};
  } finally {
    view.tableLock=false;generation.tableFrozen=false;
    if(freezeAttempted) await service.agents.request(sessionId,{kind:'table-release',path},undefined).catch(()=>{});
  }
}
