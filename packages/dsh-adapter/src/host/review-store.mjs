import {createHash} from 'node:crypto';
import {mkdir, readFile} from 'node:fs/promises';
import {join} from 'node:path';
import {KitError} from '../../../kit-host/src/index.mjs';

const validId = id => typeof id==='string' && id.length>0 && id.length<=256;
const validRecord = r => validId(r?.id) && !r.accepted && ['docx','xlsx','pptx'].includes(r.kind) &&
  typeof r.before==='string' && typeof r.after==='string' && r.revert?.anchor;

export class ReviewStore {
  constructor(directory, atomic) { this.directory=directory; this.atomic=atomic; }

  file(view) {
    // A file can be opened through different chats, relative paths or symlinks.
    // Only the host's authorized, canonical file path identifies its ledger.
    const identity=JSON.stringify([view.principal,view.realPath]);
    return join(this.directory,createHash('sha256').update(identity).digest('hex')+'.json');
  }

  async read(view) {
    let data;
    try { data=await readFile(this.file(view),'utf8'); }
    catch (error) { if(error.code==='ENOENT')return {revision:null,changes:[]}; throw error; }
    const saved=JSON.parse(data);
    if(saved.v!==1 || !validId(saved.revision) || !Array.isArray(saved.changes) || !saved.changes.every(validRecord))
      throw new Error('Invalid change history file');
    return {revision:saved.revision,changes:saved.changes};
  }

  async commit(view, patch) {
    const {base,revision,put,remove}=patch ?? {};
    if (!(base===null || validId(base)) || !validId(revision) || revision===base || !Array.isArray(put) ||
        !put.every(validRecord) || !Array.isArray(remove) || !remove.every(validId))
      throw new KitError('invalid-body','Invalid change history update.');
    await mkdir(this.directory,{recursive:true,mode:0o700});
    const file=this.file(view);
    return this.atomic.withFileLock(file,async()=>{
      const saved=await this.read(view);
      if (saved.revision===revision) return {revision}; // The previous acknowledgement may have been lost.
      if (saved.revision!==base) throw new KitError('edit-conflict','Change history was updated in another view. Reopen the document.');
      const records=new Map(saved.changes.map(r=>[r.id,r]));
      for (const id of remove) records.delete(id);
      for (const r of put) records.set(r.id,r);
      await this.atomic.writeFileAtomic(file,JSON.stringify({v:1,revision,changes:[...records.values()]}),{mode:0o600});
      return {revision};
    });
  }
}
