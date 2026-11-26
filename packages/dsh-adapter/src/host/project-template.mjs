import { runTemplateScript } from './template-python.mjs';
import {mkdir,mkdtemp,writeFile,readFile,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {SuiteError} from '../../../core/src/index.mjs';
export async function projectTemplate({root,dataRoot,metadata,signal}) {
  if(!metadata || typeof metadata!=='object' || Array.isArray(metadata) || Object.keys(metadata).some(k=>!['name','owner','period','milestoneCount'].includes(k)) || Object.entries(metadata).some(([k,v])=>k==='milestoneCount' ? !Number.isInteger(v)||v<1||v>12 : typeof v!=='string'||v.length>2000))
    throw new SuiteError('invalid-edit','Project metadata must contain text values for name, owner and period.');
  const folder=join(dataRoot,'project-generation');await mkdir(folder,{recursive:true,mode:0o700});
  const dir=await mkdtemp(join(folder,'template-'));
  try {
    const input=join(dir,'input.json'),output=join(dir,'prepared.docx');
    await writeFile(input,JSON.stringify({...metadata,startBlank:true,milestones:Array.from({length:metadata.milestoneCount ?? 4},()=>({}))}),{mode:0o600});
    await runTemplateScript('shimo-docx', 'project_plan.py', ['--source', join(root, 'project-plan.docx'), '--input', input, '--out', output], { cwd: dir, signal });
    return await readFile(output);
  } finally { await rm(dir,{recursive:true,force:true}).catch(()=>{}); }
}
