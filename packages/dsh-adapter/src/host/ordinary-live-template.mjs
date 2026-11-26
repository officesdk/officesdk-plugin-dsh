import { runTemplateScript } from './template-python.mjs';
import {mkdir,mkdtemp,readFile,rm} from 'node:fs/promises';
import {join} from 'node:path';
export async function ordinaryLiveTemplate({root,dataRoot,template,signal,rows}) {
  const folder=join(dataRoot,'ordinary-live-generation');await mkdir(folder,{recursive:true,mode:0o700});
  const dir=await mkdtemp(join(folder,'template-'));
  try {
    const output=join(dir,'prepared.docx');
    await runTemplateScript('shimo-docx', 'prepare_live_template.py', ['--source',join(root,`${template}.docx`),'--out',output,...(rows ? ['--rows',String(rows)] : []),...(template==='project-proposal'?['--preserve-rows']:[])],{cwd:dir,signal});
    return await readFile(output);
  } finally { await rm(dir,{recursive:true,force:true}).catch(()=>{}); }
}
