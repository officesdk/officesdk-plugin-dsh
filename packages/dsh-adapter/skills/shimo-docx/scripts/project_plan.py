#!/usr/bin/env python3
"""Fill the packaged Word project plan into a new file without overwriting."""
import argparse, copy, json, os, tempfile, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SPACE='{http://www.w3.org/XML/1998/namespace}space'

def replace_text(element, value):
    texts=element.findall('.//'+W+'t')
    if len(texts)!=1: raise ValueError('Unexpected project plan field')
    target=texts[0]; run=next(r for r in element.iter(W+'r') if target in list(r))
    pos=list(run).index(target);run.remove(target)
    for i,part in enumerate(str(value).split('\n')):
        if i: run.insert(pos,ET.Element(W+'br'));pos+=1
        t=ET.Element(W+'t');t.set(SPACE,'preserve');t.text=part;run.insert(pos,t);pos+=1

def build(data, out, template=None):
    template = template or Path(__file__).resolve().parent.parent/'templates/project-plan.docx'
    out=Path(out)
    if out.exists(): raise FileExistsError(f'Output already exists: {out}')
    milestones=data.get('milestones',[])
    if not isinstance(milestones,list) or not all(isinstance(x,dict) for x in milestones):raise ValueError('milestones must be an array of objects')
    out.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(template) as src:
        doc=ET.fromstring(src.read('word/document.xml'));body=doc.find(W+'body');paras=body.findall(W+'p')
        if len(paras)!=11:raise ValueError('Unexpected project plan template structure')
        for idx,key,spaces in [(1,'name',5),(2,'owner',9),(3,'period',5)]:
            texts=paras[idx].findall('.//'+W+'t')
            if len(texts)!=2:raise ValueError('Unexpected metadata field')
            texts[1].text=' '*spaces+str(data.get(key) or '待填写');texts[1].set(SPACE,'preserve')
        for idx,key in [(5,'goal'),(7,'scope'),(10,'risks')]:replace_text(paras[idx],data.get(key) or '待填写')
        if data.get('startBlank'):
            for para in paras:
                nodes=para.findall('.//'+W+'t')
                for node in nodes: node.text=''
                if nodes:
                    nodes[0].text='\u2800'; nodes[0].set(SPACE,'preserve')
                    run=next(r for r in para.iter(W+'r') if nodes[0] in list(r))
                    props=run.find(W+'rPr')
                    if props is None: props=ET.Element(W+'rPr');run.insert(0,props)
                    color=props.find(W+'color')
                    if color is None: color=ET.SubElement(props,W+'color')
                    color.set(W+'val','FFFFFF')
        table=body.find(W+'tbl'); rows=table.findall(W+'tr')
        if len(rows)!=2:raise ValueError('Unexpected project plan table')
        for _ in range(max(0,len(milestones)-1)):table.append(copy.deepcopy(rows[1]))
        rows=table.findall(W+'tr')
        for idx,row in enumerate(rows[1:]):
            item=milestones[idx] if idx<len(milestones) else {}
            cells=row.findall(W+'tc')
            if len(cells)!=4:raise ValueError('Unexpected milestone table columns')
            for cell,key in zip(cells,['stage','deliverable','owner','deadline']):replace_text(cell,item.get(key) or '待填写')
        xml=ET.tostring(doc,encoding='utf-8',xml_declaration=True)
        fd,tmp=tempfile.mkstemp(prefix='.project-plan-',suffix='.docx',dir=out.parent)
        try:
            with os.fdopen(fd,'wb') as raw,zipfile.ZipFile(raw,'w') as dst:
                for item in src.infolist():dst.writestr(item,xml if item.filename=='word/document.xml' else src.read(item.filename))
            os.link(tmp,out)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)
    print(str(out))
if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--input',required=True);parser.add_argument('--out',required=True)
    parser.add_argument('--source', type=Path)
    args=parser.parse_args()
    with open(args.input,encoding='utf-8') as f:data=json.load(f)
    build(data, args.out, args.source)
