#!/usr/bin/env python3
"""Make a blank live-writing copy of an approved document template."""
import argparse, copy, os, tempfile, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SPACE='{http://www.w3.org/XML/1998/namespace}space'

def prepare(source, output, rows=None, preserve_rows=False):
    source,output=Path(source),Path(output)
    if output.exists():raise FileExistsError(output)
    with zipfile.ZipFile(source) as src:
        doc=ET.fromstring(src.read('word/document.xml'));body=doc.find(W+'body')
        for paragraph in body.findall(W+'p'):
            nodes=paragraph.findall('.//'+W+'t')
            if not nodes or not any((node.text or '').strip() for node in nodes): continue
            for node in nodes:node.text=''
            # Sample mail-merge/date fields and their line breaks must not remain live behind the marker.
            for run in paragraph.iter(W+'r'):
                for child in list(run):
                    if child.tag in {W+'fldChar',W+'instrText'} or (child.tag==W+'br' and child.get(W+'type') not in {'page','column'}):run.remove(child)
            nodes[0].text='\u2800';nodes[0].set(SPACE,'preserve')
            run=next(r for r in paragraph.iter(W+'r') if nodes[0] in list(r))
            props=run.find(W+'rPr')
            if props is None:props=ET.Element(W+'rPr');run.insert(0,props)
            color=props.find(W+'color')
            if color is None:color=ET.SubElement(props,W+'color')
            color.set(W+'val','FFFFFF')
        if rows is not None:
            if not 1 <= rows <= 12:raise ValueError('Rows must be between 1 and 12')
            table=body.find(W+'tbl');existing=table.findall(W+'tr')
            if not preserve_rows:
                for row in existing[2:]:table.remove(row)
                for _ in range(rows-1):table.append(copy.deepcopy(existing[1]))
            elif len(existing) != rows+1:raise ValueError('Unexpected template table rows')
            for row in table.findall(W+'tr')[1:]:
                for cell in row.findall(W+'tc'):
                    for run in cell.iter(W+'r'):
                        for child in list(run):
                            if child.tag in {W+'fldChar',W+'instrText'} or (child.tag==W+'br' and child.get(W+'type') not in {'page','column'}):run.remove(child)
                    nodes=cell.findall('.//'+W+'t')
                    if nodes:
                        nodes[0].text='待填写'
                        for node in nodes[1:]:node.text=''
                    else:
                        para=cell.find('.//'+W+'p');run=para.find(W+'r')
                        if run is None:run=ET.SubElement(para,W+'r')
                        ET.SubElement(run,W+'t').text='待填写'
        xml=ET.tostring(doc,encoding='utf-8',xml_declaration=True)
        styles_xml=src.read('word/styles.xml') if 'word/styles.xml' in src.namelist() else None
        if styles_xml and source.name=='project-proposal.docx':
            styles=ET.fromstring(styles_xml)
            referenced=src.read('word/numbering.xml') if 'word/numbering.xml' in src.namelist() else b''
            for style in list(styles.findall(W+'style')):
                sid=style.get(W+'styleId','')
                if style.get(W+'type')=='numbering' and style.find(W+'pPr/'+W+'numPr') is None and sid and sid.encode() not in xml and sid.encode() not in referenced:
                    styles.remove(style)
            styles_xml=ET.tostring(styles,encoding='utf-8',xml_declaration=True)
        fd,tmp=tempfile.mkstemp(prefix='.live-',suffix='.docx',dir=output.parent)
        try:
            with os.fdopen(fd,'wb') as raw,zipfile.ZipFile(raw,'w') as dst:
                for item in src.infolist():dst.writestr(item,xml if item.filename=='word/document.xml' else styles_xml if item.filename=='word/styles.xml' and styles_xml else src.read(item.filename))
            os.link(tmp,output)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--source',required=True);p.add_argument('--out',required=True);p.add_argument('--rows',type=int);p.add_argument('--preserve-rows',action='store_true');a=p.parse_args();prepare(a.source,a.out,a.rows,a.preserve_rows)
