#!/usr/bin/env python3
"""Fill an original reporting deck, keeping its editable native chart and shapes."""
import argparse, json, math, os, re, tempfile, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'

def positive_values(values, count, label):
 if not isinstance(values,list) or len(values)!=count or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or v<0 for v in values):
  raise ValueError(f'{label} must have {count} finite nonnegative numbers')
 return values

def build(source, spec_path, out):
 source,out=Path(source),Path(out); spec=json.loads(Path(spec_path).read_text());slides=spec.get('slides')
 if out.exists(): raise FileExistsError(out)
 if not isinstance(slides,list) or len(slides)!=8 or any(not isinstance(x,dict) for x in slides): raise ValueError('slides must contain eight name-to-text objects')
 for page in slides:
  if any(not isinstance(k,str) or not isinstance(v,str) or len(v)>500 for k,v in page.items()):raise ValueError('invalid slide text')
 chart=spec.get('chart');bars=spec.get('bars'); statuses=spec.get('statuses',{})
 if not isinstance(statuses,dict) or any(k not in ('page2','page6') for k in statuses):raise ValueError('invalid statuses')
 for key,count in [('page2',4),('page6',3)]:
  if key in statuses and (not isinstance(statuses[key],list) or len(statuses[key])!=count or any(v not in ('green','amber','red','neutral') for v in statuses[key])):raise ValueError(f'{key} statuses invalid')
 if chart is not None:
  if not isinstance(chart,dict) or not isinstance(chart.get('title',''),str):raise ValueError('chart must be an object with a text title')
  vals=positive_values(chart.get('values'),3,'chart.values');cats=chart.get('categories')
  if not isinstance(cats,list) or len(cats)!=3 or any(not isinstance(c,str) or len(c)>40 for c in cats) or sum(vals)<=0:raise ValueError('chart needs three labels and nonzero values')
 if bars is not None: positive_values(bars,5,'bars')
 fd,tmp=tempfile.mkstemp(prefix='.ppt-',suffix='.pptx',dir=out.parent)
 try:
  with zipfile.ZipFile(source) as src, os.fdopen(fd,'wb') as raw, zipfile.ZipFile(raw,'w') as dst:
   for member in src.infolist():
    data=src.read(member.filename); m=re.fullmatch(r'ppt/slides/slide(\d+)\.xml',member.filename)
    if m:
     idx=int(m.group(1))-1;root=ET.fromstring(data);seen=set();tree=root.find('.//'+P+'spTree')
     maxwidth=max([int(n.get('cx')) for s in root.iter(P+'sp') if (c:=s.find('./'+P+'nvSpPr/'+P+'cNvPr')) is not None and re.fullmatch(r'data-bar-\d+',c.get('name','')) for n in s.findall('./'+P+'spPr/'+A+'xfrm/'+A+'ext')] or [0])
     for shape in list(tree):
      prop=shape.find('./'+P+'nvSpPr/'+P+'cNvPr'); name=prop.get('name','') if prop is not None else ''
      status=re.fullmatch(r'status-dot-(\d+)' if idx==1 else r'status-(\d+)' if idx==5 else r'(?!)',name)
      if status:
       color={'green':'219468','amber':'D49612','red':'CE3B3B','neutral':'A2AFBA'}[statuses.get('page2' if idx==1 else 'page6',['neutral']*(4 if idx==1 else 3))[int(status.group(1))]]
       fill=shape.find('./'+P+'spPr/'+A+'solidFill/'+A+'srgbClr')
       if fill is not None:fill.set('val',color)
      if idx==3 and re.fullmatch(r'data-bar-\d+',name):
       if bars is None or not max(bars):tree.remove(shape);continue
       ext=shape.find('./'+P+'spPr/'+A+'xfrm/'+A+'ext');ext.set('cx',str(max(1,round(maxwidth*bars[int(name[-1])]/max(bars)))))
      texts=list(shape.iter(A+'t'))
      if texts:
       value=slides[idx].get(name,'');texts[0].text=value
       if idx==2 and (name=='center-line0' or re.fullmatch(r'legend[1-3]-line0',name)):
        if chart:
         i=0 if name=='center-line0' else int(name[6])-1
         percentage=f'{100*vals[i]/sum(vals):.1f}'.rstrip('0').rstrip('.')+'%'
         value=percentage if name=='center-line0' else f'{cats[i]} {percentage}'
        else:value=''
        texts[0].text=value
       if idx==3 and re.fullmatch(r'data-value-[0-4]-line0',name) and bars is None:texts[0].text=''
       for t in texts[1:]:t.text=''
       seen.add(name)
     unknown=set(slides[idx])-seen
     if unknown:raise ValueError(f'unknown slide {idx+1} slots: {sorted(unknown)}')
     data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    elif re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',member.filename):
     root=ET.fromstring(data)
     for t in root.iter(A+'t'):t.text=''
     data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    elif member.filename=='ppt/slides/charts/chart1.xml':
     root=ET.fromstring(data);ser=root.find('.//'+C+'ser');tx=ser.find('./'+C+'tx/'+C+'v')
     if tx is not None:tx.text=chart.get('title','') if chart else ''
     for typ,values in [('cat',chart['categories'] if chart else ['','','']),('val',chart['values'] if chart else [0,0,0])]:
      for pt in ser.findall('./'+C+typ+'/*/'+C+'pt'):
       index=int(pt.get('idx'));node=pt.find(C+'v');node.text=str(values[index])
     data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    dst.writestr(member,data)
  os.link(tmp,out)
 finally:
  if os.path.exists(tmp):os.unlink(tmp)
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--source',required=True);p.add_argument('--spec',required=True);p.add_argument('--out',required=True);a=p.parse_args();build(a.source,a.spec,a.out)
