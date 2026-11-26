#!/usr/bin/env python3
"""Blank the text of a packaged editable presentation while retaining its objects and page order."""
import argparse, os, re, tempfile, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
def build(source, out, slides=None):
 source,out=Path(source),Path(out)
 if out.exists():raise FileExistsError(out)
 fd,tmp=tempfile.mkstemp(prefix='.ppt-',suffix='.pptx',dir=out.parent)
 try:
  with zipfile.ZipFile(source) as src,os.fdopen(fd,'wb') as raw,zipfile.ZipFile(raw,'w') as dst:
   for member in src.infolist():
    data=src.read(member.filename)
    if re.fullmatch(r'ppt/slides/slide\d+\.xml',member.filename) or re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',member.filename):
     root=ET.fromstring(data)
     # These original line-chart marks encode sample measurements. Keep the axes
     # and editable text slots, but never carry invented measurements into a new deck.
     if source.parent.name=='slide-008' and member.filename=='ppt/slides/slide2.xml':
      tree=root.find('.//'+P+'spTree')
      for child in list(tree):
       nv=child.find('./'+P+'nvSpPr/'+P+'cNvPr')
       if nv is not None and re.fullmatch(r'(point|segment)-\d+',nv.get('name','')):tree.remove(child)
     if source.parent.name=='slide-008' and member.filename=='ppt/slides/slide6.xml':
      tree=root.find('.//'+P+'spTree')
      for child in list(tree):
       nv=child.find('./'+P+'nvSpPr/'+P+'cNvPr')
       if nv is not None and re.fullmatch(r'bar[1-4]',nv.get('name','')):tree.remove(child)
     for node in root.iter(A+'t'):node.text=''
     if slides is not None and member.filename.startswith('ppt/slides/slide'):
      index = int(re.search(r'slide(\d+)\.xml$', member.filename).group(1)) - 1
      values = slides[index]
      used = set()
      for shape in root.iter(P+'sp'):
       info = shape.find(P+'nvSpPr/'+P+'cNvPr')
       name = info.get('name') if info is not None else None
       texts = list(shape.iter(A+'t'))
       if name in values and texts:
        texts[0].text = values[name]
        used.add(name)
      if used != set(values):
       raise ValueError('Unknown text slot on slide ' + str(index + 1))
     data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    dst.writestr(member,data)
  os.link(tmp,out)
 finally:
  if os.path.exists(tmp):os.unlink(tmp)
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--source',required=True);p.add_argument('--out',required=True);a=p.parse_args();build(a.source,a.out)
