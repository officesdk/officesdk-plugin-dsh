#!/usr/bin/env python3
"""Fill each built-in Word template with sample content, for the inspiration library's covers.

Usage: python template_samples.py --out <dir>

Each sample starts from the template file document_create opens (packages/dsh-adapter/skills/shimo-docx/templates/
*.docx) and only adds text with that template's own paragraph styles (Title, Heading 1, Heading 2, Normal), the way the
live agent writes a document. render-template-covers.mjs then opens the results in the real editor and screenshots the
first page. Writes word-<id>.docx; prints each path.
"""
import argparse
from pathlib import Path

from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[2] / 'packages' / 'dsh-adapter' / 'skills'
DOCX = ROOT / 'shimo-docx' / 'templates'


def write(doc, items, *, fill_first):
    """`items`: (style, text[, first_line_indent_pt]). The template's single empty paragraph takes the first item."""
    for index, item in enumerate(items):
        style, text = item[0], item[1]
        if index == 0 and fill_first:
            p = doc.paragraphs[0]
            p.style = doc.styles[style]
            p.add_run(text)
        else:
            p = doc.add_paragraph(text, style=doc.styles[style])
        if len(item) > 2:
            p.paragraph_format.first_line_indent = Pt(item[2])


GONGWEN = [
    ('Title', '2026年上半年工作总结'),
    ('Normal', '今年以来，在上级部门的正确领导下，我们紧紧围绕年度目标任务，坚持稳中求进工作总基调，统筹推进各项工作，取得了阶段性成效。现将上半年工作情况总结如下。'),
    ('Heading 1', '一、主要工作完成情况'),
    ('Heading 2', '（一）重点项目稳步推进'),
    ('Normal', '全年确定的12个重点项目已开工10个，完成投资4.6亿元，占年度计划的58%。其中，智慧政务平台二期已完成主体建设，进入联调测试阶段。'),
    ('Heading 2', '（二）营商环境持续优化'),
    ('Normal', '深化“一网通办”改革，政务服务事项网上可办率达到98%，企业开办时间压缩至1个工作日以内，群众满意度稳步提升。'),
    ('Heading 1', '二、存在的问题'),
    ('Normal', '一是部分项目前期手续办理进度偏慢；二是基层人员业务能力有待进一步提升。'),
]

def gongwen(out):
    doc = Document(DOCX / 'gongwen.docx')
    write(doc, GONGWEN, fill_first=True)
    doc.save(out)


SAMPLES = (
    ('word-gongwen.docx', gongwen),
)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('--out', required=True)
    out = Path(parser.parse_args().out)
    out.mkdir(parents=True, exist_ok=True)
    for name, build in SAMPLES:
        build(out / name)
        print(out / name)


if __name__ == '__main__':
    main()
