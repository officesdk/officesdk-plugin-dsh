#!/usr/bin/env python3
"""Generate an editable DOCX in the Chinese official-document style (shimo-docx).

Usage:
  python gongwen_docx.py --input doc.json --check-only
  python gongwen_docx.py --input doc.json --out doc.docx [--font-policy named|installed]

mode=plain   : 公文版式（页面、字体、字号、行距、页码），无红头、无红线。
mode=redhead : 在 plain 基础上加发文机关标志、发文字号、红线、落款、附件、版记等红头要素。
mode=auto    : 文种属公文类或给了发文机关标志时用 redhead，否则 plain。
"""
import argparse
import datetime
import json
import os
import re
import sys
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
STYLE = json.loads((ROOT / 'assets' / 'style.json').read_text(encoding='utf-8'))
T = STYLE['typography_pt']
REQUIRED = {'redhead': '发文机关标志', 'title': '标题', 'document_type': '文种', 'document_number': '发文字号',
            'recipient': '主送机关', 'signing_unit': '发文机关署名', 'date': '成文日期'}
OPTIONAL_LINES = ('serial', 'security', 'urgency', 'signer', 'subtitle', 'author')
PLACEHOLDER = re.compile(r'【(?:待填写|待确认|待核实|待补充|待定)[^】]*】|\{\{[^{}]+\}\}|待(?:填写|确认|核实|补充|确定)')
NUMBERED = re.compile(r'^\s*(\d+[.．、]|[一二三四五六七八九十]+、|[（(][\d一二三四五六七八九十]+[）)])')
KNOWN_KEYS = set(REQUIRED) | set(OPTIONAL_LINES) | {
    'mode', 'format_mode', 'seal_mode', 'colophon', 'blocks', 'body', 'attachments', 'fonts', 'title_lines',
    'draft_label', 'font_policy'}


# ---------------- helpers ----------------
def width(text):
    """Approximate width in 1/16 em units of a 16pt line (CJK = 16, ASCII = 8)."""
    return sum(16 if unicodedata.east_asian_width(c) in ('W', 'F', 'A') else 8 for c in text)


def pending(v):
    return not v or bool(PLACEHOLDER.search(v))


def line_text(data, key):
    v = data.get(key, '')
    if v is None:
        v = ''
    if not isinstance(v, str):
        raise ValueError(f'{key} 必须是字符串')
    if '\n' in v or '\r' in v:
        raise ValueError(f'{key} 必须是单行文本')
    return v.strip()


def installed_families():
    """Family names of font files in the usual macOS/Windows/Linux font folders (Pillow, no fc-list)."""
    from PIL import ImageFont
    dirs = ['/System/Library/Fonts', '/System/Library/Fonts/Supplemental', '/Library/Fonts',
            os.path.expanduser('~/Library/Fonts'), '/usr/share/fonts', '/usr/local/share/fonts',
            os.path.expanduser('~/.local/share/fonts'), os.path.expanduser('~/.fonts')]
    if os.environ.get('WINDIR'):
        dirs.append(os.path.join(os.environ['WINDIR'], 'Fonts'))
        if os.environ.get('LOCALAPPDATA'):
            dirs.append(os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'Fonts'))
    found = set()
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for root, _, files in os.walk(d):
            for f in files:
                if not f.lower().endswith(('.ttf', '.otf', '.ttc', '.otc')):
                    continue
                path = os.path.join(root, f)
                for idx in range(64 if f.lower().endswith(('.ttc', '.otc')) else 1):
                    try:
                        font = ImageFont.truetype(path, 12, index=idx)
                    except (OSError, ValueError):
                        break
                    family = font.getname()[0]
                    if family:
                        found.add(family.casefold())
    return found


def resolve_fonts(data):
    fonts = dict(STYLE['fonts'])
    overrides = data.get('fonts') or {}
    if not isinstance(overrides, dict) or any(k not in fonts for k in overrides):
        raise ValueError('fonts 只能包含 title、body、heading、secondary、footer')
    policy = data.get('font_policy', 'named')
    report = {'policy': policy, 'roles': {}}
    if policy == 'installed':
        have = installed_families()
        for role, chain in STYLE['font_fallbacks'].items():
            if role in overrides:
                continue
            pick = next((name for name in chain if name.casefold() in have), None)
            if pick:
                fonts[role] = pick
            report['roles'][role] = {'family': fonts[role], 'status': 'installed' if pick else 'not-found-kept-default'}
    elif policy != 'named':
        raise ValueError('font_policy 只能是 named 或 installed')
    for k, v in overrides.items():
        if not isinstance(v, str) or not v.strip():
            raise ValueError(f'fonts.{k} 必须是非空字体名')
        fonts[k] = v.strip()
    for role, fam in fonts.items():
        report['roles'].setdefault(role, {'family': fam, 'status': 'named'})
    return fonts, report


# ---------------- preflight ----------------
def validate_blocks(blocks, where):
    if not isinstance(blocks, list):
        raise ValueError(f'{where} 必须是数组')
    for i, b in enumerate(blocks):
        t = b.get('type') if isinstance(b, dict) else None
        if t in ('paragraph', 'heading'):
            v = b.get('text')
            if not isinstance(v, str) or not v.strip() or '\n' in v:
                raise ValueError(f'{where}[{i}].text 必须是非空单段文本')
            if t == 'heading' and b.get('level', 1) not in (1, 2, 3, 4):
                raise ValueError(f'{where}[{i}].level 必须是 1–4')
        elif t == 'table':
            cols, rows = b.get('columns'), b.get('rows')
            if not isinstance(cols, list) or not cols or not isinstance(rows, list):
                raise ValueError(f'{where}[{i}] 表格需要 columns 与 rows')
            for r in rows:
                if not isinstance(r, list) or len(r) != len(cols):
                    raise ValueError(f'{where}[{i}] 每行单元格数必须等于列数 {len(cols)}')
            ws = b.get('column_widths')
            if ws is not None and (not isinstance(ws, list) or len(ws) != len(cols)
                                   or not all(isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0 for v in ws)):
                raise ValueError(f'{where}[{i}] column_widths 必须是 {len(cols)} 个正数')
        elif t == 'list':
            if not isinstance(b.get('items'), list) or not b['items']:
                raise ValueError(f'{where}[{i}] 列表需要非空 items')
        else:
            raise ValueError(f'{where}[{i}] 类型只能是 paragraph / heading / table / list')


def choose_mode(data, fields):
    mode = data.get('mode', 'auto')
    if mode not in ('auto', 'redhead', 'plain'):
        raise ValueError('mode 只能是 auto、redhead 或 plain')
    if mode != 'auto':
        return mode
    kinds = set(STYLE['redhead_document_types'])
    kind = fields['document_type']
    if fields['redhead']:
        return 'redhead'
    if kind in ('报告', '意见'):
        # 工作报告、意见稿常是内部文档：没有发文机关或发文字号时按 plain 处理
        return 'redhead' if fields['document_number'] else 'plain'
    return 'redhead' if kind in kinds else 'plain'


def prepare(data):
    if not isinstance(data, dict):
        raise ValueError('输入必须是 JSON 对象')
    fields = {k: line_text(data, k) for k in list(REQUIRED) + list(OPTIONAL_LINES)}
    blocks = data.get('blocks', data.get('body', []))
    validate_blocks(blocks, 'blocks')
    mode = choose_mode(data, fields)
    missing, warnings = [], []
    for key in sorted(set(data) - KNOWN_KEYS):
        warnings.append(f'未识别字段 {key}，不会写入文档。')
    if not blocks:
        missing.append({'field': 'blocks', 'label': '正文'})
    for i, b in enumerate(blocks):
        if b['type'] in ('paragraph', 'heading') and PLACEHOLDER.search(b['text']):
            warnings.append(f'正文第 {i + 1} 段含待填写/待核实内容。')
    title_lines = data.get('title_lines')
    if title_lines is not None and (not isinstance(title_lines, list) or ''.join(title_lines) != fields['title']):
        raise ValueError('title_lines 拼接后必须与 title 完全一致')
    raw_att = data.get('attachments') or []
    if not isinstance(raw_att, list):
        raise ValueError('attachments 必须是数组')
    attachments = []
    for i, a in enumerate(raw_att):
        a = {'title': a} if isinstance(a, str) else a
        if not isinstance(a, dict) or not isinstance(a.get('title'), str) or not a['title'].strip():
            raise ValueError(f'附件 {i + 1} 需要 title')
        body = a.get('blocks', a.get('body', []))
        validate_blocks(body, f'attachments[{i}].blocks')
        if not body:
            warnings.append(f'附件 {i + 1} 没有正文，只写附件说明，不生成附件页。')
        attachments.append({'title': a['title'].strip(), 'blocks': body})
    if mode == 'plain':
        dropped = [k for k in ('redhead', 'document_number', 'signer', 'serial', 'security', 'urgency', 'format_mode', 'seal_mode', 'colophon')
                   if data.get(k)]
        if dropped:
            warnings.append('plain 模式不排红头要素，以下字段未写入：' + '、'.join(dropped) + '；需要红头请设 mode=redhead。')
        if pending(fields['title']):
            missing.append({'field': 'title', 'label': '标题'})
        return mode, fields, blocks, attachments, missing, warnings
    # ---- redhead checks (follow the original skill's preflight) ----
    for k, label in REQUIRED.items():
        if pending(fields[k]):
            missing.append({'field': k, 'label': label})
    fmt = data.get('format_mode', 'ordinary')
    if fmt not in ('ordinary', 'letter'):
        raise ValueError('format_mode 只能是 ordinary 或 letter')
    seal = data.get('seal_mode', 'none')
    if seal not in ('none', 'reserved'):
        raise ValueError('seal_mode 只能是 none 或 reserved')
    if 'seal_mode' not in data:
        warnings.append('未说明是否盖章，按不盖章排版；请确认。')
    if seal == 'reserved':
        warnings.append('只预留盖章位置，没有盖章；须由有权人员按规定办理。')
    kind, title = fields['document_type'], fields['title']
    if kind and title and kind != '其他' and not title.endswith(kind):
        warnings.append(f'标题没有以“{kind}”结尾，请核对标题与文种。')
    if kind in ('请示', '报告') and pending(fields['signer']):
        missing.append({'field': 'signer', 'label': '签发人'})
    if any(c in fields['signer'] for c in '、，,；;'):
        raise ValueError('只支持单个签发人')
    if fmt == 'letter' and fields['signer']:
        raise ValueError('信函格式不支持签发人')
    if fmt == 'ordinary' and kind == '函':
        warnings.append('当前为普通公文格式；需要信函格式请设 format_mode=letter。')
    date = fields['date']
    if date and not pending(date):
        m = re.fullmatch(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date)
        try:
            if not m:
                raise ValueError
            datetime.date(*map(int, m.groups()))
        except ValueError:
            warnings.append('成文日期无法自动核验，请确认。')
    if width(fields['recipient']) > 2 * 28 * 16:
        warnings.append('主送机关较长（超过约两行），须检查首页是否仍有正文；必要时人工把主送移到版记。')
    if len(fields['redhead']) > 20:
        warnings.append('发文机关标志较长，请检查压缩和换行。')
    col = data.get('colophon') or {}
    if not isinstance(col, dict) or any(k not in ('copy_to', 'issue_authority', 'issue_date') for k in col):
        raise ValueError('colophon 只支持 copy_to、issue_authority、issue_date')
    if any(not isinstance(v, str) or '\n' in v for v in col.values()):
        raise ValueError('colophon 各字段必须是单行文本')
    if fmt == 'letter' and (col.get('issue_authority') or col.get('issue_date')):
        raise ValueError('信函格式版记不编印发机关和印发日期')
    if bool(col.get('issue_authority')) != bool(col.get('issue_date')):
        missing.append({'field': 'colophon', 'label': '印发机关与印发日期需成对'})
    if len(col.get('copy_to', '')) > 24 or len(col.get('issue_authority', '')) > 14 or len(col.get('issue_date', '')) > 14:
        raise ValueError('版记字段过长（抄送≤24 字，印发机关、日期各≤14 字）')
    if col:
        warnings.append('版记固定在末页底部，请在 Word/WPS 中检查是否与正文重叠。')
    return mode, fields, blocks, attachments, missing, warnings


# ---------------- docx primitives ----------------
def run(p, text, font, size, color=None, scale=None, bold=False):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    rpr = r._element.get_or_add_rPr()
    rpr.rFonts.set(qn('w:eastAsia'), font)
    rpr.rFonts.set(qn('w:hAnsi'), font)
    if scale is not None:
        w = OxmlElement('w:w')
        w.set(qn('w:val'), str(scale))
        rpr.append(w)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def para(p, align=None, line=None, indent=False, before=0, after=0, keep=False):
    f = p.paragraph_format
    if align is not None:
        p.alignment = align
    f.space_before, f.space_after = Pt(before), Pt(after)
    f.line_spacing = Pt(line if line is not None else T['body_line'])
    f.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    if indent:
        f.first_line_indent = Pt(2 * T['body'])
    if keep:
        f.keep_with_next = True
    f.widow_control = True
    return p


def bottom_border(p, color, size, space):
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    for k, v in {'val': 'single', 'sz': str(size), 'space': str(space), 'color': color}.items():
        b.set(qn('w:' + k), v)
    bdr.append(b)
    ppr.append(bdr)


def page_number(p, font, even=False):
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if even else WD_ALIGN_PARAGRAPH.RIGHT
    if even:
        p.paragraph_format.left_indent = Pt(T['page_number'])
    else:
        p.paragraph_format.right_indent = Pt(T['page_number'])
    run(p, '— ', font, T['page_number'])
    r = run(p, '', font, T['page_number'])
    for tag, attrs, text in [('fldChar', {'fldCharType': 'begin'}, None), ('instrText', {}, ' PAGE '),
                             ('fldChar', {'fldCharType': 'separate'}, None), ('t', {}, '1'),
                             ('fldChar', {'fldCharType': 'end'}, None)]:
        el = OxmlElement('w:' + tag)
        for k, v in attrs.items():
            el.set(qn('w:' + k), v)
        if tag == 'instrText':
            el.set(qn('xml:space'), 'preserve')
        if text is not None:
            el.text = text
        r._r.append(el)
    run(p, ' —', font, T['page_number'])


def heading_font(level, fonts):
    return {1: fonts['heading'], 2: fonts['secondary'], 3: fonts['body'], 4: fonts['body']}[level]


def add_table(d, b, fonts):
    if b.get('title'):
        p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER, keep=True)
        run(p, b['title'], fonts['heading'], T['table_title'])
    cols, rows = b['columns'], b['rows']
    tbl = d.add_table(rows=len(rows) + 1, cols=len(cols))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tblpr = tbl._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side)
        for k, v in {'val': 'single', 'sz': str(STYLE['table']['border_eighth_points']), 'space': '0',
                     'color': STYLE['table']['border_color']}.items():
            e.set(qn('w:' + k), v)
        borders.append(e)
    tblpr.append(borders)
    total = Mm(STYLE['page_mm']['width'] - STYLE['page_mm']['left'] - STYLE['page_mm']['right'])
    ws = b.get('column_widths') or [1] * len(cols)
    grid = tbl._tbl.tblGrid.findall(qn('w:gridCol'))
    for i, col in enumerate(tbl.columns):
        w = int(total * ws[i] / sum(ws))
        grid[i].set(qn('w:w'), str(int(w / 635)))
        for c in col.cells:
            c.width = w
    for r_i, row in enumerate(tbl.rows):
        head = r_i == 0
        if head:
            trpr = row._tr.get_or_add_trPr()
            h = OxmlElement('w:tblHeader')
            h.set(qn('w:val'), 'true')
            trpr.append(h)
        row.height = Pt(T['table_line'] + 8)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        values = cols if head else rows[r_i - 1]
        for c_i, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.CENTER if head or c_i > 0 else WD_ALIGN_PARAGRAPH.LEFT
            para(p, align, line=T['table_line'])
            run(p, '' if values[c_i] is None else str(values[c_i]), fonts['heading'] if head else fonts['body'], T['table'])
    after = para(d.add_paragraph(), line=T['body_line'] // 2)
    if b.get('note'):
        run(after, b['note'], fonts['body'], T['table'])


def add_blocks(d, blocks, fonts):
    for b in blocks:
        t = b['type']
        if t == 'table':
            add_table(d, b, fonts)
            continue
        if t == 'list':
            for i, item in enumerate(b['items'], 1):
                p = para(d.add_paragraph(), indent=True)
                prefix = f'{i}. ' if b.get('numbered', True) and not NUMBERED.match(str(item)) else ''
                run(p, prefix + str(item).strip(), fonts['body'], T['body'])
            continue
        p = para(d.add_paragraph(), indent=True, keep=(t == 'heading'))
        font = heading_font(b.get('level', 1), fonts) if t == 'heading' else fonts['body']
        run(p, b['text'].strip(), font, T['body'])


def letter_rule(container, fonts, text=None):
    tbl = container.add_table(rows=1, cols=1, width=Mm(170))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    p = tbl.cell(0, 0).paragraphs[0]
    para(p, WD_ALIGN_PARAGRAPH.CENTER, line=48 if text else 4, before=5 if text else 0)
    if text:
        run(p, text, fonts['title'], 30, 'FF0000')
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    for k, v in {'val': 'thickThinSmallGap' if text else 'thinThickSmallGap', 'sz': '16',
                 'space': '5' if text else '0', 'color': 'FF0000'}.items():
        b.set(qn('w:' + k), v)
    bdr.append(b)
    ppr.append(bdr)


def cell_border(cell, side, size=8):
    tcpr = cell._tc.get_or_add_tcPr()
    bd = tcpr.find(qn('w:tcBorders'))
    if bd is None:
        bd = OxmlElement('w:tcBorders')
        tcpr.append(bd)
    e = OxmlElement('w:' + side)
    for k, v in {'val': 'single', 'sz': str(size), 'color': '000000'}.items():
        e.set(qn('w:' + k), v)
    bd.append(e)


def add_colophon(d, info, fonts, letter):
    copy, auth, date = (info.get(k, '').strip() for k in ('copy_to', 'issue_authority', 'issue_date'))
    if not (copy or auth or date):
        return
    rows = int(bool(copy)) + int(bool(auth or date))
    tbl = d.add_table(rows=rows, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.columns[0].width = Mm(156)
    fl = OxmlElement('w:tblpPr')
    for k, v in {'vertAnchor': 'page', 'horzAnchor': 'margin', 'tblpX': '0',
                 'tblpY': str(Mm(262).twips - rows * Pt(25).twips), 'leftFromText': '0', 'rightFromText': '0',
                 'topFromText': '40', 'bottomFromText': '0'}.items():
        fl.set(qn('w:' + k), v)
    tbl._tbl.tblPr.insert(0, fl)
    i = 0
    if copy:
        row = tbl.rows[i]
        row.height, row.height_rule = Pt(25), WD_ROW_HEIGHT_RULE.EXACTLY
        p = para(row.cells[0].paragraphs[0], line=20)
        p.paragraph_format.left_indent = p.paragraph_format.right_indent = Pt(14)
        run(p, '抄送：' + copy.rstrip('。') + '。', fonts['body'], 14)
        if not letter:
            cell_border(row.cells[0], 'top')
            cell_border(row.cells[0], 'bottom', 6 if rows > 1 else 8)
        i += 1
    if auth or date:
        row = tbl.rows[i]
        row.height, row.height_rule = Pt(25), WD_ROW_HEIGHT_RULE.EXACTLY
        p = para(row.cells[0].paragraphs[0], line=20)
        p.paragraph_format.left_indent = p.paragraph_format.right_indent = Pt(14)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(151), WD_TAB_ALIGNMENT.RIGHT)
        run(p, auth or '【待填写印发机关】', fonts['body'], 14)
        run(p, '\t' + (date or '【待填写印发日期】').removesuffix('印发') + '印发', fonts['body'], 14)
        if rows == 1:
            cell_border(row.cells[0], 'top')
        cell_border(row.cells[0], 'bottom')
    tail = para(d.add_paragraph(), line=1)
    run(tail, '', fonts['body'], 1)


def keep_last_with_next(d):
    """Keep the last body paragraph with the signature so it never sits alone on a page."""
    for p in reversed(d.paragraphs):
        if p.text.strip():
            p.paragraph_format.keep_with_next = True
            return


def signature(d, fields, fonts, reserved, shown):
    keep_last_with_next(d)
    sig, date = shown('signing_unit'), shown('date')
    sw, dw = width(sig), width(date)
    if reserved:
        p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.RIGHT, line=20, before=14, keep=True)
        p.paragraph_format.right_indent = Pt(64)
        run(p, '【待盖章位置】', fonts['heading'], 10, '666666')
        date_indent = 64
        sig_indent = max(0, date_indent + (dw - sw) / 2)
    else:
        sig_indent = max(32, 64 + dw - sw)
        date_indent = max(32, sig_indent + sw - 32 - dw)
    p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.RIGHT, before=0 if reserved else T['body_line'], keep=True)
    p.paragraph_format.right_indent = Pt(sig_indent)
    run(p, sig, fonts['body'], T['body'])
    p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.RIGHT)
    p.paragraph_format.right_indent = Pt(date_indent)
    run(p, date, fonts['body'], T['body'])


# ---------------- build ----------------
def setup(d, fonts, letter, label, first_page_number=True):
    s = d.sections[0]
    m = STYLE['page_mm']
    s.page_width, s.page_height = Mm(m['width']), Mm(m['height'])
    s.top_margin = Mm(70 if letter else m['top'])
    s.bottom_margin, s.left_margin, s.right_margin = Mm(m['bottom']), Mm(m['left']), Mm(m['right'])
    s.header_distance, s.footer_distance = Mm(m['header']), Mm(m['footer'])
    s.different_first_page_header_footer = True
    d.settings.odd_and_even_pages_header_footer = True
    if label:
        for h in (s.first_page_header, s.header, s.even_page_header):
            p = para(h.paragraphs[0], WD_ALIGN_PARAGRAPH.RIGHT, line=14)
            run(p, label, fonts['heading'], T['draft_label'], '666666')
    if first_page_number:
        page_number(s.first_page_footer.paragraphs[0], fonts['footer'])
    page_number(s.footer.paragraphs[0], fonts['footer'])
    page_number(s.even_page_footer.paragraphs[0], fonts['footer'], even=True)
    normal = d.styles['Normal']
    normal.font.name = fonts['body']
    normal.font.size = Pt(T['body'])
    normal._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fonts['body'])
    return s


def title_para(d, lines, fonts, after):
    p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER, line=T['title_line'], after=after, keep=True)
    for i, ln in enumerate(lines):
        if i:
            p.add_run().add_break()
        run(p, ln, fonts['title'], T['title'], '000000')
    return p


def build_plain(data, fields, blocks, fonts, out):
    d = Document()
    label = STYLE['draft_label'] if data.get('draft_label') is True else None
    setup(d, fonts, False, label)
    title_para(d, data.get('title_lines') or [fields['title'] or '【待填写标题】'], fonts,
               T['body_line'] if not (fields['subtitle'] or fields['author']) else 0)
    for key in ('subtitle', 'author'):
        if fields[key]:
            p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER, after=T['body_line'] if key == 'author' or not fields['author'] else 0)
            run(p, fields[key], fonts['secondary'], T['subtitle'])
    if fields['recipient']:
        p = para(d.add_paragraph(), keep=True)
        rec = fields['recipient']
        run(p, rec if rec.endswith(('：', ':')) else rec + '：', fonts['body'], T['body'])
    add_blocks(d, blocks, fonts)
    if fields['signing_unit'] or fields['date']:
        signature(d, fields, fonts, False, lambda k: fields[k])
    d.save(out)


def build_redhead(data, fields, blocks, attachments, fonts, out):
    d = Document()
    d.core_properties.subject = '红头文件草稿，待核定'
    letter = data.get('format_mode', 'ordinary') == 'letter'
    label = None if data.get('draft_label') is False else STYLE['draft_label']
    s = setup(d, fonts, letter, label, first_page_number=not letter)
    shown = lambda k: fields[k] if fields[k] else f'【待填写{REQUIRED[k]}】'
    if letter:
        letter_rule(s.first_page_header, fonts, shown('redhead'))
        letter_rule(s.first_page_footer, fonts)
    meta = 0
    for key in ('serial', 'security', 'urgency'):
        if fields[key]:
            run(para(d.add_paragraph(), keep=True), fields[key], fonts['heading'], T['body'])
            meta += 1
    if not letter:
        before = max(0, 98 - meta * T['body_line'])
        p = para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER, line=T['redhead_line'], before=before,
                 after=2 * T['body_line'] - 16, keep=True)
        red = shown('redhead')
        w = max(1, width(red) * T['redhead'] / 16)
        scale = max(20, min(100, int(Mm(156).pt / w * 100 * 0.93)))
        run(p, red, fonts['title'], T['redhead'], STYLE['red_line']['color'], scale)
    has_signer = bool(fields['signer']) or fields['document_type'] in ('请示', '报告')
    align = WD_ALIGN_PARAGRAPH.RIGHT if letter else (WD_ALIGN_PARAGRAPH.LEFT if has_signer else WD_ALIGN_PARAGRAPH.CENTER)
    p = para(d.add_paragraph(), align, after=2 * T['body_line'], keep=True)
    if has_signer and not letter:
        p.paragraph_format.tab_stops.add_tab_stop(Pt(16))
        p.paragraph_format.tab_stops.add_tab_stop(Mm(150), WD_TAB_ALIGNMENT.RIGHT)
        run(p, '\t' + shown('document_number') + '\t签发人：', fonts['body'], T['body'])
        run(p, fields['signer'] or '【待填写签发人】', fonts['secondary'], T['body'])
    else:
        run(p, shown('document_number'), fonts['body'], T['body'])
    if not letter:
        bottom_border(p, STYLE['red_line']['color'], STYLE['red_line']['eighth_points'], 11)
    title_para(d, data.get('title_lines') or [shown('title')], fonts, T['body_line'])
    p = para(d.add_paragraph(), keep=True)
    rec = shown('recipient')
    run(p, rec if rec.endswith(('：', ':')) else rec + '：', fonts['body'], T['body'])
    add_blocks(d, blocks or [{'type': 'paragraph', 'text': '【待填写正文】'}], fonts)
    n = len(attachments)
    for i, a in enumerate(attachments):
        p = para(d.add_paragraph(), before=T['body_line'] if i == 0 else 0, keep=True)
        start = max(112, 88 + width(f'{n}.')) if n > 1 else 80
        p.paragraph_format.left_indent = Pt(start)
        p.paragraph_format.first_line_indent = Pt((32 if i == 0 else 80) - start)
        p.paragraph_format.tab_stops.add_tab_stop(Pt(start))
        run(p, ('附件：' if i == 0 else '') + (f'{i + 1}.' if n > 1 else '') + '\t' + a['title'], fonts['body'], T['body'])
    signature(d, fields, fonts, data.get('seal_mode', 'none') == 'reserved', shown)
    for i, a in enumerate(attachments):
        if not a['blocks']:
            continue
        d.add_page_break()
        run(para(d.add_paragraph(), after=T['body_line'], keep=True), '附件' + (str(i + 1) if n > 1 else ''),
            fonts['heading'], T['body'])
        run(para(d.add_paragraph(), WD_ALIGN_PARAGRAPH.CENTER, line=T['title_line'], after=T['body_line'], keep=True),
            a['title'], fonts['title'], T['title'])
        add_blocks(d, a['blocks'], fonts)
    add_colophon(d, data.get('colophon') or {}, fonts, letter)
    d.save(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--input', required=True, type=Path)
    ap.add_argument('--out', type=Path)
    ap.add_argument('--check-only', action='store_true')
    ap.add_argument('--font-policy', choices=['named', 'installed'])
    ap.add_argument('--overwrite', action='store_true')
    a = ap.parse_args()
    if not a.check_only:
        if not a.out or a.out.suffix.lower() != '.docx':
            ap.error('生成时 --out 必须是 .docx')
        if a.out.exists() and not a.overwrite:
            ap.error(f'{a.out} 已存在；换个文件名或加 --overwrite')
    try:
        data = json.loads(a.input.read_text(encoding='utf-8'))
        if a.font_policy:
            data['font_policy'] = a.font_policy
        mode, fields, blocks, attachments, missing, warnings = prepare(data)
        fonts, font_report = resolve_fonts(data)
        report = {'style': STYLE['name'], 'version': STYLE['version'], 'mode': mode, 'missing': missing,
                  'warnings': warnings, 'fonts': font_report,
                  'note': '字体只写名字、不嵌入文件；打开文档的电脑没装这些字体时会被替换，分页可能变化。'}
        if mode == 'plain' and attachments:
            warnings.append('plain 模式不排附件页；附件请写进正文或改用 redhead。')
        if not a.check_only:
            a.out.parent.mkdir(parents=True, exist_ok=True)
            if mode == 'plain':
                build_plain(data, fields, blocks, fonts, a.out)
            else:
                build_redhead(data, fields, blocks, attachments, fonts, a.out)
            Document(str(a.out))
            report['output'] = str(a.out.resolve())
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, json.JSONDecodeError, KeyError, TypeError, AttributeError) as e:
        print(f'处理失败：{type(e).__name__}: {e}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
