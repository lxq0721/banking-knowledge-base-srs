# -*- coding: utf-8 -*-
"""把测试方案 markdown 转成 docx，样式对齐《核心系统迁移知识库需求规格书》。"""
import re, sys, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = sys.argv[1] if len(sys.argv) > 1 else '核心系统迁移知识库测试方案_v0.8.md'
OUT = sys.argv[2] if len(sys.argv) > 2 else '核心系统迁移知识库测试方案_v0.8.docx'
VERSION = 'v0.8'
DOC_TITLE = '核心系统迁移知识库'
DOC_SUB = '测试方案'
HEADER_TEXT = '核心系统迁移知识库　测试方案'

INK = RGBColor(0x26, 0x26, 0x26)
GREY = RGBColor(0x59, 0x59, 0x59)
ACCENT = RGBColor(0x1F, 0x3B, 0x5C)

BODY_PT = 10.5
TBL_PT = 9
LATIN = 'Times New Roman'
EA = '宋体'
EA_H = '微软雅黑'


def set_font(run, size=None, bold=None, color=None, latin=LATIN, ea=EA, italic=None):
    run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    rFonts.set(qn('w:eastAsia'), ea)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color is not None: run.font.color.rgb = color


def para_fmt(p, before=0, after=0, line=1.5, first=0, left=0, keep_next=False, align=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.first_line_indent = Cm(first)
    pf.left_indent = Cm(left)
    pf.keep_with_next = keep_next
    if align is not None: pf.alignment = align


def add_left_bar(p, color='BFBFBF'):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single'); left.set(qn('w:sz'), '12')
    left.set(qn('w:space'), '6'); left.set(qn('w:color'), color)
    pBdr.append(left)
    pPr.append(pBdr)


TOKEN = re.compile(r'(\*\*.+?\*\*|`[^`]+`)')


def add_inline(p, text, size=BODY_PT, color=INK, bold=False, ea=EA):
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            r = p.add_run(part[2:-2]); set_font(r, size, True, color, ea=ea)
        elif part.startswith('`') and part.endswith('`'):
            r = p.add_run(part[1:-1]); set_font(r, size - 0.5, bold, color, latin='Consolas', ea='Consolas')
        else:
            r = p.add_run(part); set_font(r, size, bold, color, ea=ea)


def strip_md_cell(s):
    return s.strip()


def cell_text_len(s):
    s = re.sub(r'\*\*|`', '', s)
    n = 0
    for ch in s:
        n += 2 if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f' or '\uff00' <= ch <= '\uffef' else 1
    return n


def set_cell_bg(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def set_table_borders(table, color='9E9E9E', sz='4'):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), sz)
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
        borders.append(el)
    tblPr.append(borders)
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)


def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    for tag in ('w:tblHeader', 'w:cantSplit'):
        el = OxmlElement(tag); el.set(qn('w:val'), 'true'); trPr.append(el)


def set_row_cantsplit(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:cantSplit'); el.set(qn('w:val'), 'true'); trPr.append(el)


def add_table(doc, rows, avail_cm):
    ncol = max(len(r) for r in rows)
    rows = [r + [''] * (ncol - len(r)) for r in rows]
    lens = []
    for i in range(ncol):
        lens.append(max(cell_text_len(r[i]) for r in rows))
    lens = [max(3, min(v, 46)) for v in lens]
    total = sum(lens)
    widths = [avail_cm * v / total for v in lens]
    # 最小宽度保护
    minw = [1.7 if ncol >= 6 else (2.0 if ncol == 5 else 2.8)] * ncol
    for i in range(ncol):
        if widths[i] < minw[i]:
            widths[i] = minw[i]
    scale = avail_cm / sum(widths)
    widths = [w * scale for w in widths]

    table = doc.add_table(rows=len(rows), cols=ncol)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    for i, w in enumerate(widths):
        table.columns[i].width = Cm(w)
    for ri, row in enumerate(rows):
        tr = table.rows[ri]
        set_repeat_header(tr) if ri == 0 else set_row_cantsplit(tr)
        for ci, val in enumerate(row):
            cell = tr.cells[ci]
            cell.width = Cm(widths[ci])
            p = cell.paragraphs[0]
            para_fmt(p, before=1.5, after=1.5, line=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
            txt = strip_md_cell(val)
            if ri == 0:
                add_inline(p, txt, TBL_PT, INK, bold=True)
                set_cell_bg(cell, 'EDEFF2')
            else:
                add_inline(p, txt, TBL_PT, INK)
    # 表后留白
    sp = doc.add_paragraph()
    para_fmt(sp, before=0, after=4, line=1.0)
    sp.runs and None
    return table


def add_field(p, instr, size=9, color=GREY):
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instr)
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    szel = OxmlElement('w:sz'); szel.set(qn('w:val'), str(int(size * 2))); rPr.append(szel)
    t = OxmlElement('w:t'); t.text = '1'
    r.append(rPr); r.append(t); fld.append(r)
    p._p.append(fld)


def main():
    lines = open(SRC, encoding='utf-8').read().split('\n')
    doc = Document()

    # 页面
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.54)
    s.top_margin = s.bottom_margin = Cm(2.29)
    s.header_distance = Cm(1.3)
    s.footer_distance = Cm(1.3)

    # 默认样式
    st = doc.styles['Normal']
    st.font.name = LATIN
    st.font.size = Pt(BODY_PT)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), EA)
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)

    avail = 21 - 2.54 * 2  # 15.92cm

    # ---------- 封面 ----------
    p = doc.add_paragraph(); para_fmt(p, before=110, after=6)
    r = p.add_run(DOC_TITLE); set_font(r, 13, True, INK, ea=EA_H)
    p = doc.add_paragraph(); para_fmt(p, before=0, after=10)
    r = p.add_run(DOC_SUB); set_font(r, 30, True, ACCENT, ea=EA_H)
    p = doc.add_paragraph(); para_fmt(p, before=0, after=2)
    r = p.add_run(VERSION); set_font(r, 11, False, GREY, ea=EA_H)

    # 封面信息表
    cover = [
        ['文档名称', '核心系统迁移知识库项目 测试方案'],
        ['版本', VERSION],
        ['状态', '待评审'],
        ['编制', '测试组'],
        ['需评审方', '产品 / 开发方 / 委托方业务与项目管理部门'],
        ['验收依据', '《核心系统迁移知识库需求规格书》v3.18 及配套验收标准清单'],
        ['实现参考', '委托方技术方案 V1.0（只影响取数与用例准备，不作判定依据）'],
    ]
    for i in (0, 1, 2, 3, 4, 5, 6):
        pass
    gap = doc.add_paragraph(); para_fmt(gap, before=0, after=0, line=1.0)
    gap.add_run('')
    # 手动构建两列信息表
    tb = doc.add_table(rows=len(cover), cols=2)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    set_table_borders(tb, color='D0D0D0')
    w0, w1 = 3.6, avail - 3.6
    for ri, (k, v) in enumerate(cover):
        tr = tb.rows[ri]; set_row_cantsplit(tr)
        for ci, val in enumerate((k, v)):
            cell = tr.cells[ci]
            cell.width = Cm(w0 if ci == 0 else w1)
            pp = cell.paragraphs[0]
            para_fmt(pp, before=2, after=2, line=1.2)
            add_inline(pp, val, 9.5, INK, bold=(ci == 0))
            if ci == 0:
                set_cell_bg(cell, 'F2F3F5')
    for i, w in enumerate((w0, w1)):
        tb.columns[i].width = Cm(w)

    doc.add_page_break()

    # ---------- 目录 ----------
    p = doc.add_paragraph(); para_fmt(p, before=0, after=10)
    r = p.add_run('目　录'); set_font(r, 15, True, INK, ea=EA_H)
    toc_p = doc.add_paragraph()
    para_fmt(toc_p, before=0, after=0, line=1.3)
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), r'TOC \o "1-2" \h \z \u')
    run = OxmlElement('w:r'); t = OxmlElement('w:t')
    t.text = '（目录将在打开文档后自动更新：全选后按 F9，或在 Word 中右键"更新域"）'
    run.append(t); fld.append(run)
    toc_p._p.append(fld)

    doc.add_page_break()

    # ---------- 正文 ----------
    i = 0
    first_h1 = True
    while i < len(lines):
        ln = lines[i]
        raw = ln.rstrip()
        if not raw.strip():
            i += 1; continue
        if raw.startswith('# '):           # 已由封面承接
            i += 1; continue
        if raw.startswith('## '):
            txt = raw[3:].strip()
            p = doc.add_paragraph(style='Heading 1')
            para_fmt(p, before=14 if not first_h1 else 0, after=7, line=1.3, keep_next=True)
            r = p.add_run(txt); set_font(r, 16, True, ACCENT, ea=EA_H)
            first_h1 = False
            # 章前细分隔线
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr'); bot = OxmlElement('w:bottom')
            bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6')
            bot.set(qn('w:space'), '4'); bot.set(qn('w:color'), 'C9CDD4')
            pBdr.append(bot); pPr.append(pBdr)
            i += 1; continue
        if raw.startswith('### '):
            txt = raw[4:].strip()
            p = doc.add_paragraph(style='Heading 2')
            para_fmt(p, before=11, after=5, line=1.3, keep_next=True)
            r = p.add_run(txt); set_font(r, 12.5, True, INK, ea=EA_H)
            i += 1; continue
        if raw.startswith('> '):
            buf = []
            while i < len(lines) and lines[i].startswith('> '):
                buf.append(lines[i][2:].strip()); i += 1
            for b in buf:
                p = doc.add_paragraph()
                para_fmt(p, before=2, after=2, line=1.4, left=0.5)
                add_left_bar(p)
                add_inline(p, b, 9.5, GREY)
            continue
        if raw.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].lstrip().startswith('|'):
                row = lines[i].strip()
                if re.match(r'^\|[\s\-\|:]+\|$', row):
                    i += 1; continue
                cells = [c for c in row.strip('|').split('|')]
                rows.append([c.strip() for c in cells])
                i += 1
            add_table(doc, rows, avail)
            continue
        if raw.strip() == '---':
            i += 1; continue
        if raw.startswith('- '):
            while i < len(lines) and lines[i].startswith('- '):
                txt = lines[i][2:].strip(); i += 1
                p = doc.add_paragraph()
                para_fmt(p, before=1, after=1, line=1.45, left=0.75, first=-0.45)
                r = p.add_run('· '); set_font(r, BODY_PT, False, GREY)
                add_inline(p, txt)
            continue
        m = re.match(r'^(\d+)\. (.+)$', raw)
        if m:
            while i < len(lines):
                m2 = re.match(r'^(\d+)\. (.+)$', lines[i].rstrip())
                if not m2:
                    break
                num, txt = m2.group(1), m2.group(2).strip()
                i += 1
                p = doc.add_paragraph()
                para_fmt(p, before=1, after=1, line=1.45, left=0.95, first=-0.55)
                r = p.add_run(num + '. '); set_font(r, BODY_PT, True, ACCENT)
                add_inline(p, txt)
            continue
        if raw.startswith('*') and raw.endswith('*') and not raw.startswith('**'):
            p = doc.add_paragraph()
            para_fmt(p, before=8, after=0, line=1.4)
            r = p.add_run(raw.strip('*')); set_font(r, 9, False, GREY, italic=True)
            i += 1; continue
        # 普通段落
        p = doc.add_paragraph()
        para_fmt(p, before=2, after=2, line=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        add_inline(p, raw.strip())
        i += 1

    # ---------- 页眉页脚 ----------
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]
    para_fmt(hp, before=0, after=0, line=1.0)
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = hp.add_run(HEADER_TEXT); set_font(r, 8.5, False, GREY)
    r = hp.add_run('　　' + VERSION); set_font(r, 8.5, False, GREY)
    pPr = hp._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr'); bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '4')
    bot.set(qn('w:space'), '4'); bot.set(qn('w:color'), 'BFBFBF')
    pBdr.append(bot); pPr.append(pBdr)

    fp = sec.footer.paragraphs[0]
    para_fmt(fp, before=0, after=0, line=1.0)
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run('第 '); set_font(r, 9, False, GREY)
    add_field(fp, 'PAGE')
    r = fp.add_run(' 页　共 '); set_font(r, 9, False, GREY)
    add_field(fp, 'NUMPAGES')
    r = fp.add_run(' 页'); set_font(r, 9, False, GREY)

    cp = doc.core_properties
    cp.title = '核心系统迁移知识库 测试方案'
    cp.subject = '测试方案 ' + VERSION
    cp.author = '测试组'
    cp.last_modified_by = '测试组'
    cp.comments = ''
    cp.category = ''
    cp.keywords = ''

    doc.save(OUT)
    print('saved', OUT)


if __name__ == '__main__':
    main()
