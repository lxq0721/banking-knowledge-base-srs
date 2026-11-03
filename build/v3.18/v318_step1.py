# -*- coding: utf-8 -*-
"""v3.18 step1：资料失效的知识层处置（137 → 138）
1) 新增 FR-GEN-33「资料失效后的知识层处置」（必须）+ 表 6-9 三情形表
2) FR-IMP-16 收窄为入口动作，处置范围指向 FR-GEN-33（避免同一规则两处写）
3) 表 4-5 加「疑似同一份」行，题注改四种情形；FR-IMP-10 判定依据补文件名主干
4) FR-GEN-32 汇总清单补「项目知识总览」并注明按条目刷新
5) 表 7-4「已废止」的进入条件补资料失效
6) 术语表加「生成单元」
7) 表 A-3 加「跨应用系统的资料改动」
"""
import sys, io
from copy import deepcopy
import docx
from docx.text.paragraph import Paragraph
from docx.table import _Row

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = '_v318_base.docx'
DST = '_v318_step1.docx'
d = docx.Document(SRC)
ps = list(d.paragraphs)


def find_para(prefix, style=None):
    for p in ps:
        if p.style.name.startswith('toc'):
            continue
        if style and p.style.name != style:
            continue
        if p.text.strip().startswith(prefix):
            return p
    raise KeyError(prefix)


def find_tbl(heads):
    for tb in d.tables:
        got = [c.text.strip() for c in tb.rows[0].cells]
        if got[:len(heads)] == heads:
            return tb
    raise KeyError(str(heads))


def row_exact(tb, key, col=0):
    for r in tb.rows:
        if r.cells[col].text.strip() == key:
            return r
    raise KeyError(key)


def fill_cell(cell, text):
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p0 = cell.paragraphs[0]
    assert p0.runs, cell.text[:30]
    p0.runs[0].text = text
    for r in p0.runs[1:]:
        r._element.getparent().remove(r._element)


def set_para(p, head, body, keep_next):
    """保留首个 run 的格式：runs[0]=前缀，runs[1]=正文，其余删除"""
    p.runs[0].text = head
    p.runs[1].text = body
    for r in p.runs[2:]:
        r._element.getparent().remove(r._element)
    p.paragraph_format.keep_with_next = keep_next
    p.paragraph_format.keep_together = True
    return p


def rewrite_desc(prefix_para, new_text):
    """描述段整段改写（保留『描述：』ru n 的格式）"""
    p = prefix_para
    assert p.text.strip().startswith('描述：'), p.text[:20]
    p.runs[0].text = '描述：'
    p.runs[1].text = new_text
    for r in p.runs[2:]:
        r._element.getparent().remove(r._element)
    return p


# ============ 1. 新增 FR-GEN-33 ============
title_tpl = find_para('FR-GEN-32\u3000')
assert title_tpl.text.strip() == 'FR-GEN-32\u3000资料变化后的重算范围', title_tpl.text
pri_tpl = ps[ps.index(title_tpl) + 1]
cons_tpl = ps[ps.index(title_tpl) + 2]
desc_tpl = ps[ps.index(title_tpl) + 3]
assert pri_tpl.text.strip() == '优先级：必须', pri_tpl.text
assert cons_tpl.text.strip().startswith('约束：'), cons_tpl.text[:20]
assert desc_tpl.text.strip().startswith('描述：'), desc_tpl.text[:20]

CODE = 'FR-GEN-33'
TITLE = '资料失效后的知识层处置'
CONS = ('负向需求：以「撤销批次后知识文档仍可被检索」「资料停用后其切片仍进入回答依据」'
        '「整包重传未包含的程序仍按在用版本作答」为反例验收。')
DESC = ('资料被撤销、停用或不再提交后，由它生成的产物应同步失效：知识文档转为「已废止」，'
        '其切片与向量退出检索，图谱节点与关系按 FR-STO-27 下线，经它确认的经验条目一并失效；'
        '该资料参与构成的跨文件汇总产物（数据实体说明、业务术语表、系统结构说明、项目知识总览）'
        '随之重算。原件的只读快照保留，指向该资料的引用改为占位并标「资料已失效」。'
        '三种触发情形的判定与处置见表 6-9。')

t_el = deepcopy(title_tpl._p)
tp = Paragraph(t_el, d)
assert len(tp.runs) >= 2
tp.runs[0].text = CODE
tp.runs[1].text = '\u3000' + TITLE
for r in tp.runs[2:]:
    r._element.getparent().remove(r._element)
tp.paragraph_format.keep_with_next = True
tp.paragraph_format.keep_together = True

p_el = deepcopy(pri_tpl._p)
pp = Paragraph(p_el, d)
pp.paragraph_format.keep_with_next = True
pp.paragraph_format.keep_together = True

c_el = deepcopy(cons_tpl._p)
cp = Paragraph(c_el, d)
set_para(cp, '约束：', CONS, True)

d_el = deepcopy(desc_tpl._p)
dp = Paragraph(d_el, d)
set_para(dp, '描述：', DESC, True)   # 后面紧跟表题注，故 keep_with_next

# ============ 2. 表 6-9 ============
cap_tpl = find_para('表 6-8\u3000', style='图表题注')
assert cap_tpl.text.strip() == '表 6-8\u3000来源不一致的三种情形与处置', cap_tpl.text
cap_el = deepcopy(cap_tpl._p)
cap = Paragraph(cap_el, d)
assert len(cap.runs) >= 3, [r.text for r in cap.runs]
cap.runs[0].text = '表'
cap.runs[1].text = ' 6-9'
cap.runs[2].text = '\u3000资料失效的三种情形与处置'
for r in cap.runs[3:]:
    r._element.getparent().remove(r._element)
cap.paragraph_format.keep_with_next = True
cap.paragraph_format.keep_together = True

t68 = find_tbl(['情形', '判定', '处置', '计入疑点'])
new_tbl = deepcopy(t68._tbl)
assert len(t68.rows) == 4, len(t68.rows)
from docx.table import Table
nt = Table(new_tbl, d)
rows = [
    ['情形', '判定', '处置', '可恢复方式'],
    ['整批撤销',
     '操作人对已完成的导入批次执行撤销（FR-IMP-16）',
     '该批次全部资料的知识层产物失效，原件快照保留',
     '重新提交该批资料'],
    ['单份停用',
     '操作人在导入清单对单份资料执行停用',
     '仅该资料的产物失效，同批其他资料不受影响',
     '解除停用后按 FR-GEN-32 重算'],
    ['整包重传未出现',
     '操作人声明本次为整包重传，与该范围上一版比对，存在上一版有而本次无的资料',
     '按单份停用处置，导入清单标「本次未提交」',
     '下次提交时补齐该资料'],
]
for tr, vals in zip(nt.rows, rows):
    for c, v in zip(tr.cells, vals):
        fill_cell(c, v)
print('表 6-9 →', len(nt.rows), '行 |', [c.text for c in nt.rows[3].cells])

# 插入顺序：GEN-32 描述段之后 → GEN-33 四段 → 表 6-9 题注 → 表
anchor = desc_tpl._p
for el in (t_el, p_el, c_el, d_el, cap_el):
    anchor.addnext(el)
    anchor = el
anchor.addnext(new_tbl)      # 表格跟题注，已在文档内
print('GEN-33 段落与表 6-9 已插入')

# ============ 3. FR-IMP-16 收窄 ============
imp16 = find_para('FR-IMP-16\u3000')
i16 = ps.index(imp16)
assert ps[i16 + 1].text.strip() == '优先级：应该'
rewrite_desc(ps[i16 + 2],
             '系统应支持按批次撤销已完成的导入；撤销后该批次资料按 FR-GEN-33 处置，原件快照保留。')
print('FR-IMP-16 描述已收窄')

# ============ 4. 表 4-5 加「疑似同一份」+ FR-IMP-10 ============
imp10 = find_para('FR-IMP-10\u3000')
i10 = ps.index(imp10)
rewrite_desc(ps[i10 + 2],
             '系统应在导入前计算每份资料的内容指纹，并把本次提交的资料按路径、内容指纹与文件名主干'
             '归入表 4-5 的一种情形；只对「新增」、「同路径变更」以及经人指定接续的「疑似同一份」'
             '资料执行后续环节。')
print('FR-IMP-10 描述已改')

t45 = find_tbl(['情形', '判定', '处置'])
assert len(t45.rows) == 4, len(t45.rows)
last = t45.rows[-1]
tr_el = deepcopy(last._tr)
last._tr.addnext(tr_el)
nr = _Row(tr_el, t45)
vals = ['疑似同一份',
        '路径与内容指纹均未命中，但文件名主干（去掉路径、扩展名与版本号后缀）与已有记录相同',
        '预检中列为候选，由人指定是否接续：接续的按「同路径变更」，否则按「新增」']
for c, v in zip(nr.cells, vals):
    fill_cell(c, v)
print('表 4-5 →', len(t45.rows), '行')

cap45 = find_para('表 4-5\u3000', style='图表题注')
assert cap45.text.strip() == '表 4-5\u3000再次提交的三种情形与处置', cap45.text
assert cap45.runs[-1].text.endswith('再次提交的三种情形与处置'), [r.text for r in cap45.runs]
cap45.runs[-1].text = '\u3000再次提交的四种情形与处置'
print('表 4-5 题注 →', cap45.text.strip())

# ============ 5. FR-GEN-32 补项目知识总览 ============
rewrite_desc(desc_tpl,
             '已导入资料新增或变更后，系统只重算直接引用该资料的单元与跨文件聚合的汇总产物'
             '（数据实体说明、业务术语表、系统结构说明、项目知识总览），其中项目知识总览按条目刷新、'
             '不整份重生成；未引用它的单元不得重算，也不得因补一份资料触发整个应用系统重新生成。'
             '重算按 FR-GEN-05 另出新版本，已修正内容按 FR-REV-24 保留，'
             '范围外对象只更新占位引用，见 FR-GEN-03。')
print('FR-GEN-32 描述已补项目知识总览')

# ============ 6. 表 7-4 已废止进入条件 ============
t74 = find_tbl(['状态', '进入条件', '可执行动作', '可用方式'])
r74 = row_exact(t74, '已废止')
old = r74.cells[1].text.strip()
assert old == '源文件变更、内容被替代或人工废止', old
fill_cell(r74.cells[1], '源文件变更或资料失效（FR-GEN-33）、内容被替代或人工废止')
print('表 7-4 已废止 →', r74.cells[1].text.strip())

# ============ 7. 术语表加「生成单元」 ============
tt = find_tbl(['术语', '含义'])
src_row = row_exact(tt, '知识文档')
tr_el = deepcopy(src_row._tr)
src_row._tr.addnext(tr_el)
nr = _Row(tr_el, tt)
fill_cell(nr.cells[0], '生成单元')
fill_cell(nr.cells[1], '知识生成与重算的基本对象：一份程序、一份文档或一张工作表各对应一份产物；'
                       '一次生成的范围由若干生成单元构成，见 FR-GEN-02 与 FR-GEN-32。')
print('术语表 →', len(tt.rows), '行')

# ============ 8. 表 A-3 加跨应用系统 ============
ta3 = find_tbl(['场景', '现状', '建议的处置方向'])
lastrow = ta3.rows[-1]
tr_el = deepcopy(lastrow._tr)
lastrow._tr.addnext(tr_el)
nr = _Row(tr_el, ta3)
fill_cell(nr.cells[0], '跨应用系统的资料改动')
fill_cell(nr.cells[1], '直接引用只在应用系统内闭合，甲系统改动接口字段时，乙系统的对应知识是否重算未定')
fill_cell(nr.cells[2], '按应用系统边界处理，改动侧登记疑点；待委托方确认系统边界后另行确定')
print('表 A-3 →', len(ta3.rows), '行')

d.save(DST)
print('saved', DST)

# ============ 校验 ============
d2 = docx.Document(DST)
P = list(d2.paragraphs)
i = [k for k, p in enumerate(P) if p.text.strip().startswith('FR-GEN-33\u3000')][0]
for j in range(i - 1, i + 7):
    print(j, '|', P[j].style.name, '|', repr(P[j].text[:70]))
