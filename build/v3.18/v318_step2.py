# -*- coding: utf-8 -*-
"""v3.18 step2：统计口径同步 + 附录 C + 版本号"""
import sys, io
from copy import deepcopy
import docx
from docx.text.paragraph import Paragraph

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = '_v318_step1.docx'
DST = '_v318_step2.docx'
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


def fill_cell(cell, text):
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p0 = cell.paragraphs[0]
    assert p0.runs, cell.text[:30]
    p0.runs[0].text = text
    for r in p0.runs[1:]:
        r._element.getparent().remove(r._element)


# ---------- 1. 3.1 概述条数 ----------
p31 = find_para('七个功能域共')
assert '130 条' in p31.text and '137 条' in p31.text, p31.text
new31 = p31.text.replace('共 130 条', '共 131 条').replace('合计 137 条', '合计 138 条')
assert new31 != p31.text
p31.runs[0].text = new31
for r in p31.runs[1:]:
    r._element.getparent().remove(r._element)
print('3.1 →', p31.text.strip()[:80])

# ---------- 2. 表 1-3 必须条数 ----------
t13 = find_tbl(['优先级', '定义', '本文中条目', '裁剪影响'])
for r in t13.rows:
    if r.cells[0].text.strip() == '必须':
        assert r.cells[2].text.strip() == '共 104 条', r.cells[2].text
        fill_cell(r.cells[2], '共 105 条')
        print('表 1-3 必须 →', r.cells[2].text.strip())

# ---------- 3. 表 A-1 P1 ----------
ta1 = find_tbl(['批次', '定位与取舍原则', '覆盖范围'])
for r in ta1.rows:
    if r.cells[0].text.strip().startswith('P1'):
        t = r.cells[2].text.strip()
        assert t.startswith('101 条：104 条'), t
        new = (t.replace('101 条', '102 条').replace('104 条「必须」', '105 条「必须」'))
        fill_cell(r.cells[2], new)
        print('表 A-1 P1 →', new[:60])

# ---------- 4. 表 3-2 / 表 B-1 的 GEN 触发与范围组 ----------
t32 = find_tbl(['功能域', '功能组', '需求编号', '条数'])
for r in t32.rows:
    if r.cells[1].text.strip() == '触发与范围':
        assert r.cells[2].text.strip() == 'FR-GEN-01~05、32', r.cells[2].text
        fill_cell(r.cells[2], 'FR-GEN-01~05、32~33')
        fill_cell(r.cells[3], '7')
        print('表 3-2 →', r.cells[2].text.strip(), r.cells[3].text.strip())

tb1 = find_tbl(['功能域', '章节', '功能组', '需求编号', '条数', '原型屏'])
for r in tb1.rows:
    if r.cells[2].text.strip() == '触发与范围':
        assert r.cells[3].text.strip() == 'FR-GEN-01~05、32', r.cells[3].text
        fill_cell(r.cells[3], 'FR-GEN-01~05、32~33')
        fill_cell(r.cells[4], '7')
        print('表 B-1 →', r.cells[3].text.strip(), r.cells[4].text.strip())

# ---------- 5. 表 B-1 注 ----------
pnote = find_para('注：合计 137 条需求')
assert '133 条功能需求' in pnote.text and '130 条' in pnote.text
new = (pnote.text.replace('合计 137 条', '合计 138 条')
       .replace('133 条功能需求', '134 条功能需求')
       .replace('七个功能域 130 条', '七个功能域 131 条'))
pnote.runs[0].text = new
for r in pnote.runs[1:]:
    r._element.getparent().remove(r._element)
print('表 B-1 注 →', pnote.text.strip()[:70])

# ---------- 5b. 表 A-3 前的引导段：四个场景 → 五个 ----------
plead = find_para('下列四个场景')
assert plead.text.strip().startswith('下列四个场景'), plead.text[:20]
new = plead.text.replace('下列四个场景', '下列五个场景')
plead.runs[0].text = new
for r in plead.runs[1:]:
    r._element.getparent().remove(r._element)
print('表 A-3 引导段 →', plead.text.strip()[:40])

# ---------- 6. 附录 C 插入 v3.18 ----------
anchor = find_para('v3.17 修订')
assert anchor is not None
tpl_head = anchor
tpl_item = ps[ps.index(anchor) + 1]
assert tpl_item.text.strip().startswith('·'), tpl_item.text[:20]

HEAD = 'v3.18 修订\u3000补资料失效的处置规则，封住「资料撤了但知识还在」的口子：'
ITEMS = [
    '· 新增 FR-GEN-33「资料失效后的知识层处置」（必须）：资料被撤销、停用或不再提交后，'
    '其知识文档转为「已废止」，切片与向量退出检索，图谱节点与关系按 FR-STO-27 下线，'
    '经它确认的经验条目一并失效，该资料参与构成的跨文件汇总产物随之重算；原件快照保留，'
    '指向它的引用改为占位并标「资料已失效」。三种触发情形（整批撤销、单份停用、整包重传未出现）'
    '的判定与处置见表 6-9。原有规则只在 FR-IMP-16 规定了撤销批次时清除「导入记录、解析产物与索引条目」，'
    '知识文档、图谱节点与经验条目都不在其内，撤销后仍可被检索作答；单份停用与整包重传少交文件'
    '两种情形则没有任何入口，被删的程序会按在用版本继续作答。',

    '· FR-IMP-16 的描述收窄为入口动作，处置范围改指 FR-GEN-33；表 6-9 的「整包重传未出现」一项以'
    '操作人声明整包范围为判定前提，与 FR-IMP-21「系统不自行扫描目录、无法判断一批是否到齐」保持一致。',

    '· 表 4-5 由三种情形扩为四种，补「疑似同一份」：路径与内容指纹均未命中、但文件名主干与已有记录相同时，'
    '在预检中列为候选，由人指定是否接续。原判定只看路径与指纹，改名的同时改了内容、或换格式重新提交时'
    '会被判为「新增」，旧记录成为孤儿、同一份资料在库中记两遍。FR-IMP-10 的判定依据同步补入文件名主干。',

    '· FR-GEN-32 的汇总产物清单补「项目知识总览」（FR-GEN-15），并注明其按条目刷新、不整份重生成；'
    '表 7-4「已废止」的进入条件补「资料失效（FR-GEN-33）」，此前只写了源文件变更，全篇没有能触发它的动作。',

    '· 术语表补「生成单元」，与 FR-GEN-02 的「生成范围」区分；表 A-3 补「跨应用系统的资料改动」一项，'
    '甲系统改接口字段时乙系统的对应知识是否重算，待委托方确认系统边界后另行确定。',

    '· 需求总数由 137 条增至 138 条，「必须」由 104 条增至 105 条，「应该」32 条与「可选」1 条不变；'
    '表 1-3、表 3-2、表 A-1、表 B-1 同步，表 3-2 与表 B-1 的条数合计均校验为 138 条。'
    'P1 由 101 条增至 102 条，P2 35 条、P3 1 条不变。验收标准由 481 条增至 485 条。',

    '· 编号提示：FR-GEN-32、FR-GEN-33 两个编号在 v3.14 并入 FR-GEN-31 后留空，此后分别由'
    '「资料变化后的重算范围」（v3.16）与「资料失效后的知识层处置」（本版）使用；'
    'v3.14 一段中的这两个编号指的是被并入的原条目，不指向现行条目。',
]

h_el = deepcopy(tpl_head._p)
hp = Paragraph(h_el, d)
hp.runs[0].text = HEAD
for r in hp.runs[1:]:
    r._element.getparent().remove(r._element)
hp.paragraph_format.keep_with_next = True

els = [h_el]
for txt in ITEMS:
    e = deepcopy(tpl_item._p)
    pp = Paragraph(e, d)
    pp.runs[0].text = txt
    for r in pp.runs[1:]:
        r._element.getparent().remove(r._element)
    els.append(e)

for e in els:
    anchor._p.addprevious(e)
print('附录 C 插入 v3.18，共', len(ITEMS), '条')

# ---------- 7. 版本号三处 ----------
# 7a 封面标题下独立段落
pver = find_para('v3.17')
assert pver.text.strip() == 'v3.17', pver.text
pver.runs[0].text = 'v3.18'
print('封面版本段 →', pver.text.strip())

# 7b 封面表格「文档版本」行
t0 = find_tbl(['项目', '内容'])
for r in t0.rows:
    if r.cells[0].text.strip() == '文档版本':
        old = r.cells[1].text.strip()
        assert old.startswith('v3.17'), old
        fill_cell(r.cells[1],
                  'v3.18（在 v3.17 基础上补资料失效的处置：新增 FR-GEN-33 与表 6-9，'
                  '把整批撤销、单份停用、整包重传未出现三种情形下的知识层失效写成一条规则；'
                  '表 4-5 补「疑似同一份」、FR-GEN-32 补项目知识总览、表 7-4 补「已废止」的触发动作）')
        print('封面文档版本 →', r.cells[1].text[:60])
    if r.cells[0].text.strip() == '日期':
        fill_cell(r.cells[1], '2026-10-30')
        print('日期 →', r.cells[1].text.strip())

# 7c 页眉
for s in d.sections:
    for p in s.header.paragraphs:
        if 'v3.17' in p.text:
            done = False
            for r in p.runs:
                if 'v3.17' in r.text:
                    r.text = r.text.replace('v3.17', 'v3.18')
                    done = True
            assert done, [r.text for r in p.runs]
            print('页眉 →', p.text.strip())

d.save(DST)
print('saved', DST)

# ---------- 校验 ----------
d2 = docx.Document(DST)
print('--- 附录 C 首段')
P = list(d2.paragraphs)
k = [i for i, p in enumerate(P) if p.text.strip().startswith('v3.18 修订')][0]
for j in range(k - 1, k + 3):
    print(j, '|', P[j].text.strip()[:90])
