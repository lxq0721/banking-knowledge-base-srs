# -*- coding: utf-8 -*-
"""v3.18 清单：新增 FR-GEN-33 三条 AC 与 FR-IMP-10 一条 AC，同步三条需求的描述，重排编号，出 PDF"""
import os
import sys
import io
import shutil
import time
import openpyxl
import win32com.client as win32

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = os.path.abspath('../../release/v3.17/核心系统迁移知识库需求规格书_v3.17_验收标准清单.xlsx')
OUT_DIR = os.path.abspath('../../release/v3.18')
os.makedirs(OUT_DIR, exist_ok=True)
XLSX = os.path.join(OUT_DIR, '核心系统迁移知识库需求规格书_v3.18_验收标准清单.xlsx')
PDF = os.path.join(OUT_DIR, '核心系统迁移知识库需求规格书_v3.18_验收标准清单.pdf')

shutil.copyfile(SRC, XLSX)


def retry(fn, n=12, wait=3):
    le = None
    for i in range(n):
        try:
            return fn()
        except Exception as e:
            le = e
            time.sleep(wait)
    raise le


# ---------- 1. COM 插入空行（保留上方行的格式） ----------
xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
try:
    x = retry(lambda: xl.Workbooks.Open(XLSX, ReadOnly=False))
    s = x.Worksheets('验收清单')
    # 先插大行号：136 行（FR-GEN-32 末条 AC-132）之后插 3 行
    for _ in range(3):
        retry(lambda: s.Rows(137).Insert(-4121))
    # 再插小行号：45 行（FR-IMP-10 末条 AC-41）之后插 1 行
    retry(lambda: s.Rows(46).Insert(-4121))
    retry(lambda: x.Save())
    retry(lambda: x.Close(SaveChanges=0))
    print('插入 4 行空行完成')
finally:
    try:
        xl.Quit()
    except Exception:
        pass

# ---------- 2. openpyxl 写内容 ----------
DESC_IMP10 = ('系统应在导入前计算每份资料的内容指纹，并把本次提交的资料按路径、内容指纹与文件名主干'
              '归入表 4-5 的一种情形；只对「新增」、「同路径变更」以及经人指定接续的「疑似同一份」'
              '资料执行后续环节。')
DESC_IMP16 = '系统应支持按批次撤销已完成的导入；撤销后该批次资料按 FR-GEN-33 处置，原件快照保留。'
DESC_GEN32 = ('已导入资料新增或变更后，系统只重算直接引用该资料的单元与跨文件聚合的汇总产物'
              '（数据实体说明、业务术语表、系统结构说明、项目知识总览），其中项目知识总览按条目刷新、'
              '不整份重生成；未引用它的单元不得重算，也不得因补一份资料触发整个应用系统重新生成。'
              '重算按 FR-GEN-05 另出新版本，已修正内容按 FR-REV-24 保留，范围外对象只更新占位引用，'
              '见 FR-GEN-03。')

NEW_IMP10_AC = ('同一份资料改名并修改内容后再次提交，导入清单在新建记录前列出候选；'
                '由人指定接续后并入原记录，清单中该资料只有一条记录。')

GEN33_ROWS = [
    ('FR-GEN-33', '资料失效后的知识层处置', '6 知识生成', '6.3.1 触发与范围', '必须',
     '资料被撤销、停用或不再提交后，由它生成的产物应同步失效：知识文档转为「已废止」，'
     '其切片与向量退出检索，图谱节点与关系按 FR-STO-27 下线，经它确认的经验条目一并失效；'
     '该资料参与构成的跨文件汇总产物（数据实体说明、业务术语表、系统结构说明、项目知识总览）随之重算。'
     '原件的只读快照保留，指向该资料的引用改为占位并标「资料已失效」。三种触发情形的判定与处置见表 6-9。',
     '撤销一个已完成的导入批次后，该批次资料的知识文档状态为「已废止」，其切片与向量不再出现在检索结果中，'
     '图谱中由这些资料建立的节点下线，原件快照仍可打开。'),
    ('FR-GEN-33', '资料失效后的知识层处置', '6 知识生成', '6.3.1 触发与范围', '必须',
     '资料被撤销、停用或不再提交后，由它生成的产物应同步失效：知识文档转为「已废止」，'
     '其切片与向量退出检索，图谱节点与关系按 FR-STO-27 下线，经它确认的经验条目一并失效；'
     '该资料参与构成的跨文件汇总产物（数据实体说明、业务术语表、系统结构说明、项目知识总览）随之重算。'
     '原件的只读快照保留，指向该资料的引用改为占位并标「资料已失效」。三种触发情形的判定与处置见表 6-9。',
     '对同一批次中的一份资料执行停用，只有该资料的产物失效，同批其他资料的知识文档状态与检索命中不变；'
     '解除停用后按 FR-GEN-32 重算，该资料恢复为可取用。'),
    ('FR-GEN-33', '资料失效后的知识层处置', '6 知识生成', '6.3.1 触发与范围', '必须',
     '资料被撤销、停用或不再提交后，由它生成的产物应同步失效：知识文档转为「已废止」，'
     '其切片与向量退出检索，图谱节点与关系按 FR-STO-27 下线，经它确认的经验条目一并失效；'
     '该资料参与构成的跨文件汇总产物（数据实体说明、业务术语表、系统结构说明、项目知识总览）随之重算。'
     '原件的只读快照保留，指向该资料的引用改为占位并标「资料已失效」。三种触发情形的判定与处置见表 6-9。',
     '声明整包重传且提交包中缺少一份原有资料时，该资料在导入清单标「本次未提交」且产物失效；'
     '未声明整包的提交中缺少同一文件时，已有记录与其产物不受影响（与 FR-IMP-21 一致）。'),
]

wb = openpyxl.load_workbook(XLSX)
ws = wb['验收清单']
LAST = ws.max_row

# 2a 既有行：改描述
n = {'FR-IMP-10': 0, 'FR-IMP-16': 0, 'FR-GEN-32': 0}
for r in range(5, LAST + 1):
    code = ws.cell(r, 2).value
    if code in n:
        ws.cell(r, 7).value = {'FR-IMP-10': DESC_IMP10,
                               'FR-IMP-16': DESC_IMP16,
                               'FR-GEN-32': DESC_GEN32}[code]
        n[code] += 1
print('描述改写', n)
assert n == {'FR-IMP-10': 9, 'FR-IMP-16': 3, 'FR-GEN-32': 3}, n

# 2b FR-GEN-32 的 AC-131 补项目知识总览
hit = 0
for r in range(5, LAST + 1):
    if ws.cell(r, 8).value == 'AC-131':
        ws.cell(r, 9).value = ('补充资料后，数据实体说明、业务术语表与项目知识总览同步更新，条目内容按新资料更新，'
                               '不存在沿用旧口径的条目。')
        hit += 1
assert hit == 1, hit

# 2c 新增行内容（46 为 FR-IMP-10 新 AC；138~140 为 FR-GEN-33 三条，行号已含 46 行插入的位移）
src = 45
vals = [ws.cell(src, c).value for c in range(1, 13)]
row46 = {2: 'FR-IMP-10', 3: '再次提交的处置', 4: '4 资料导入', 5: '4.3.6 导入治理',
         6: '应该', 7: DESC_IMP10, 9: NEW_IMP10_AC}
for c in range(1, 13):
    ws.cell(46, c).value = row46.get(c, None)
assert ws.cell(45, 2).value == 'FR-IMP-10' and ws.cell(47, 2).value == 'FR-IMP-15', \
    (ws.cell(45, 2).value, ws.cell(47, 2).value)

for i, (code, title, chap, grp, pri, desc, ac) in enumerate(GEN33_ROWS):
    r = 138 + i
    ws.cell(r, 2).value = code
    ws.cell(r, 3).value = title
    ws.cell(r, 4).value = chap
    ws.cell(r, 5).value = grp
    ws.cell(r, 6).value = pri
    ws.cell(r, 7).value = desc
    ws.cell(r, 9).value = ac
assert ws.cell(141, 2).value == 'FR-GEN-06', ws.cell(140, 2).value

# 2d 重排序号与 AC 编号
for r in range(5, LAST + 1):
    ws.cell(r, 1).value = r - 4
    ws.cell(r, 8).value = f'AC-{r - 4}'
TOTAL = LAST - 4
print('清单行数', TOTAL, '（应为 485）')

# 2e 顶部说明
ws.cell(2, 1).value = ('对应文档：核心系统迁移知识库需求规格书_v3.18.docx\u3000｜\u3000'
                       '本清单是验收标准的唯一出处，需求书正文不重复列示')
ws.cell(3, 1).value = ('共 138 条需求 / %d 条验收标准\u3000｜\u3000'
                       '优先级：必须 105 · 应该 32 · 可选 1\u3000｜\u3000'
                       '交付批次：P1 102 · P2 35 · P3 1' % TOTAL)
# 筛选与校验范围
ws.auto_filter.ref = f'A4:L{TOTAL + 4}'
for dv in ws.data_validations.dataValidation:
    dv.sqref = f'J5:J{TOTAL + 4}'
cf = ws.conditional_formatting
for rng in list(cf):
    rng.sqref = openpyxl.worksheet.cell_range.MultiCellRange(f'J5:J{TOTAL + 4}')

# ---------- 3. 统计页 ----------
st = wb['统计与说明']


def put(r, c, v):
    old = st.cell(r, c).value
    st.cell(r, c).value = v
    print(f'  统计 {r},{c}: {old} → {v}')


put(3, 2, f'数据源：核心系统迁移知识库需求规格书_v3.18.docx\u3000 138 条需求 / {TOTAL} 条验收标准')

# 先打印两张统计表的表头，确认列位（教训：按行号猜列会覆盖名称列与验收条数列）
for r in (6, 19):
    print('  表头', r, [st.cell(r, c).value for c in range(1, 7)])

put(7, 4, 57)            # 表一：4 资料导入 验收条数 56 → 57
put(9, 3, 28)            # 表一：6 知识生成 需求条数 27 → 28
put(9, 4, 106)           # 表一：6 知识生成 验收条数 103 → 106
put(9, 5, 23)            # 表一：6 知识生成 必须 22 → 23
put(16, 3, 138)          # 表一：合计 需求条数
put(16, 4, TOTAL)        # 表一：合计 验收条数
put(16, 5, 105)          # 表一：合计 必须
# 表二：B=优先级名称 C=需求条数 D=验收条数 E=需求占比 F=裁剪说明
put(20, 3, 105)          # 必须 需求条数
put(20, 4, 378)          # 必须 验收条数 375 → 378
put(20, 5, round(105 / 138, 3))
put(21, 3, 32)           # 应该 需求条数（不变，显式回写防错位）
put(21, 4, 105)          # 应该 验收条数 104 → 105
put(21, 5, round(32 / 138, 3))
put(22, 3, 1)
put(22, 4, 2)
put(22, 5, round(1 / 138, 3))
put(23, 3, 138)          # 合计 需求条数
put(23, 4, TOTAL)        # 合计 验收条数
put(23, 5, 1)

# 回读校验：名称列与合计必须仍在原位
assert st.cell(20, 2).value == '必须', st.cell(20, 2).value
assert st.cell(21, 2).value == '应该', st.cell(21, 2).value
assert st.cell(22, 2).value == '可选', st.cell(22, 2).value
assert st.cell(23, 2).value == '合计', st.cell(23, 2).value
assert st.cell(20, 3).value == 105 and st.cell(20, 4).value == 378
assert st.cell(21, 4).value == 105 and st.cell(22, 4).value == 2
assert st.cell(23, 3).value == 138 and st.cell(23, 4).value == TOTAL
assert st.cell(20, 4).value + st.cell(21, 4).value + st.cell(22, 4).value == TOTAL
print('  统计页校验通过')

wb.save(XLSX)
print('saved', XLSX)

# ---------- 4. AutoFit ----------
wl = win32.DispatchEx('Excel.Application')
wl.Visible = False
wl.DisplayAlerts = False
try:
    x = retry(lambda: wl.Workbooks.Open(XLSX, ReadOnly=False))
    s = x.Worksheets('验收清单')
    s.Rows(f'5:{TOTAL + 4}').AutoFit()
    retry(lambda: x.Save())
    for name in ('验收清单', '统计与说明'):
        sh = x.Worksheets(name)
        sh.PageSetup.Orientation = 2
        sh.PageSetup.Zoom = False
        sh.PageSetup.FitToPagesWide = 1
        sh.PageSetup.FitToPagesTall = False
        sh.PageSetup.PrintTitleRows = '$4:$4' if name == '验收清单' else '$1:$1'
    retry(lambda: x.ExportAsFixedFormat(0, PDF))
    retry(lambda: x.Close(SaveChanges=0))
    print('pdf', PDF)
finally:
    try:
        wl.Quit()
    except Exception:
        pass
