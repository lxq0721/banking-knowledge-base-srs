# -*- coding: utf-8 -*-
"""v3.18 step2c：表 6-9 文案压缩 + 列宽重分配（原继承表 6-8 列宽，判定列为英文混排后出现孤字）"""
import sys, io
import docx
from docx.oxml.ns import qn

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = '_v318_step2.docx'
d = docx.Document(SRC)

ROWS = [
    ['情形', '判定', '处置', '可恢复方式'],
    ['整批撤销',
     '对已完成的导入批次执行撤销',
     '该批次全部资料的知识层产物失效',
     '重新提交该批'],
    ['单份停用',
     '对单份资料执行停用',
     '仅该资料的产物失效，同批不受影响',
     '解除停用后重算'],
    ['整包重传未出现',
     '声明整包重传后，与上一版比对缺失项',
     '按单份停用处置，标「本次未提交」',
     '下次提交补齐'],
]
WIDTHS = [1900, 2800, 2700, 1626]


def fill_cell(cell, text):
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p0 = cell.paragraphs[0]
    assert p0.runs, cell.text[:30]
    p0.runs[0].text = text
    for r in p0.runs[1:]:
        r._element.getparent().remove(r._element)


TB = None
for tb in d.tables:
    if [c.text.strip() for c in tb.rows[0].cells] == ['情形', '判定', '处置', '可恢复方式']:
        TB = tb
assert TB is not None

for tr, vals in zip(TB.rows, ROWS):
    for c, v in zip(tr.cells, vals):
        fill_cell(c, v)

grid = TB._tbl.find(qn('w:tblGrid'))
gcs = grid.findall(qn('w:gridCol'))
assert len(gcs) == 4
for gc, w in zip(gcs, WIDTHS):
    gc.set(qn('w:w'), str(w))
for row in TB.rows:
    for c, w in zip(row.cells, WIDTHS):
        tcW = c._tc.tcPr.find(qn('w:tcW'))
        tcW.set(qn('w:w'), str(w))
        tcW.set(qn('w:type'), 'dxa')
print('表 6-9 列宽 →', WIDTHS, '合计', sum(WIDTHS))
for r in TB.rows:
    print(' | '.join(c.text.strip() for c in r.cells))

d.save(SRC)
print('saved', SRC)
