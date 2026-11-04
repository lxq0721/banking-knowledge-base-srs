# -*- coding: utf-8 -*-
"""v3.19 第二步：同步十处表格与两处正文统计。

按题注定位表格，不按索引——插入行/表会让后续索引全变。
表 3-2 与表 B-1 列序不同（编号分别在 cells[2] / cells[3]），各自按列名取。
"""
import os
from copy import deepcopy

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19.docx")

CHANGED = []


def body_items(doc):
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            out.append(Paragraph(ch, doc))
        elif ch.tag.endswith("}tbl"):
            out.append(Table(ch, doc))
    return out


def table_by_caption(doc, cap):
    last = ""
    for obj in body_items(doc):
        if isinstance(obj, Paragraph):
            t = obj.text.strip()
            if t.startswith("表 "):
                last = t
        else:
            if last.startswith(cap):
                return obj
    raise KeyError(cap)


def rep(cell, old, new):
    """单元格内替换；跨 run 时把文本并到首 run。成功返回 True。"""
    joined = cell.text
    if old not in joined:
        return False
    if joined.replace(old, new) == joined:
        return False
    # 单 run 直接替换
    for p in cell.paragraphs:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)
                CHANGED.append((old, new))
                return True
    # 跨 run：写回首 run
    p = cell.paragraphs[0]
    p.runs[0].text = joined.replace(old, new)
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)
    CHANGED.append((old + " (跨run)", new))
    return True


def row_by(doc, cap, col, value):
    tb = table_by_caption(doc, cap)
    for r in tb.rows:
        if r.cells[col].text.strip() == value:
            return tb, r
    raise KeyError("%s / %s" % (cap, value))


def insert_row_after(tb, anchor_row, values):
    new_tr = deepcopy(anchor_row._tr)
    anchor_row._tr.addnext(new_tr)
    row = tb.rows[[r._tr for r in tb.rows].index(new_tr)]
    for c, v in zip(row.cells, values):
        p = c.paragraphs[0]
        if p.runs:
            p.runs[0].text = v
            for r in p.runs[1:]:
                r._element.getparent().remove(r._element)
        else:
            p.add_run(v)
    CHANGED.append(("+row", values[0]))
    return row


def main():
    doc = Document(DOC)

    # ── 表 1-3 优先级分布：必须 105 -> 108
    tb = table_by_caption(doc, "表 1-3")
    _, r = row_by(doc, "表 1-3", 0, "必须")
    assert rep(r.cells[2], "共 105 条", "共 108 条")

    # ── 表 3-2 功能结构分解（编号 cells[2]，条数 cells[3]）
    tb, r = row_by(doc, "表 3-2", 1, "对象产物")
    assert rep(r.cells[2], "FR-GEN-08~14、29~30", "FR-GEN-08~14、29~30、34")
    assert rep(r.cells[3], "9", "10")
    tb, r = row_by(doc, "表 3-2", 1, "向量化与索引")
    assert rep(r.cells[2], "FR-STO-12、18、20~22", "FR-STO-12、18、20~22、30")
    assert rep(r.cells[3], "5", "6")
    tb, r = row_by(doc, "表 3-2", 1, "来源约束")
    assert rep(r.cells[2], "FR-QA-07~09、22", "FR-QA-07~09、22~23")
    assert rep(r.cells[3], "4", "5")

    # ── 表 B-1 追溯矩阵（列序不同：编号 cells[3]、条数 cells[4]）
    tb, r = row_by(doc, "表 B-1", 2, "对象产物")
    assert rep(r.cells[3], "FR-GEN-08~14、29~30", "FR-GEN-08~14、29~30、34")
    assert rep(r.cells[4], "9", "10")
    tb, r = row_by(doc, "表 B-1", 2, "向量化与索引")
    assert rep(r.cells[3], "FR-STO-12、18、20~22", "FR-STO-12、18、20~22、30")
    assert rep(r.cells[4], "5", "6")
    tb, r = row_by(doc, "表 B-1", 2, "来源约束")
    assert rep(r.cells[3], "FR-QA-07~09、22", "FR-QA-07~09、22~23")
    assert rep(r.cells[4], "4", "5")

    # ── 表 A-1 交付批次：P1 102 -> 105
    _, r = row_by(doc, "表 A-1", 0, "P1 演示核心")
    assert rep(r.cells[2], "102 条：105 条「必须」", "105 条：108 条「必须」")

    # ── 表 6-1 交易行的源码必出列
    _, r = row_by(doc, "表 6-1", 0, "交易（一交易一份）")
    assert rep(r.cells[1], "交易入口与调用链", "交易入口与调用链、交易说明")

    # ── 表 6-5 两条通道：补交易说明（索引通道，L1a）
    tb, r = row_by(doc, "表 6-5", 0, "作业步骤与 DD 清单")
    insert_row_after(tb, r, ["交易入口与调用链、链上读写表", "索引", "否", "L1a",
                             "是，不经人工放行"])

    # ── 表 8-4 切片口径：知识文档行并入交易说明
    _, r = row_by(doc, "表 8-4", 0, "知识文档")
    assert rep(r.cells[0], "知识文档", "知识文档、交易说明")

    # ── 表 8-5 图谱类型：节点补「交易」、关系补「触发」
    tb, r = row_by(doc, "表 8-5", 0, "节点类型")
    assert rep(r.cells[1], "数据库表、作业、DD 语句", "数据库表、作业、交易、DD 语句")
    assert rep(r.cells[2], "结构事实与经人工确认的经验条目",
               "结构事实与经人工确认的经验条目；交易节点由资源定义的交易码绑定产出")
    tb, r = row_by(doc, "表 8-5", 0, "关系类型")
    assert rep(r.cells[1], "调用、包含", "调用、触发、包含")
    assert rep(r.cells[2], "结构事实中的调用边与访问边",
               "结构事实中的调用边、触发边与访问边")

    # ── 表 3-3 产物清单：补交易说明
    tb, r = row_by(doc, "表 3-3", 0, "知识文档（10 节）")
    insert_row_after(tb, r, ["交易与交易链路说明", "知识生成", "知识文档区（一交易一份）",
                             "检索与图谱浏览，作为交易维度的导航"])

    # ── 表 1-2 术语表：补「交易说明」
    tb, r = row_by(doc, "表 1-2", 0, "知识文档")
    insert_row_after(tb, r, ["交易说明",
                             "按交易码成文的链路说明，含交易码与入口程序、从入口到末端的调用链、"
                             "链上各程序读写的表；只由结构事实组成，属索引通道，见 FR-GEN-34。"])

    doc.save(DOC)
    for a, b in CHANGED:
        print("  %-42s -> %s" % (a[:42], b[:46]))
    print("共 %d 处" % len(CHANGED))


if __name__ == "__main__":
    main()
