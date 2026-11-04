# -*- coding: utf-8 -*-
"""v3.20 第二步：同步汇总表与表 8-5 的出处列。

按题注定位表格，不按索引——插入行/表会让后续索引全变。
表 3-2 与表 B-1 列序不同（编号分别在 cells[2] / cells[3]），各自按列名取。
"""
import os

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")

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
    for p in cell.paragraphs:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)
                CHANGED.append((old, new))
                return True
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


def main():
    doc = Document(DOC)

    # ── 表 1-3 优先级分布：必须 108 -> 111
    _, r = row_by(doc, "表 1-3", 0, "必须")
    assert rep(r.cells[2], "共 108 条", "共 111 条")

    # ── 表 3-2（编号 cells[2]、条数 cells[3]）
    tb, r = row_by(doc, "表 3-2", 1, "图谱建立")
    assert rep(r.cells[2], "FR-STO-06、23~24、26~27", "FR-STO-06、23~24、26~27、31~32")
    assert rep(r.cells[3], "5", "7")
    tb, r = row_by(doc, "表 3-2", 1, "来源约束")
    assert rep(r.cells[2], "FR-QA-07~09、22~23", "FR-QA-07~09、22~24")
    assert rep(r.cells[3], "5", "6")

    # ── 表 B-1（编号 cells[3]、条数 cells[4]）
    tb, r = row_by(doc, "表 B-1", 2, "图谱建立")
    assert rep(r.cells[3], "FR-STO-06、23~24、26~27", "FR-STO-06、23~24、26~27、31~32")
    assert rep(r.cells[4], "5", "7")
    tb, r = row_by(doc, "表 B-1", 2, "来源约束")
    assert rep(r.cells[3], "FR-QA-07~09、22~23", "FR-QA-07~09、22~24")
    assert rep(r.cells[4], "5", "6")

    # ── 表 A-1 交付批次：P1 105 -> 108
    _, r = row_by(doc, "表 A-1", 0, "P1 演示核心")
    assert rep(r.cells[2], "105 条：108 条「必须」", "108 条：111 条「必须」")

    # ── 表 8-5：节点补「作业步骤」、两类出处列补全
    tb, r = row_by(doc, "表 8-5", 0, "节点类型")
    assert rep(r.cells[1], "数据库表、作业、交易、DD 语句",
               "数据库表、作业、作业步骤、交易、DD 语句")
    assert rep(r.cells[2], "交易节点由资源定义的交易码绑定产出；占位对象按 FR-STO-24 落图",
               "交易节点由资源定义的交易码绑定产出；业务术语节点由术语表与人工确认产出；"
               "知识条目节点随知识成文产生；应用系统节点按分图配置产出；"
               "占位对象按 FR-STO-24 落图")

    tb, r = row_by(doc, "表 8-5", 0, "关系类型")
    assert rep(r.cells[2], "结构事实中的调用边、触发边与访问边、术语与代码对象的绑定、条目对其来源的引用",
               "结构事实中的调用边、触发边与访问边、包含边与定义边、术语与代码对象的绑定、"
               "条目对其来源的引用；归属边按分图配置产出；"
               "别名边按资料中给出的同一对象的两种名称（如数据集的逻辑名与物理名）产出")

    doc.save(DOC)
    for a, b in CHANGED:
        print("  %-46s -> %s" % (a[:46], b[:60]))
    print("共 %d 处" % len(CHANGED))


if __name__ == "__main__":
    main()
