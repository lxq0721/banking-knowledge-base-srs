# -*- coding: utf-8 -*-
"""v3.20 第四步：压缩表 8-5「由何产出」两格文案，使三行表不跨页。

v3.20 补出处后表 8-5 长出约三行，「必带属性」行被挤到次页，只剩表头加一行。
文案里的「节点」「产出」在列名中已重复，逐处删掉即可省两行。
"""
import os

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")

NODE_SRC = ("结构事实与经人工确认的经验条目；交易由资源定义的交易码绑定，业务术语由术语表，"
            "知识条目随成文产生，应用系统按分图配置，占位对象按 FR-STO-24 落图")

REL_SRC = ("结构事实中的调用、触发、访问、包含、定义边，术语绑定，条目对来源的引用；"
           "归属按分图配置，别名取资料给出的同一对象两名称（如数据集逻辑名与物理名）")


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


def set_cell(cell, text):
    """整格改写：首段承载全文，其余段落删除。"""
    ps = cell.paragraphs
    ps[0].runs[0].text = text
    for r in ps[0].runs[1:]:
        r._element.getparent().remove(r._element)
    for p in ps[1:]:
        p._p.getparent().remove(p._p)


def main():
    doc = Document(DOC)
    tb = table_by_caption(doc, "表 8-5")
    node = next(r for r in tb.rows if r.cells[0].text.strip() == "节点类型")
    rel = next(r for r in tb.rows if r.cells[0].text.strip() == "关系类型")
    print("原节点出处 %d 字 -> %d 字" % (len(node.cells[2].text), len(NODE_SRC)))
    print("原关系出处 %d 字 -> %d 字" % (len(rel.cells[2].text), len(REL_SRC)))
    set_cell(node.cells[2], NODE_SRC)
    set_cell(rel.cells[2], REL_SRC)
    doc.save(DOC)

    doc = Document(DOC)
    tb = table_by_caption(doc, "表 8-5")
    for r in tb.rows:
        print(" |", " | ".join(c.text.strip().replace("\n", " / ")[:70] for c in r.cells))


if __name__ == "__main__":
    main()
