# -*- coding: utf-8 -*-
"""v3.20 第五步：调表 8-5 列宽，避免「经验条目」在条目中间断行。

表 8-5 加「作业步骤」后取值列变挤，「知识条目、经验条目、占位对象」被 Word 断成
「经验条 / 目」。总宽不变，只把「项」列的富余宽度让给「取值」列：
1474 / 3855 / 3685  ->  1200 / 4131 / 3683

tblGrid/gridCol 与每格 tcW 必须同时改，只改一处 Word 仍按旧列数渲染。
"""
import os

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w"
NEW = [1200, 4131, 3683]


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


def main():
    doc = Document(DOC)
    tb = table_by_caption(doc, "表 8-5")
    tbl = tb._tbl

    grid = tbl.find(".//{%s}tblGrid" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    gcs = grid.findall("{%s}gridCol" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    assert len(gcs) == 3, len(gcs)
    for gc, w in zip(gcs, NEW):
        gc.set(W + "w", str(w))

    total = 0
    for r in tb.rows:
        for c, w in zip(r.cells, NEW):
            tcw = c._tc.find(".//{%s}tcW" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            tcw.set(W + "w", str(w))
            tcw.set(W + "type", "dxa")
        total += 1

    tw = tbl.find(".//{%s}tblW" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    tw.set(W + "w", str(sum(NEW)))
    tw.set(W + "type", "dxa")

    doc.save(DOC)

    doc = Document(DOC)
    tb = table_by_caption(doc, "表 8-5")
    grid = tb._tbl.find(".//{%s}tblGrid" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    print("gridCol:", [c.get(W + "w") for c in grid], "合计", sum(NEW))
    print("行数", len(tb.rows), "每行 tcW 一致:",
          all([c._tc.find('.//{%s}tcW' % "http://schemas.openxmlformats.org/wordprocessingml/2006/main").get(W + "w")
               for c in r.cells] == [str(x) for x in NEW] for r in tb.rows))


if __name__ == "__main__":
    main()
