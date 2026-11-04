# -*- coding: utf-8 -*-
"""v3.19 第四步：封面日期、FR-STO-23 段落结构修正。

FR-STO-23 原为「编号 / 优先级 / 描述」三段，约束内容被并进描述段，成了
「描述：约束：类型清单见表 8-5……」。本版改表 8-5，顺带恢复四段结构。
"""
import os
from copy import deepcopy

from docx import Document
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19.docx")


def set_text(p, text):
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)


def main():
    doc = Document(DOC)

    # ── 封面日期
    hit = False
    for tb in doc.tables:
        for r in tb.rows:
            if len(r.cells) >= 2 and r.cells[0].text.strip() == "日期":
                set_text(r.cells[1].paragraphs[0], "2026-11-04")
                hit = True
                break
        if hit:
            break
    assert hit, "未找到日期行"
    print("封面日期 -> 2026-11-04")

    # ── FR-STO-23 恢复四段
    ps = [p for p in doc.paragraphs if p.text.strip()]
    i = next(k for k, p in enumerate(ps) if p.text.strip().startswith("FR-STO-23"))
    bad = ps[i + 2]
    assert bad.text.strip().startswith("描述：约束："), bad.text[:40]

    el = deepcopy(bad._p)
    bad._p.addnext(el)
    set_text(bad, "约束：类型清单见表 8-5，清单外类型不得入库。")
    desc = Paragraph(el, bad._parent)
    set_text(desc, "描述：图谱的节点与关系应只使用表 8-5 列出的类型，新增类型须先更新清单再入库；"
                   "入库过程中不得自行产生清单外的类型名。")
    print("FR-STO-23 约束段与描述段已分开")

    doc.save(DOC)

    doc = Document(DOC)
    ps = [p for p in doc.paragraphs if p.text.strip()]
    i = next(k for k, p in enumerate(ps) if p.text.strip().startswith("FR-STO-23"))
    for j in range(i, i + 4):
        print("  ", ps[j].text[:90])
    for tb in doc.tables[:1]:
        for r in tb.rows:
            if r.cells[0].text.strip() == "日期":
                print("  日期:", r.cells[1].text.strip())


if __name__ == "__main__":
    main()
