# -*- coding: utf-8 -*-
"""v3.20 第三步：正文统计、附录 C 新版本段、版本号三处、文档属性。

注意：附录 C 新版本必须用 anchor.addprevious 逐段插入（addnext 会让顺序颠倒）；
定位段落时要跳过 toc 样式，否则会先命中目录行把内容插到文档开头。
"""
import os
from copy import deepcopy

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")

VER = "v3.20"
DATE = "2026-11-04"

CHANGES = [
    # 1 正文：七个功能域条数
    ("七个功能域共 134 条功能需求；连同数据需求的 3 条与非功能需求的 4 条，合计 141 条需求。",
     "七个功能域共 137 条功能需求；连同数据需求的 3 条与非功能需求的 4 条，合计 144 条需求。"),
    # 2 表 B-1 表注
    ("注：合计 141 条需求，含 137 条功能需求（七个功能域 134 条 + 数据需求 3 条）与 4 条非功能需求",
     "注：合计 144 条需求，含 140 条功能需求（七个功能域 137 条 + 数据需求 3 条）与 4 条非功能需求"),
]

C_HEAD = ("v3.20 修订　补边的粒度、遍历方向与检索入口三处口径，使「答案落在这条链路的某一行」"
          "真正能查到：")

C_ITEMS = [
    "· 新增 FR-STO-31「边的节点粒度」（必须）：读写、调用、触发边应挂在能够定位到的最细节点上——"
    "程序内挂段落，作业内挂作业步骤，交易挂交易节点；同一程序对同一文件的多处访问落为多条段落边，"
    "不得合并为一条程序级边，程序级关系由段落边沿「包含」向上聚合。原文只在表 8-5 列出了节点与"
    "关系的类型，没有一处规定一条边挂在哪一级节点上，于是「§1500 读 A 文件、§1600 写 B 文件」"
    "这种同一程序内两段代码管两个文件的情形，被一条程序级边抹平，问「账户状态在哪一段被改」"
    "只能给出整个程序。证据位置本已要求落到源码行（FR-STO-06），但证据到行而节点在程序，"
    "这条要求实际落不下去。",

    "· 新增 FR-STO-32「可达闭包的遍历方向」（必须）：可达闭包应逐类关系定义可遍历方向；"
    "「调用」与「读写」须支持反向，由被访问的文件或表反查访问它的程序与段落；「包含」沿出边；"
    "「归属」只用于限定应用系统，「依据」与「引用」不参与闭包扩展。FR-STO-30 写的是「可达闭包」，"
    "而图论的可达默认沿出边，读写边的方向是程序→文件，于是以数据为起点的链路问题取不到任何对象——"
    "这不是理解分歧，是方向口径未定，按字面实现该条直接空转。反向得到的路径须标注方向，"
    "不得展示为正向调用顺序。",

    "· 新增 FR-QA-24「问题中的对象名锚定」（必须）：问题中出现可按唯一键匹配的对象名时，"
    "应先以该节点为检索入口，再在 FR-STO-30 的闭包内按相似度排序；匹配在问题所指应用系统内按"
    "FR-STO-24 的唯一键进行，命中多个同名对象并列呈现并标注所属应用系统；未匹配时退回相似度检索"
    "并说明未找到该对象。原文的检索入口只有相似度一条，问题里明写着 ACCTFILE、FLAG-POSTPONE，"
    "走向量可能命中表结构，也可能命中某句提到「账户文件」的段落，入口本身不确定。",

    "· 表 8-5 节点类型补「作业步骤」，并把「由何产出」两格补全：节点格原只交代了结构事实与交易"
    "节点的产出，应用系统、业务术语、知识条目三类没有出处；关系格原只覆盖调用、触发、读写、"
    "引用、依据五类，「包含」「定义」「归属」「别名」四类无出处，其中「别名」全文仅有此一处出现，"
    "入库时每人都能给出一种解释。此处按 FR-STO-31 的粒度要求补作业步骤节点（表 8-4 早已把"
    "作业步骤作为 JCL 的切片单元），并逐类补出处，不改变类型清单的取值范围。",

    "· 需求总数由 141 条增至 144 条，「必须」由 108 条增至 111 条，「应该」32 条与「可选」1 条不变；"
    "表 1-3、表 3-2、表 A-1、表 B-1 同步，表 3-2 与表 B-1 的条数合计均校验为 144 条。"
    "P1 由 105 条增至 108 条，P2 35 条、P3 1 条不变——三条都卡在「答案能不能落到链路上的某一行」，"
    "不属可延后的治理项。",

    "· 编号提示：FR-STO-31、FR-STO-32、FR-QA-24 分别接在各域当前最大编号之后新增，"
    "未占用 FR-STO-25、FR-STO-28、FR-QA-03、FR-QA-15 等已留空的编号。",
]

COVER_DESC = ("v3.20（在 v3.19 基础上补三处口径：新增边的节点粒度、可达闭包的遍历方向、"
              "问题中的对象名锚定三条必须需求；表 8-5 补作业步骤节点并补全节点与关系的出处。"
              "需求 144 条，其中必须 111 条）")


def items_of(doc):
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            out.append(Paragraph(ch, doc))
        elif ch.tag.endswith("}tbl"):
            out.append(Table(ch, doc))
    return out


def set_text(p, text):
    """保留首个 run 的格式，改写整段文本。"""
    assert p.runs, p.text[:40]
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)


def main():
    doc = Document(DOC)
    ps = [p for p in doc.paragraphs]

    # ── 正文两处统计
    done = 0
    for old, new in CHANGES:
        hit = False
        for p in ps:
            if old in p.text:
                set_text(p, p.text.replace(old, new))
                hit = True
                done += 1
                break
        assert hit, old[:30]
    print("正文统计替换", done, "处")

    # ── 附录 C：v3.20 段插在 v3.19 段之前
    anchor = next(p for p in ps
                  if p.style.name == "Normal" and p.text.strip().startswith("v3.19 修订"))
    tpl_head = anchor._p
    tpl_item = next(p for p in ps[ps.index(anchor):]
                    if p.style.name == "Normal" and p.text.strip().startswith("· "))._p

    for txt, tpl in [(C_HEAD, tpl_head)] + [(t, tpl_item) for t in C_ITEMS]:
        el = deepcopy(tpl)
        anchor._p.addprevious(el)
        new_p = Paragraph(el, anchor._parent)
        set_text(new_p, txt)
    print("附录 C 插入", 1 + len(C_ITEMS), "段")

    # ── 版本号三处
    ps = [p for p in doc.paragraphs]
    i = next(k for k, p in enumerate(ps) if p.text.strip() == "需求规格书")
    assert ps[i + 1].text.strip().startswith("v3.")
    set_text(ps[i + 1], VER)
    for tb in doc.tables:
        for r in tb.rows:
            k = r.cells[0].text.strip()
            if k == "文档版本":
                set_text(r.cells[1].paragraphs[0], COVER_DESC)
            elif k == "日期":
                set_text(r.cells[1].paragraphs[0], DATE)
    for s in doc.sections:
        for p in s.header.paragraphs:
            if "需求规格书" in p.text and "v3." in p.text:
                import re
                set_text(p, re.sub(r"v3\.\d+", VER, p.text))

    # ── 文档属性
    cp = doc.core_properties
    cp.title = "核心系统迁移知识库需求规格书"
    cp.subject = "%s" % VER
    cp.author = "核心系统迁移知识库项目组"
    cp.last_modified_by = "核心系统迁移知识库项目组"
    cp.comments = ""
    cp.keywords = ""
    cp.category = ""
    cp.revision = 1

    doc.save(DOC)
    print("saved")

    # ── 回读校验
    doc = Document(DOC)
    ps = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for t in ps:
        if t.startswith("v3.20 修订"):
            print("\n[OK] 附录 C 首段:", t[:70])
            break
    print("[OK] 版本段落:", ps[:3])
    print("[OK] 页眉:", doc.sections[0].header.paragraphs[-1].text)
    print("[OK] 日期行:", [p.text for p in doc.tables[0].rows[-1].cells] if doc.tables else "")


if __name__ == "__main__":
    main()
