# -*- coding: utf-8 -*-
"""v3.19 第三步：正文统计、附录 C 新版本段、版本号三处、文档属性。

注意：附录 C 新版本必须用 anchor.addprevious 逐段插入（addnext 会让顺序颠倒）；
定位段落时要跳过 toc 样式，否则会先命中目录行把内容插到文档开头。
"""
import os
from copy import deepcopy

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19.docx")

VER = "v3.19"
DATE = "2026-11-04"

CHANGES = [
    # 1 正文：七个功能域条数
    ("七个功能域共 131 条功能需求；连同数据需求的 3 条与非功能需求的 4 条，合计 138 条需求。",
     "七个功能域共 134 条功能需求；连同数据需求的 3 条与非功能需求的 4 条，合计 141 条需求。"),
    # 2 表 B-1 表注
    ("注：合计 138 条需求，含 134 条功能需求（七个功能域 131 条 + 数据需求 3 条）与 4 条非功能需求",
     "注：合计 141 条需求，含 137 条功能需求（七个功能域 134 条 + 数据需求 3 条）与 4 条非功能需求"),
]

C_HEAD = ("v3.19 修订　补联机交易的整链产物与知识反查规则，堵住「跨程序的问题只能逐个程序翻」"
          "与「命中代码要现场重编说明」两个口子：")

C_ITEMS = [
    "· 新增 FR-GEN-34「交易与交易链路说明」（必须）：交易按交易码逐个成文，内容含交易码与入口程序、"
    "从入口到末端的调用链、链上各程序读写的表；交易码到入口程序的绑定取自资源定义，资源定义不可得时"
    "才取程序内的交易码分支。表 6-1 早已列出「交易（一交易一份）」的源码必出产物，2.2 也写了"
    "「程序、作业、交易、数据实体各自独立成文」，但功能需求只到 FR-GEN-11 的作业说明，交易这一层"
    "没有条目承接——联机交易只有资源定义给出的第一跳，往下要逐个程序翻第 7 节才能拼出链路。",

    "· 新增 FR-STO-30「关系链路问题的检索覆盖」（必须）：关系链路类问题的检索结果应覆盖问题所涉对象"
    "在图谱中的可达闭包，闭包内对象在检索库中缺切片的应在应答中显式标注，不得静默略过；"
    "以命中条数替代覆盖的做法不满足要求。FR-STO-21 只限定了知识版本与应用系统两个维度，"
    "没有约束「链路上那几个对象该不该在候选集里」，而漏掉的那一段措辞与问题本就不相似，"
    "排序永远把它排在尾部，调大条数带进来的是噪声。",

    "· 新增 FR-QA-23「命中源码时的知识取用」（必须）：命中源码切片时，按切片元数据中的对象编号"
    "取该对象已生成的知识节作为解释，并按对象类型分别落到程序知识文档、所属程序的流程与业务规则节、"
    "数据实体说明、交易说明；该对象尚未成文时只呈现原文并标注「该对象未成文」，不补写解释。"
    "8.1 原本只写了「切片 → 知识文档段落 → 源码行」这一个方向，反向没有条目，"
    "现场生成的内容不落库、不可复现，与 FR-QA-05、FR-QA-06 的稳定要求冲突。",

    "· 表 8-5 的节点类型补「交易」、关系类型补「触发」：FR-GRP-01 写的是「起点可为程序、交易或数据集」，"
    "而类型清单中没有交易节点，浏览视图的这个起点取不到；FR-STO-23 又规定清单外类型不得入库，"
    "两处只能改一处，此处补清单。",

    "· 表 6-5 补交易说明一行（索引通道、L1a、不经人工放行）；表 8-4 的「知识文档」切片行并入交易说明；"
    "表 3-3 补产物「交易与交易链路说明」；表 1-2 补术语「交易说明」；表 6-1 交易行的源码必出列"
    "由「交易入口与调用链」改为「交易入口与调用链、交易说明」，与 FR-GEN-34 对应。",

    "· 需求总数由 138 条增至 141 条，「必须」由 105 条增至 108 条，「应该」32 条与「可选」1 条不变；"
    "表 1-3、表 3-2、表 A-1、表 B-1 同步，表 3-2 与表 B-1 的条数合计均校验为 141 条。"
    "P1 由 102 条增至 105 条，P2 35 条、P3 1 条不变——三条都是链路闭合所必需，不属可延后的治理项。",

    "· 编号提示：FR-GEN-34、FR-STO-30、FR-QA-23 分别接在各域当前最大编号之后新增，"
    "未占用 FR-GEN-23、FR-STO-19、FR-STO-25、FR-QA-03、FR-QA-15 等已留空的编号。",
]

COVER_DESC = ("v3.19（在 v3.18 基础上补三处缺口：新增 FR-GEN-34 交易与交易链路说明、"
              "FR-STO-30 关系链路问题的检索覆盖、FR-QA-23 命中源码时的知识取用；"
              "表 8-5 补交易节点与触发关系。需求 141 条，其中必须 108 条）")


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

    # ── 附录 C：v3.19 段插在 v3.18 段之前
    anchor = next(p for p in ps
                  if p.style.name == "Normal" and p.text.strip().startswith("v3.18 修订"))
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
    # (1) 封面标题下的独立段落
    ps = [p for p in doc.paragraphs]
    i = next(k for k, p in enumerate(ps) if p.text.strip() == "需求规格书")
    assert ps[i + 1].text.strip().startswith("v3.")
    set_text(ps[i + 1], VER)
    # (2) 封面表格「文档版本」行
    for tb in doc.tables:
        for r in tb.rows:
            if r.cells[0].text.strip() == "文档版本":
                set_text(r.cells[1].paragraphs[0], COVER_DESC)
                break
            if r.cells[0].text.strip() == "日期":
                set_text(r.cells[1].paragraphs[0], DATE)
    # (3) 页眉
    for s in doc.sections:
        for p in s.header.paragraphs:
            if "需求规格书" in p.text and "v3." in p.text:
                set_text(p, p.text.replace("v3.18", VER))

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
        if t.startswith("v3.19 修订"):
            print("\n[OK] 附录 C 首段:", t[:70])
            break
    print("[OK] 版本段落:", [t for t in ps[:4]])
    print("[OK] 页眉:", doc.sections[0].header.paragraphs[-1].text)


if __name__ == "__main__":
    main()
