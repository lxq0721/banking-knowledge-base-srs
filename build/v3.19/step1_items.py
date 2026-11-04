# -*- coding: utf-8 -*-
"""v3.19 第一步：插入三条新条目。

FR-GEN-34 交易与交易链路说明            6.3.3 对象产物，紧跟 FR-GEN-12
FR-STO-30 关系链路问题的检索覆盖        8.3.3 向量化与索引，紧跟 FR-STO-21
FR-QA-23  命中源码时的知识取用          10.3.4 来源约束，紧跟 FR-QA-08

模板取 FR-STO-24（编号 / 优先级 / 约束 / 描述 四段齐全），deepcopy 整段
继承字体、颜色与分页属性；插入后改写文本并裁掉多余 run。
"""
import os
import re
from copy import deepcopy

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19.docx")


def body_items(doc):
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            out.append(Paragraph(ch, doc))
        elif ch.tag.endswith("}tbl"):
            out.append(Table(ch, doc))
    return out


def find_para(items, prefix):
    for p in items:
        if isinstance(p, Paragraph) and p.text.strip().startswith(prefix):
            return p
    raise KeyError(prefix)


def item_block(items, prefix):
    """取一个条目的连续段落：从编号段起，到下一个条目的编号段之前。"""
    idx = next(i for i, p in enumerate(items)
               if isinstance(p, Paragraph) and p.text.strip().startswith(prefix))
    out = []
    for p in items[idx:]:
        if not isinstance(p, Paragraph):
            break
        if out and re.match(r"^(FR|NFR)-[A-Z]{2,3}-\d+", p.text.strip()):
            break
        if out and p.style.name.startswith("Heading"):
            break
        out.append(p)
    return out


def rewrite(paras, num, title, prio, cons, desc):
    """按「编号 / 优先级 / 约束 / 描述」顺序改写复制出来的四段。"""
    for p, (head, body) in zip(paras, ((num, "\u3000" + title), (None, prio),
                                       (None, cons), (None, desc))):
        runs = p.runs
        assert len(runs) >= 2, (p.text[:40], len(runs))
        if head is not None:
            runs[0].text = head
            runs[1].text = body
        else:
            runs[1].text = body
        # 描述段正文常被英文词切成多 run，只留前两个
        for r in runs[2:]:
            r._element.getparent().remove(r._element)


ITEMS = [
    dict(
        anchor="FR-GEN-13",
        num="FR-GEN-34", title="交易与交易链路说明", prio="必须",
        cons="批处理作业链由 FR-GEN-11 覆盖，本条只补联机交易；同一业务既有作业又有交易时，"
             "两份产物各自成文并互相引用，不合并为一份。",
        desc="交易应按交易码逐个成文，内容包含交易码与入口程序、从入口到末端的调用链、"
             "链上各程序读写的表。交易码到入口程序的绑定取自资源定义；资源定义不可得时取自"
             "程序内的交易码分支，仅由委托方人工整理的交易清单属文档主张，不替代绑定。"
             "调用链与表访问只取结构事实，源码与资源定义中均未出现的调用不得补入；"
             "一个交易码存在多个入口时按入口分列。缺少整链产物时，使用者问「这笔交易走哪些程序、"
             "动了哪些表」只能得到单个程序的说明，跨程序的部分要靠人逐个翻。",
    ),
    dict(
        anchor="FR-STO-22",
        num="FR-STO-30", title="关系链路问题的检索覆盖", prio="必须",
        cons="本条约束候选集的覆盖范围，与 FR-STO-21 的版本与应用系统限定叠加生效，不替代它；"
             "闭包深度按问题涉及的关系类型确定，不受 FR-GRP-02 浏览视图 2 跳的限制。",
        desc="关系链路类问题的检索结果，应覆盖该问题所涉对象在图谱中的可达闭包；"
             "闭包内对象在检索库中没有对应切片的，应答中应显式标注该对象缺切片，不得静默略过。"
             "仅按相似度取前若干切片、以命中条数替代覆盖的做法不满足本要求——"
             "链路中某一段的措辞与问题本就不相似，排序永远把它排在尾部，把条数调大带进来的是噪声。",
    ),
    dict(
        anchor="FR-QA-09",
        num="FR-QA-23", title="命中源码时的知识取用", prio="必须",
        cons="本条只改变解释的取用来源，不改变 FR-QA-08 对事实类结论的来源限制。",
        desc="检索命中源码切片时，应按切片元数据中的对象编号取该对象已生成的知识节作为解释，"
             "不得由模型对命中的源码现场生成说明。取用落点按对象类型确定：程序取其知识文档，"
             "段落取所属程序的流程与业务规则节，字段与表取数据实体说明，交易取交易说明。"
             "该对象尚未成文时，只呈现源码原文并标注「该对象未成文」，不补写解释。"
             "取回的知识节保留原有的来源级别与内容状态，不因被引用而改变。"
             "现场生成的内容不落库、不可复现，同一问题两次提问会给出不同措辞，"
             "与 FR-QA-05、FR-QA-06 的稳定要求冲突。",
    ),
]


def main():
    doc = Document(DOC)
    items = body_items(doc)
    tpl = item_block(items, "FR-STO-24")
    assert len(tpl) == 4, [p.text[:20] for p in tpl]

    for cfg in ITEMS:
        anchor = find_para(items, cfg["anchor"])
        new_ps = []
        for p in tpl:
            el = deepcopy(p._p)
            anchor._p.addprevious(el)
            new_ps.append(Paragraph(el, anchor._parent))
        rewrite(new_ps, cfg["num"], cfg["title"], cfg["prio"], cfg["cons"], cfg["desc"])
        print("inserted", cfg["num"], "->", cfg["anchor"])

    doc.save(DOC)

    # 回读校验
    doc = Document(DOC)
    ps = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for key in ("FR-GEN-34", "FR-STO-30", "FR-QA-23"):
        i = next(k for k, t in enumerate(ps) if t.startswith(key))
        print("\n====", ps[i])
        for t in ps[i + 1:i + 5]:
            if re.match(r"^(FR|NFR)-", t):
                break
            print("   ", t[:120])


if __name__ == "__main__":
    main()
