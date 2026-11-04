# -*- coding: utf-8 -*-
"""v3.20 第一步：插入三条新条目。

FR-STO-31  边的节点粒度            8.3.4 图谱建立，组尾（锚点为 Heading「存储分工与准入」）
FR-STO-32  可达闭包的遍历方向      8.3.4 图谱建立，紧跟 FR-STO-31
FR-QA-24   问题中的对象名锚定      10.3.4 来源约束，组尾（锚点为 Heading「经验沉淀」）

组尾插入的锚点用标题而非条目编号：addprevious 总是紧贴锚点之前，
两条都锚到同一标题、按编号升序执行，落出来就是 31、32、标题。

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
DOC = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")


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


def find_heading(items, text, level):
    """按标题文本精确匹配，避开同名正文与目录行。"""
    for p in items:
        if not isinstance(p, Paragraph):
            continue
        if p.style.name == "Heading %d" % level and p.text.strip() == text:
            return p
    raise KeyError("Heading %d / %s" % (level, text))


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
        anchor=("heading", "存储分工与准入", 3),
        num="FR-STO-31", title="边的节点粒度", prio="必须",
        cons="负向需求：以「一个程序只落一条读写边」作为反例验收；程序级与作业级关系由"
             "下级边聚合得出，不单独落边。",
        desc="读写、调用、触发边应挂在能够定位到的最细节点上：程序内挂段落节点，作业内挂"
             "作业步骤，交易挂交易节点；段落与作业步骤分别以「包含」关系连到所属程序与作业。"
             "一条边对应它所能落到的那处语句，同一程序内对同一文件的多处访问应落为多条段落边，"
             "不得合并为一条程序级边。程序级「谁用过什么」由段落边沿「包含」向上聚合，不另存一份。"
             "证据位置按 FR-STO-06 落到行或 STEP；只能定位到程序整体、无法落到段落或步骤的，"
             "允许落程序级边并在证据位置注明「仅到程序」。",
    ),
    dict(
        anchor=("heading", "存储分工与准入", 3),
        num="FR-STO-32", title="可达闭包的遍历方向", prio="必须",
        cons="与 FR-STO-30 配套生效，不改变边本身的语义方向；以「从文件节点出发取不到"
             "任何对象」为反例验收。",
        desc="可达闭包应逐类关系定义可遍历方向，不得一律沿出边。「调用」与「读写」须支持"
             "反向：由被访问的文件或表反查访问它的程序与段落，由被调用程序反查调用者；"
             "「包含」沿出边，由容器到其内部节点；「归属」只用于限定应用系统，不作为链路扩展；"
             "「依据」与「引用」不参与闭包扩展。反向得到的路径在应答中应标注方向并说明"
             "「由被访问方反查得到」，不得展示为正向调用顺序——先后顺序只能由成文的交易说明"
             "或调用点的行号给出。",
    ),
    dict(
        anchor=("heading", "经验沉淀", 3),
        num="FR-QA-24", title="问题中的对象名锚定", prio="必须",
        cons="只改变检索入口的确定方式，不改变 FR-QA-08 对事实类结论的来源限制，"
             "也不放宽 FR-STO-21 的版本与应用系统限定。",
        desc="问题中出现可按唯一键匹配的对象名（程序名、文件与数据集名、表名、字段名、交易码、"
             "作业名）时，应先以该节点为检索入口，再在 FR-STO-30 的闭包内按相似度排序；"
             "不得仅由相似度决定入口。匹配按 FR-STO-24 的唯一键在问题所指应用系统内进行；"
             "命中多个同名对象时并列呈现并标注所属应用系统，不得任选其一。问题所指对象未能匹配时，"
             "退回按相似度检索，并在回答中说明未匹配到该对象，不得静默改按措辞取候选。",
    ),
]


def main():
    doc = Document(DOC)
    items = body_items(doc)
    tpl = item_block(items, "FR-STO-24")
    assert len(tpl) == 4, [p.text[:20] for p in tpl]

    for cfg in ITEMS:
        kind, text, lvl = cfg["anchor"]
        anchor = find_heading(items, text, lvl)
        new_ps = []
        for p in tpl:
            el = deepcopy(p._p)
            anchor._p.addprevious(el)
            new_ps.append(Paragraph(el, anchor._parent))
        rewrite(new_ps, cfg["num"], cfg["title"], cfg["prio"], cfg["cons"], cfg["desc"])
        print("inserted", cfg["num"], "-> 标题「%s」前" % text)

    doc.save(DOC)

    # 回读校验
    doc = Document(DOC)
    ps = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for key in ("FR-STO-31", "FR-STO-32", "FR-QA-24"):
        i = next(k for k, t in enumerate(ps) if t.startswith(key))
        print("\n====", ps[i])
        for t in ps[i + 1:i + 5]:
            if re.match(r"^(FR|NFR)-", t):
                break
            print("   ", t[:120])
    print("\n条数", len([t for t in ps if re.match(r"^((?:FR|NFR)-[A-Z]{2,3}-\d+)", t)]))


if __name__ == "__main__":
    main()
