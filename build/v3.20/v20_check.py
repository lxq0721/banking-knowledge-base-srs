# -*- coding: utf-8 -*-
"""v3.20 一致性校验：条目总数、分域条数、四张汇总表的合计与逐域对齐、新条目段落结构。"""
import os
import re
import collections

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")

TOTAL, MUST, SHOULD, OPT = 144, 111, 32, 1
OK, BAD = [], []


def count_ids(s):
    """展开区间与省略前缀的续号：FR-PRS-01~02、04~05 计 4 条；
    FR-QA-07~09、22~23 中的「22~23」继承前一个前缀，计 2 条。"""
    n = 0
    pre = None
    for part in re.split(r"[、，,]", s):
        part = part.strip()
        m = re.match(r"^((?:FR-[A-Z]{2,3}|NFR)-)(\d+)(?:~(\d+))?$", part)
        if m:
            pre = m.group(1)
            a = int(m.group(2))
            b = int(m.group(3)) if m.group(3) else a
            n += b - a + 1
            continue
        m = re.match(r"^(\d+)(?:~(\d+))?$", part)
        if m and pre:
            a = int(m.group(1))
            b = int(m.group(2)) if m.group(2) else a
            n += b - a + 1
    return n


def chk(cond, msg):
    (OK if cond else BAD).append(msg)


def items_of(doc):
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            out.append(Paragraph(ch, doc))
        elif ch.tag.endswith("}tbl"):
            out.append(Table(ch, doc))
    return out


def table_by_caption(doc, cap):
    last = ""
    for obj in items_of(doc):
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
    ps = [p for p in doc.paragraphs if p.text.strip()]

    # 条目与优先级
    ids = []
    for p in ps:
        # NFR 只有两位编号（NFR-01），不能走 [A-Z]{2,3}- 这条
        m = re.match(r"^((?:FR-[A-Z]{2,3}-\d+)|(?:NFR-\d+))", p.text.strip())
        if m:
            ids.append(m.group(1))
    chk(len(ids) == TOTAL, "条目总数 = %d（期望 %d）" % (len(ids), TOTAL))
    chk(len(set(ids)) == len(ids), "编号无重复")

    prio = collections.Counter()
    cur = None
    for p in ps:
        t = p.text.strip()
        if re.match(r"^((?:FR|NFR)-)", t):
            cur = t.split("\u3000")[0]
        elif t.startswith("优先级：") and cur:
            prio[t.replace("优先级：", "").strip()] += 1
            cur = None
    chk(prio["必须"] == MUST, "必须 %d（期望 %d）" % (prio["必须"], MUST))
    chk(prio["应该"] == SHOULD, "应该 %d（期望 %d）" % (prio["应该"], SHOULD))
    chk(prio["可选"] == OPT, "可选 %d（期望 %d）" % (prio["可选"], OPT))

    # 分域条数
    dom = collections.Counter()
    for x in ids:
        m = re.match(r"^FR-([A-Z]{2,3})-", x)
        dom[m.group(1) if m else "NFR"] += 1
    print("分域:", dict(dom))
    chk(dom["STO"] == 27, "STO 域 %d 条（期望 27）" % dom["STO"])
    chk(dom["QA"] == 22, "QA 域 %d 条（期望 22）" % dom["QA"])

    # 表 1-3
    tb = table_by_caption(doc, "表 1-3")
    got = {r.cells[0].text.strip(): r.cells[2].text.strip() for r in tb.rows[1:]}
    chk(str(MUST) in got["必须"], "表 1-3 必须行 = %s" % got["必须"])

    # 表 3-2 / 表 B-1：逐域合计与表格合计、编号串展开后等于条数
    for cap, c_num, c_cnt in (("表 3-2", 2, 3), ("表 B-1", 3, 4)):
        tb = table_by_caption(doc, cap)
        by_dom = collections.OrderedDict()
        cur_dom = ""
        for r in tb.rows[1:]:
            name = r.cells[0].text.strip()
            if name:
                cur_dom = name
            by_dom.setdefault(cur_dom, 0)
            by_dom[cur_dom] += int(r.cells[c_cnt].text.strip())
        total = sum(by_dom.values())
        chk(total == TOTAL, "%s 合计 = %d（期望 %d）" % (cap, total, TOTAL))
        norm = lambda x: re.sub(r"^\d+\s*", "", x).strip()
        by_dom = {norm(k): v for k, v in by_dom.items()}
        expect = {"资料导入": dom["IMP"], "内容解析": dom["PRS"], "知识生成": dom["GEN"],
                  "质量核验": dom["REV"], "入库建图": dom["STO"], "图谱浏览": dom["GRP"],
                  "问答应答": dom["QA"], "数据需求": dom["DAT"], "非功能需求": dom["NFR"]}
        bad = [(k, by_dom.get(k), v) for k, v in expect.items() if by_dom.get(k) != v]
        chk(not bad, "%s 逐域对齐%s" % (cap, "" if not bad else " 不符: %s" % bad))
        for r in tb.rows[1:]:
            n = count_ids(r.cells[c_num].text)
            cnt = int(r.cells[c_cnt].text.strip())
            if n != cnt:
                chk(False, "%s 行「%s」按编号计 %d 条但条数列为 %d" %
                    (cap, r.cells[1].text.strip(), n, cnt))

    # 表 A-1
    tb = table_by_caption(doc, "表 A-1")
    p_sum = 0
    for r in tb.rows[1:]:
        m = re.match(r"(\d+) 条", r.cells[2].text.strip())
        p_sum += int(m.group(1))
    chk(p_sum == TOTAL, "表 A-1 三批合计 = %d（期望 %d）" % (p_sum, TOTAL))

    # 正文残留旧数：附录 C 之前才算残留
    cut = next(k for k, p in enumerate(ps) if p.style.name == "Heading 1"
               and p.text.strip().startswith("附录 C"))
    for p in ps[:cut]:
        for bad_num in ("141 条", "107 条"):
            if bad_num in p.text:
                chk(False, "正文残留「%s」: %s" % (bad_num, p.text[:60]))

    # 三条新条目的段落结构：编号 / 优先级 / 约束 / 描述
    # 组尾条目后面紧跟下一个组的标题，标题不算条目段
    for code in ("FR-STO-31", "FR-STO-32", "FR-QA-24"):
        i = next(k for k, p in enumerate(ps) if p.text.strip().startswith(code))
        seg = []
        for p in ps[i + 1:i + 6]:
            t = p.text.strip()
            if re.match(r"^(FR|NFR)-", t) or p.style.name.startswith("Heading"):
                break
            seg.append(t)
        chk(len(seg) == 3 and seg[0].startswith("优先级：") and seg[1].startswith("约束：")
            and seg[2].startswith("描述："), "%s 四段结构完整" % code)

    # 表 8-5：作业步骤节点 + 两类出处列
    tb = table_by_caption(doc, "表 8-5")
    node = next(r for r in tb.rows if r.cells[0].text.strip() == "节点类型")
    rel = next(r for r in tb.rows if r.cells[0].text.strip() == "关系类型")
    chk("作业步骤" in node.cells[1].text, "表 8-5 节点含作业步骤")
    chk("交易" in node.cells[1].text, "表 8-5 节点含交易")
    chk(all(k in node.cells[2].text for k in ("业务术语", "知识条目", "应用系统")),
        "表 8-5 节点出处补齐三类")
    chk(all(k in rel.cells[2].text for k in ("包含", "定义", "归属", "别名")),
        "表 8-5 关系出处补齐四类")
    chk("触发" in rel.cells[1].text, "表 8-5 关系含触发")

    # 版本号三处
    chk(any(p.text.strip() == "v3.20" for p in ps), "封面标题下版本段 = v3.20")
    hdr = [p.text for s in doc.sections for p in s.header.paragraphs]
    chk(any("v3.20" in t for t in hdr), "页眉版本 = v3.20")

    print()
    for m in OK:
        print("  [OK] ", m)
    for m in BAD:
        print("  [!!] ", m)
    print("\n通过 %d 项，失败 %d 项" % (len(OK), len(BAD)))


if __name__ == "__main__":
    main()
