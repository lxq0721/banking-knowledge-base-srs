# -*- coding: utf-8 -*-
"""v3.19 一致性校验：条目总数、分域条数、四张汇总表的合计与逐域对齐。"""
import os
import re
import collections

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19.docx")

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
        m = re.match(r"^((?:FR|NFR)-[A-Z]{2,3}-\d+|NFR-\d+)", p.text.strip())
        if m:
            ids.append(m.group(1))
    chk(len(ids) == 141, "条目总数 = %d（期望 141）" % len(ids))
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
    chk(prio["必须"] == 108, "必须 %d（期望 108）" % prio["必须"])
    chk(prio["应该"] == 32, "应该 %d（期望 32）" % prio["应该"])
    chk(prio["可选"] == 1, "可选 %d（期望 1）" % prio["可选"])

    # 分域条数
    dom = collections.Counter()
    for x in ids:
        m = re.match(r"^FR-([A-Z]{2,3})-", x)
        dom[m.group(1) if m else "NFR"] += 1
    print("分域:", dict(dom))

    # 表 1-3
    tb = table_by_caption(doc, "表 1-3")
    got = {r.cells[0].text.strip(): r.cells[2].text.strip() for r in tb.rows[1:]}
    chk("108" in got["必须"], "表 1-3 必须行 = %s" % got["必须"])
    chk("32" in got["应该"], "表 1-3 应该行 = %s" % got["应该"])

    # 表 3-2 / 表 B-1：逐域合计与表格合计
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
        chk(total == 141, "%s 合计 = %d（期望 141）" % (cap, total))
        # 逐域比对
        norm = lambda x: re.sub(r"^\d+\s*", "", x).strip()
        by_dom = {norm(k): v for k, v in by_dom.items()}
        expect = {"资料导入": dom["IMP"], "内容解析": dom["PRS"], "知识生成": dom["GEN"],
                  "质量核验": dom["REV"], "入库建图": dom["STO"], "图谱浏览": dom["GRP"],
                  "问答应答": dom["QA"], "数据需求": dom["DAT"], "非功能需求": dom["NFR"]}
        bad = [(k, by_dom.get(k), v) for k, v in expect.items() if by_dom.get(k) != v]
        chk(not bad, "%s 逐域对齐%s" % (cap, "" if not bad else " 不符: %s" % bad))
        # 编号串展开区间后应等于条数
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
    chk(p_sum == 141, "表 A-1 三批合计 = %d（期望 141）" % p_sum)

    # 正文残留 138：附录 C 之前才算残留（附录 C 里的历史记录本应保留旧数）
    cut = next(k for k, p in enumerate(ps) if p.style.name == "Heading 1"
               and p.text.strip().startswith("附录 C"))
    for p in ps[:cut]:
        if "138 条" in p.text:
            chk(False, "正文残留「138 条」: %s" % p.text[:60])

    # 表 6-5 / 表 8-5 新行
    tb = table_by_caption(doc, "表 6-5")
    chk(any("交易入口与调用链" in r.cells[0].text for r in tb.rows), "表 6-5 已补交易说明行")
    tb = table_by_caption(doc, "表 8-5")
    node = next(r for r in tb.rows if r.cells[0].text.strip() == "节点类型")
    rel = next(r for r in tb.rows if r.cells[0].text.strip() == "关系类型")
    chk("交易" in node.cells[1].text, "表 8-5 节点含交易")
    chk("触发" in rel.cells[1].text, "表 8-5 关系含触发")

    print()
    for m in OK:
        print("  [OK] ", m)
    for m in BAD:
        print("  [!!] ", m)
    print("\n通过 %d 项，失败 %d 项" % (len(OK), len(BAD)))


if __name__ == "__main__":
    main()
