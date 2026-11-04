# -*- coding: utf-8 -*-
"""把 v2.0 派生的两份工作簿导出为 jsonl。

    py build/eval-set/export_derived.py

输出：
    eval-set/jsonl/eval_basic.jsonl      基础版（人工逐条执行）
    eval-set/jsonl/eval_deepeval.jsonl   DeepEval 版（框架字段）

为什么不用同目录的 runner/export_jsonl.py：那个脚本按 v2.0 的 18 列表头取字段，
而这两份为了给人看改过列名（「问题」而不是「问题原文」、`input(问题)` 这类带
后缀的框架字段），直接喂进去会报「表头缺字段」。映射关系写死在本脚本里，改列
名时同步改下面的 BASIC_MAP / DEEPEVAL_MAP。

值变换只有三处，都是表里存不下、导出时必须定的：
  1. 期望要点在表里是「一行一条」，基础版补 `1）2）3）` 序号，DeepEval 版用
     「；」连成一段——`expected_output` 是单个字符串，塞不了列表。
  2. 期望出处同样按行拆成列表。
  3. 「是否需角色」表里写是/否，jsonl 转布尔；「自研判据」的
     `must:X | !Y | /re/` 解析成 {must, must_not, regex}，全空则判 manual。
"""

import json
import os
import re
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, "eval-set", "银行知识库评测集_基础版_v1.0.xlsx")
SRC_DE = os.path.join(ROOT, "eval-set", "银行知识库评测集_DeepEval版_v1.0.xlsx")
OUT_DIR = os.path.join(ROOT, "eval-set", "jsonl")

BASIC_MAP = {
    "id": "编号",
    "category": "分类",
    "type": "题型",
    "difficulty": "难度",
    "question": "问题",
    "points": "期望要点",
    "expect_sources": "期望出处",
    "expect_level": "来源级别",
    "expect_state": "内容状态",
    "in_kb": "有无答案",
    "needs_role": "是否需角色",
    "note": "备注",
}

DEEPEVAL_MAP = {
    "id": "编号",
    "input": "input(问题)",
    "expected_output": "expected_output(期望要点)",
    "expect_sources": "期望出处",
    "expect_level": "来源级别",
    "expect_state": "内容状态",
    "in_kb": "有无答案",
    "category": "分类",
    "type": "题型",
    "difficulty": "难度",
    "needs_role": "是否需角色",
    "rules": "自研判据 rules",
}


def lines(v):
    """单元格按行拆成列表，去掉空行与首尾空白。"""
    if v is None:
        return []
    return [s.strip() for s in str(v).splitlines() if s.strip()]


def yes_no(v, default=False):
    if v is None:
        return default
    return str(v).strip() == "是"


def parse_rules(v):
    """`must:X | !Y | /re/` → {'must': [...], 'must_not': [...], 'regex': [...]}"""
    out = {"must": [], "must_not": [], "regex": []}
    for tok in str(v or "").split("|"):
        tok = tok.strip()
        if not tok:
            continue
        if tok.startswith("must_not:"):
            out["must_not"].append(tok[len("must_not:"):].strip())
        elif tok.startswith("must:"):
            out["must"].append(tok[len("must:"):].strip())
        elif tok.startswith("!"):
            out["must_not"].append(tok[1:].strip())
        elif tok.startswith("/") and tok.endswith("/") and len(tok) > 2:
            out["regex"].append(tok[1:-1])
        else:
            sys.exit("[FATAL] 认不出判定 token：%r" % tok)
    return out


def read_sheet(path, sheet_marker):
    """按工作表名关键字取表；表头固定在第 1 行，数据从第 2 行起。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    names = [n for n in wb.sheetnames if sheet_marker in n]
    if len(names) != 1:
        sys.exit("[FATAL] %s 里含「%s」的工作表有 %d 个：%s"
                 % (os.path.basename(path), sheet_marker, len(names), wb.sheetnames))
    ws = wb[names[0]]
    header = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        rows.append(dict(zip(header, r)))
    return header, rows


def need(header, mapping, path):
    missing = [v for v in mapping.values() if v not in header]
    if missing:
        sys.exit("[FATAL] %s 缺列：%s\n实际表头：%s"
                 % (os.path.basename(path), missing, header))


def dump(records, name):
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, name)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        for rec in records:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    print("  %-22s %3d 条  %6.1f KB" % (name, len(records), os.path.getsize(out) / 1024))


def main():
    # ── 基础版：人工执行用，字段与服务端判分无关的一律不带 ──
    header, rows = read_sheet(SRC, "手动执行")
    need(header, BASIC_MAP, SRC)
    basic = []
    for r in rows:
        g = lambda k: r.get(BASIC_MAP[k])
        basic.append({
            "id": g("id"),
            "category": g("category"),
            "type": g("type"),
            "difficulty": g("difficulty"),
            "question": g("question"),
            # 表里一行一条要点，导出补序号，人工对着打分时好定位
            "points": ["%d）%s" % (i + 1, s) for i, s in enumerate(lines(g("points")))],
            "expect_sources": lines(g("expect_sources")),
            "expect_level": g("expect_level"),
            "expect_state": g("expect_state"),
            "in_kb": g("in_kb"),
            "needs_role": yes_no(g("needs_role")),
            "note": g("note") or "",
        })

    # ── DeepEval 版：字段名即框架字段，运行时两列留空待填 ──
    header_de, rows_de = read_sheet(SRC_DE, "DeepEval")
    need(header_de, DEEPEVAL_MAP, SRC_DE)
    deepeval = []
    for r in rows_de:
        g = lambda k: r.get(DEEPEVAL_MAP[k])
        rules = parse_rules(g("rules"))
        deepeval.append({
            "id": g("id"),
            "input": g("input"),
            "expected_output": "；".join(lines(g("expected_output"))),
            "actual_output": "",       # 运行时由被测接口填
            "retrieval_context": [],   # 运行时由检索侧填
            "in_kb": g("in_kb"),
            "expect_sources": lines(g("expect_sources")),
            "expect_level": g("expect_level"),
            "expect_state": g("expect_state"),
            "category": g("category"),
            "type": g("type"),
            "difficulty": g("difficulty"),
            "needs_role": yes_no(g("needs_role")),
            "rules": rules,
            # 一条判据都定不出来的只能人工判（v2.0 里是空输入、越权索取、指令注入）
            "judge_mode": "auto" if any(rules.values()) else "manual",
        })

    if len(basic) != len(deepeval):
        sys.exit("[FATAL] 两份条数不一致：基础版 %d / DeepEval %d"
                 % (len(basic), len(deepeval)))
    mismatch = [a["id"] for a, b in zip(basic, deepeval) if a["id"] != b["id"]]
    if mismatch:
        sys.exit("[FATAL] 两份编号对不上：%s" % mismatch[:10])

    print("导出到 eval-set/jsonl/：")
    dump(basic, "eval_basic.jsonl")
    dump(deepeval, "eval_deepeval.jsonl")
    manual = sum(1 for x in deepeval if x["judge_mode"] == "manual")
    print("  %d 条；机判 %d / 人工判 %d" % (len(basic), len(basic) - manual, manual))


if __name__ == "__main__":
    main()
