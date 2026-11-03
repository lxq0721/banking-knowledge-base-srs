# -*- coding: utf-8 -*-
"""
判定器用例自检：八种情形各一条，确认 judge() 的每个分支都按预期落。

  python selftest_judge.py      # 全过退出码 0
"""
import sys

import run_eval as R

BASE = {
    "id": "T-1",
    "in_kb": "是",
    "expect_level": "代码直出",
    "expect_state": "可取用",
    "rules": {"must": ["ACCT-BAL"], "must_not": ["不支持"], "regex": []},
    "expect_sources": ["a.cpy:1"],
    "checks": "",
}

NO_RULE = {"must": [], "must_not": [], "regex": []}

CASES = [
    ("全部命中", dict(BASE), "ACCT-BAL 定义在 a.cpy", "通过"),
    ("一条都没命中", dict(BASE), "不知道", "不通过"),
    ("出现禁词", dict(BASE), "ACCT-BAL 支持，但暂不支持该场景", "不通过"),
    ("无规则走人工", dict(BASE, rules=NO_RULE), "随便说说", "待人工"),
    ("幻觉未声明", dict(BASE, in_kb="否", rules={"must": [], "must_not": ["不支持"], "regex": []}),
     "应该可以办理，流程如下", "不通过"),
    ("幻觉正确拒答", dict(BASE, in_kb="否", rules={"must": [], "must_not": ["不支持"], "regex": []}),
     "知识库中未查询到相关信息", "通过"),
    ("要求标注却漏标", dict(BASE, expect_level="模型生成·核验通过"), "ACCT-BAL 定义在 a.cpy", "不通过"),
    ("标注齐全", dict(BASE, expect_level="模型生成·核验通过"),
     "ACCT-BAL 定义在 a.cpy（模型生成·未核实）", "通过"),
    ("正则命中", dict(BASE, rules={"must": [], "must_not": [], "regex": [r"a\.cpy:\d+"]}),
     "见 a.cpy:214", "通过"),
    ("正则未命中", dict(BASE, rules={"must": [], "must_not": [], "regex": [r"a\.cpy:\d+"]}),
     "见 a.cpy", "不通过"),
    ("仅禁词规则且未出现", dict(BASE, rules={"must": [], "must_not": ["不支持"], "regex": []}),
     "ACCT-BAL 支持", "通过"),
]


def main():
    bad = 0
    for name, rec, ans, want in CASES:
        got = R.judge(rec, ans, [], 200, 900)["verdict"]
        ok = got == want
        bad += 0 if ok else 1
        print(("PASS  " if ok else "FAIL  ") + f"{name:14s} 期望 {want:<5s} 实际 {got}")
    print(f"\n{len(CASES)} 条用例，通过 {len(CASES) - bad} 条")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
