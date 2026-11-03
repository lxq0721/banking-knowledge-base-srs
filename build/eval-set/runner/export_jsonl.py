# -*- coding: utf-8 -*-
"""
把《银行知识库评测集模板》导出成脚本可执行的 eval_set.jsonl。

设计原则
  1. 按表头文字取列，不按列号——表里加减列、调整顺序都不会读错。
  2. 导出前先校验，问题一次列全，不边跑边发现。
  3. 只导出字段，不改原表；示例行标 is_sample，跑批默认跳过。

用法
  python export_jsonl.py 银行知识库评测集模板_v0.2.xlsx
  python export_jsonl.py <xlsx> -o eval_set.jsonl --keep-sample
"""
import argparse
import json
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

SHEET = "评测集"
HDR_ROW = 2                 # 第 1 行是分组表头，第 2 行才是字段名
FIRST_DATA_ROW = 3

LEVELS = {"代码直出", "引自文档", "人工修正", "模型生成·核验通过", "模型生成·核验未通过", "—", ""}
STATES = {"可取用", "已修正", "有争议", "已废止", "—", ""}
IN_KB = {"是", "否", "—", ""}
BASIS = {"原始资料", "开发补全物", "推断", "—", ""}
ID_RE = re.compile(r"^EV-[A-Z]{2,6}-\d{3}$")
POINT_MARKS = "①②③④⑤⑥⑦⑧⑨⑩"


def norm(v):
    """单元格值规整：None→''，数字→'12'，去首尾空白。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def split_points(text):
    """把「①…②…③…」切成要点数组；没有圈号时按换行/分号切。"""
    text = norm(text)
    if not text:
        return []
    pos = [i for i, ch in enumerate(text) if ch in POINT_MARKS]
    if len(pos) >= 2:
        out = []
        for i, p in enumerate(pos):
            end = pos[i + 1] if i + 1 < len(pos) else len(text)
            seg = text[p + 1:end].strip().strip("；;。 \n")
            if seg:
                out.append(seg)
        return out
    parts = re.split(r"[\n；;]", text)
    return [p.strip(" 。") for p in parts if p.strip(" 。")]


def split_sources(text):
    text = norm(text)
    if not text or text.startswith("—"):
        return []
    parts = re.split(r"[\n；;、]", text)
    return [p.strip() for p in parts if p.strip() and p.strip() != "—"]


def parse_rules(text):
    """解析机器判定规则：普通串=必须出现，!前缀=不得出现，/…/=正则。"""
    raw = norm(text)
    rules = {"must": [], "must_not": [], "regex": []}
    if not raw:
        return rules, []
    errs = []
    for tok in re.split(r"[|\n]", raw):
        tok = tok.strip().strip("。 ")
        if not tok:
            continue
        if tok.startswith("!") and len(tok) > 1:
            rules["must_not"].append(tok[1:].strip())
        elif tok.startswith("/"):
            if not tok.endswith("/") or len(tok) < 3:
                errs.append(f"正则未闭合：{tok}")
                continue
            try:
                re.compile(tok[1:-1])
            except re.error as e:
                errs.append(f"正则写错：{tok}（{e}）")
                continue
            rules["regex"].append(tok[1:-1])
        else:
            rules["must"].append(tok)
    if not (rules["must"] or rules["must_not"] or rules["regex"]):
        errs.append("规则列有内容但解析结果为空")
    return rules, errs


def split_scope(scope):
    """「核心存款 / 境内」→ (核心存款, 境内)。没有分隔符时 region 留空。"""
    scope = norm(scope)
    if not scope:
        return "", ""
    parts = [p.strip() for p in re.split(r"[/／]", scope) if p.strip()]
    if len(parts) >= 2:
        return parts[0], "/".join(parts[1:])
    return parts[0], ""


def load_rows(xlsx, sheet=SHEET, hdr_row=None, first_row=None):
    hdr_row = hdr_row or HDR_ROW
    first_row = first_row or FIRST_DATA_ROW
    wb = load_workbook(xlsx, data_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f"[FATAL] 找不到工作表「{sheet}」，实际有：{wb.sheetnames}")
    ws = wb[sheet]

    header = {}
    for c in range(1, ws.max_column + 1):
        name = norm(ws.cell(row=hdr_row, column=c).value)
        if name:
            header[name] = c
    need = ["编号", "所属范围", "问题原文", "问题类型", "难度", "需分角色执行",
            "库内是否存在该信息", "事实基准", "期望答案要点", "期望出处",
            "期望来源等级", "期望内容状态", "重点校验项", "机器判定规则", "备注"]
    missing = [n for n in need if n not in header]
    if missing:
        sys.exit(f"[FATAL] 表头缺字段：{missing}\n实际表头：{list(header)}")

    rows = []
    for r in range(first_row, ws.max_row + 1):
        def cell(name):
            return norm(ws.cell(row=r, column=header[name]).value)

        rid = cell("编号")
        if not rid or not ID_RE.match(rid):
            continue          # 跳过空行、分组标题行、末尾统计块
        rows.append((r, {n: cell(n) for n in header}))
    return rows


def build_record(row_no, cells):
    errs = []
    rid = cells["编号"]
    qtype = cells["问题类型"]
    type_no = qtype[0] if qtype and qtype[0].isdigit() else ""
    subsystem, region = split_scope(cells["所属范围"])
    points = split_points(cells["期望答案要点"])
    rules, rule_errs = parse_rules(cells["机器判定规则"])
    errs += [f"第 {row_no} 行 {rid}：{e}" for e in rule_errs]

    if not ID_RE.match(rid):
        errs.append(f"第 {row_no} 行：编号「{rid}」不符合 EV-<域>-<三位序号>")
    if not type_no:
        errs.append(f"第 {row_no} 行 {rid}：问题类型没选（应为 1~9 类之一）")
    if not cells["问题原文"] and type_no != "6":
        errs.append(f"第 {row_no} 行 {rid}：问题原文为空（仅第 6 类异常题可空）")
    if cells["库内是否存在该信息"] not in IN_KB:
        errs.append(f"第 {row_no} 行 {rid}：库内是否存在该信息取值非法")
    if cells["事实基准"] not in BASIS:
        errs.append(f"第 {row_no} 行 {rid}：事实基准取值非法")
    if cells["期望来源等级"] not in LEVELS:
        errs.append(f"第 {row_no} 行 {rid}：期望来源等级取值非法「{cells['期望来源等级']}」")
    if cells["期望内容状态"] not in STATES:
        errs.append(f"第 {row_no} 行 {rid}：期望内容状态取值非法「{cells['期望内容状态']}」")
    if cells["库内是否存在该信息"] == "是" and len(points) < 2:
        errs.append(f"第 {row_no} 行 {rid}：期望答案要点不足 2 条，无法按点打分")
    if cells["库内是否存在该信息"] == "是" and not split_sources(cells["期望出处"]):
        errs.append(f"第 {row_no} 行 {rid}：期望出处为空")

    note = cells["备注"]
    rec = {
        "id": rid,
        "scope": cells["所属范围"],
        "subsystem": subsystem,
        "region": region,
        "question": cells["问题原文"],
        "type": qtype,
        "type_no": type_no,
        "difficulty": cells["难度"],
        "needs_role": cells["需分角色执行"] == "是",
        "in_kb": cells["库内是否存在该信息"],
        "basis": cells["事实基准"],
        "points": points,
        "expect_sources": split_sources(cells["期望出处"]),
        "expect_level": cells["期望来源等级"],
        "expect_state": cells["期望内容状态"],
        "checks": cells["重点校验项"],
        "rules": rules,
        "judge_mode": "auto" if (rules["must"] or rules["must_not"] or rules["regex"]) else "manual",
        "note": note,
        "category": cells.get("分类", ""),
        "is_sample": note.strip().startswith("示例"),   # 只认「示例」开头的备注，避免正文里出现「示例」二字被误判
        "source_row": row_no,
    }
    return rec, errs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="评测集 xlsx 路径")
    ap.add_argument("-o", "--out", default="eval_set.jsonl")
    ap.add_argument("-s", "--sheet", default=SHEET,
                    help='工作表名，多个用逗号隔开（如 "01 开发与迁移,02 业务分析"），默认「评测集」')
    ap.add_argument("--hdr-row", type=int, default=None,
                    help="字段名所在行号（默认 2）")
    ap.add_argument("--first-row", type=int, default=None,
                    help="第一条数据行号（默认 3）")
    ap.add_argument("--keep-sample", action="store_true", help="把示例行也导出（默认导出）")
    ap.add_argument("--drop-sample", action="store_true", help="导出时剔除示例行")
    args = ap.parse_args()

    sheets = [s.strip() for s in args.sheet.split(",") if s.strip()]
    rows = []
    for sh in sheets:
        got = load_rows(args.xlsx, sh, args.hdr_row, args.first_row)
        if not got:
            print(f"[WARN] 工作表「{sh}」没有任何一行填了编号，已跳过")
            continue
        rows += [(sh, rn, cells) for rn, cells in got]
    if not rows:
        sys.exit("[FATAL] 指定的工作表里没有任何一行填了编号")

    multi = len(sheets) > 1
    records, errors = [], []
    seen = {}
    for sh, row_no, cells in rows:
        rid = cells["编号"]
        where = f"[{sh}] 第 {row_no} 行" if multi else f"第 {row_no} 行"
        if rid in seen:
            errors.append(f"{where}：编号 {rid} 与 {seen[rid]} 重复")
        seen[rid] = f"{sh} 第 {row_no} 行" if multi else f"第 {row_no} 行"
        rec, errs = build_record(row_no, cells)
        rec["view"] = sh
        rec["source_sheet"] = sh
        errors += [e.replace(f"第 {row_no} 行", where) for e in errs]
        records.append(rec)

    if args.drop_sample:
        before = len(records)
        records = [r for r in records if not r["is_sample"]]
        print(f"[INFO] 已剔除示例行 {before - len(records)} 条")

    out = Path(args.out)
    with out.open("w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

    auto = sum(1 for r in records if r["judge_mode"] == "auto")
    print(f"[OK] 导出 {len(records)} 条 → {out}")
    print(f"     可自动判定 {auto} 条 / 需人工 {len(records) - auto} 条 "
          f"（自动率 {auto / len(records):.0%}）")
    print(f"     示例行 {sum(1 for r in records if r['is_sample'])} 条（跑批默认跳过）")

    if errors:
        print(f"\n[校验] 发现 {len(errors)} 处问题，建议改表后重新导出：")
        for e in errors:
            print("   -", e)
        sys.exit(1)
    print("\n[校验] 通过：编号唯一、必填齐全、要点可拆、规则语法正确")


if __name__ == "__main__":
    main()
