# -*- coding: utf-8 -*-
"""
按 eval_set.jsonl 跑批并判定，输出 results.csv。

判定顺序（不可颠倒）
  1 硬失败优先：5xx / 连接失败 / 超时 → 不通过，不再看内容
  2 无答案题专项：该说「未查询到」而没说，或给出了结论 → 不通过（幻觉，P0）
  3 级别标注：要求标「模型生成」的题，回答里没有该标注 → 不通过（标注缺失）
  4 规则匹配：must 全中 / must_not 全不中 / regex 全命中 → 通过；命中一部分 → 部分通过
  5 无规则 → 待人工（没判据就等于没判，不能算通过）

用法
  python run_eval.py                         # dry-run，用内置桩把判定逻辑走一遍
  python run_eval.py --endpoint http://host/qa
  python run_eval.py --endpoint ... --roles developer,ba,architect --include-sample
"""
import argparse
import csv
import hashlib
import json
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

NO_ANSWER_HINTS = ("未查询到", "未找到", "没有相关", "无相关", "无法确定", "知识库中未", "暂无相关")
LABEL_HINTS = ("模型生成", "未核实", "仅供参考")
NEED_LABEL = {"模型生成·核验通过", "模型生成·核验未通过"}
FORBIDDEN_AS_BASIS = {"已废止", "有争议"}


# ---------------------------------------------------------------- 判定
def judge(rec, answer, sources, status_code, latency_ms, error=None):
    """返回 dict：verdict / mode / hit / miss / violation / reason"""
    res = {"verdict": "待人工", "mode": "auto", "hit": [], "miss": [],
           "violation": [], "reason": ""}
    text = answer or ""
    low = text.lower()

    # 1 硬失败
    if error or (status_code and status_code >= 500):
        res.update(verdict="不通过", violation=["技术失败"],
                   reason=f"5xx/连接失败：{error or status_code}")
        return res
    if status_code == 408 or (latency_ms and latency_ms >= 60000):
        res.update(verdict="不通过", violation=["超时"], reason="超过 60s")
        return res
    if not text.strip():
        res.update(verdict="不通过", violation=["空响应"], reason="answer 为空")
        return res

    rules = rec["rules"]
    must, must_not, regexes = rules["must"], rules["must_not"], rules["regex"]

    # 2 无答案题
    if rec["in_kb"] == "否":
        res["mode"] = "auto"
        said_no = any(h in text for h in NO_ANSWER_HINTS)
        if not said_no:
            res.update(verdict="不通过", violation=["幻觉：未声明未查到"],
                       reason="库内无此信息，回答却未说明「未查询到」")
            return res
        if any(t.lower() in low for t in must_not):
            res.update(verdict="不通过", violation=["幻觉：给出推测结论"],
                       reason="出现禁用结论词")
            return res
        res.update(verdict="通过", hit=["未声明编造"], reason="正确声明未查询到")
        return res

    # 3 级别标注
    if rec["expect_level"] in NEED_LABEL and not any(h in text for h in LABEL_HINTS):
        res.update(verdict="不通过", violation=["标注缺失"],
                   reason=f"应标注「{rec['expect_level']}」但回答无标注")
        return res

    # 4 规则匹配（must_not 只作否决项，不参与命中率分母）
    total = len(must) + len(regexes)
    if total == 0 and not must_not:
        res.update(verdict="待人工", mode="manual",
                   reason="该题无机器判定规则，需人工判语义质量")
        return res

    hit, miss, bad = [], [], []
    for t in must:
        (hit if t.lower() in low else miss).append(t)
    for t in must_not:
        if t.lower() in low:
            bad.append(f"出现禁词「{t}」")
    for p in regexes:
        (hit if re.search(p, text, re.I) else miss).append(f"/{p}/")

    ok = len(hit)
    res["hit"], res["miss"], res["violation"] = hit, miss, bad
    if bad:
        res.update(verdict="不通过", reason="；".join(bad))
    elif ok == total:
        res.update(verdict="通过", reason="规则全部满足")
    elif ok == 0:
        res.update(verdict="不通过", reason="规则一条都未命中")
    else:
        res.update(verdict="部分通过", reason=f"命中 {ok}/{total}")
    return res


# ---------------------------------------------------------------- 桩（dry-run）
def stub_answer(rec):
    """按编号哈希稳定地模拟「好 / 半 / 差 / 编造」四种响应，用于自测判定链路。
    这只是桩，不代表任何真实系统质量。"""
    h = int(hashlib.md5(rec["id"].encode()).hexdigest(), 16) % 4
    t = rec["type_no"]
    must = rec["rules"]["must"]
    srcs = rec["expect_sources"]

    if rec["in_kb"] == "否":
        if h == 3:
            return "支持，数字人民币钱包可以正常办理余额结转。", srcs               # 编造
        return "知识库中未查询到相关信息，当前资料不涉及该功能。", []

    if t == "6":                                    # 异常题：桩返回空响应，应判不通过
        return "", []

    base = "；".join(must) if must else rec["question"]
    if h == 0:
        txt = f"（桩）{base}。依据：{srcs[0] if srcs else '—'}。"
        if rec["expect_level"] in NEED_LABEL:
            txt += "（模型生成·未核实）"
        return txt, srcs
    if h == 1:
        keep = must[:-1] if len(must) > 1 else must
        txt = "（桩）" + "；".join(keep) + "。"
        if rec["expect_level"] in NEED_LABEL:
            txt += "（模型生成·未核实）"
        return txt, srcs
    if h == 2:
        return "（桩）未查询到相关信息。", []
    txt = f"（桩）{'；'.join(must)}。"
    if rec["expect_level"] in NEED_LABEL:           # 故意漏标注
        txt = txt.replace("（模型生成·未核实）", "")
    return txt, srcs


# ---------------------------------------------------------------- 接口
def http_call(endpoint, rec, role, timeout):
    body = {
        "question": rec["question"],
        "filters": {"subsystem": rec["subsystem"], "region": rec["region"]},
        "role": role,
    }
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(endpoint, data=data,
                                 headers={"Content-Type": "application/json"})
    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
            ms = int((time.time() - t0) * 1000)
            node = payload.get("data", payload)
            answer = node.get("answer") or node.get("content") or ""
            srcs = node.get("sources") or node.get("contexts") or []
            if srcs and isinstance(srcs[0], dict):
                srcs = [s.get("loc") or s.get("location") or s.get("file") or str(s) for s in srcs]
            return answer, list(srcs), resp.status, ms, None
    except urllib.error.HTTPError as e:
        return "", [], e.code, int((time.time() - t0) * 1000), f"HTTP {e.code}"
    except Exception as e:                                  # 连接失败 / 超时
        return "", [], None, int((time.time() - t0) * 1000), f"{type(e).__name__}: {e}"


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="eval_set.jsonl")
    ap.add_argument("--endpoint", default="", help="不给则走 dry-run 桩")
    ap.add_argument("--roles", default="", help="需分角色执行的题按此列表各跑一遍，逗号分隔")
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument("--out", default="results.csv")
    ap.add_argument("--include-sample", action="store_true")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    recs = [json.loads(l) for l in Path(args.input).read_text(encoding="utf-8").splitlines() if l.strip()]
    if not args.include_sample:
        n = len(recs)
        recs = [r for r in recs if not r["is_sample"]]
        if n != len(recs):
            print(f"[INFO] 跳过示例行 {n - len(recs)} 条（加 --include-sample 才跑）")
    if args.limit:
        recs = recs[:args.limit]
    if not recs:
        sys.exit("[FATAL] 没有可执行的题目——主表是不是只填了示例行？")

    roles = [r.strip() for r in args.roles.split(",") if r.strip()] or ["default"]
    mode = "http" if args.endpoint else "dry"

    rows, stat = [], {}
    flags = {"halluc": 0, "label": 0, "hard": 0}
    for rec in recs:
        run_roles = roles if rec["needs_role"] else ["default"]
        for role in run_roles:
            if mode == "http":
                answer, srcs, code, ms, err = http_call(args.endpoint, rec, role, args.timeout)
            else:
                answer, srcs = stub_answer(rec)
                code, ms, err = 200, 800 + int(hashlib.md5(rec["id"].encode()).hexdigest()[:3], 16), None
            j = judge(rec, answer, srcs, code, ms, err)
            stat[j["verdict"]] = stat.get(j["verdict"], 0) + 1
            if any(v.startswith("幻觉") for v in j["violation"]):
                flags["halluc"] += 1
            if "标注缺失" in j["violation"]:
                flags["label"] += 1
            if {"技术失败", "超时", "空响应"} & set(j["violation"]):
                flags["hard"] += 1

            # 错引：返回了期望之外的出处；违规项里除幻觉/标注之外的也归到这里
            wrong = [s for s in srcs if not any(s in e or e in s for e in rec["expect_sources"])]
            wrong += [v for v in j["violation"] if not v.startswith(("幻觉", "标注缺失"))]
            label_state = ("缺失" if "标注缺失" in j["violation"]
                           else "不适用" if rec["expect_level"] not in NEED_LABEL else "齐全")
            hit_src = [s for s in srcs if s not in wrong]
            rows.append({
                "执行批次": time.strftime("B%Y%m%d-%H%M"),
                "执行时间": time.strftime("%Y-%m-%d %H:%M:%S"),
                "编号": rec["id"],
                "本次角色 / 范围": f"{role} / {rec['scope']}" if rec["scope"] else role,
                "问题原文（快照）": rec["question"],
                "实际输出（摘要）": answer.replace("\n", " ")[:300],
                "命中要点": " / ".join(j["hit"]),
                "漏要点": " / ".join(j["miss"]),
                "错引 / 越界引用": " / ".join(wrong),
                "标注是否齐全": label_state,
                "耗时 ms": ms,
                "判定结果": j["verdict"],
                "判定方式": "自动" if j["mode"] == "auto" else "人工",
                "执行人": "脚本",
                "证据位置": (hit_src[0] if hit_src else "—"),
                "备注": j["reason"],
            })

    out = Path(args.out)
    with out.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)

    judged = [r for r in rows if r["判定结果"] != "待人工"]
    passed = sum(1 for r in judged if r["判定结果"] == "通过")
    partial = sum(1 for r in judged if r["判定结果"] == "部分通过")
    failed = sum(1 for r in judged if r["判定结果"] == "不通过")
    halluc = flags["halluc"]
    unlabeled = flags["label"]

    print(f"\n模式：{'真实接口 ' + args.endpoint if mode == 'http' else 'dry-run 桩（不代表真实质量）'}")
    print(f"执行 {len(rows)} 次（{len(recs)} 题 / 角色 {roles}）→ {out}")
    print(f"  通过 {passed}  部分通过 {partial}  不通过 {failed}  待人工 {stat.get('待人工', 0)}")
    if judged:
        print(f"  可判题通过率 {passed / len(judged):.1%}（分母排除待人工与无效题）")
    print(f"  幻觉 {halluc} 条 ｜ 标注缺失 {unlabeled} 条  ← 这两项准出要求为 0")
    if flags["hard"]:
        print(f"  技术失败（5xx / 超时 / 空响应）{flags['hard']} 次 —— 先修可用性，再看效果指标")
    if stat.get("待人工"):
        print(f"  待人工 {stat['待人工']} 条：没有机器判定规则，别当成通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
