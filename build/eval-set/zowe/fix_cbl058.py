# -*- coding: utf-8 -*-
"""就地修正 v1.0 100 题里的 EV-CBL-058（记录类型来源判定有误）。"""
import os
import math
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# 默认改仓库 eval-set/ 下那一份；换路径时用环境变量 EVAL_OUT 覆盖
P = os.environ.get("EVAL_OUT") or os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "..", "..", "eval-set", "银行知识库评测集_ZOpenEditor实测100题_v1.0.xlsx"))
FN = "微软雅黑"

NEW = {
    "问题原文": "SAM1 新增一条客户记录时，记录类型这一位是从哪里来的？如果交易键最后一位不是 A 会怎样？",
    "期望答案要点": (
        "①210-PROCESS-ADD-TRAN 先 MOVE SPACES 清空整条记录，再只填 KEY、余额、订单、12 个月（SAM1.cbl 第 323-330 行）\n"
        "②第 324 行 MOVE TRAN-KEY TO WS-CUST-KEY 是 6 字节整组赋值，5 位客户号与 1 位记录类型一起带过来\n"
        "③所以记录类型取自交易键第 6 字节；样例新增键为 '00004A'，结尾是 A，新增后就是客户记录\n"
        "④若键尾不是 A（例如只给 5 位客户号、第 6 位补空格），740 段按 IF WS-CUST-REC-TYPE = 'A' 判断会走 ELSE，"
        "整条记录被写成 69 字节联系人格式（第 441-446 行）\n"
        "⑤代码对这一位没有任何校验，正确性完全依赖调用方把 A 放在键尾"
    ),
    "期望出处": "COBOL/SAM1.cbl:323-330; COBOL/SAM1.cbl:440-446; COPYBOOK/CUSTCOPY.cpy:25-27",
    "重点校验项": "结论要落在「记录类型取自交易键第 6 字节、无校验」，不得写成「未设 REC-TYPE 导致必错」",
    "机器判定规则": "第 6 字节|324|441|!未设置",
    "备注": "实测·仓库主干·缺校验（原判定「未赋值必走联系人分支」已更正：样例键尾为 A，能正确写成客户记录）",
}


def est_height(texts_widths, line_h=12.5, pad=6, minh=34):
    lines = 1
    for t, w in texts_widths:
        if not t:
            continue
        cap = max(4, (w - 1.2) * 1.05)
        n = 0
        for seg in str(t).split("\n"):
            u = sum(2 if ord(ch) > 127 else 1 for ch in seg)
            n += max(1, math.ceil(u / cap))
        lines = max(lines, n)
    return max(minh, lines * line_h + pad)


wb = load_workbook(P)
ws = wb["评测集"]
hdr = {ws.cell(row=2, column=c).value: c for c in range(1, ws.max_column + 1)}
need = ["问题原文", "期望答案要点", "期望出处", "重点校验项", "机器判定规则", "备注"]
missing = [n for n in need if n not in hdr]
assert not missing, missing

target = None
for r in range(3, ws.max_row + 1):
    if ws.cell(row=r, column=hdr["编号"]).value == "EV-CBL-058":
        target = r
        break
assert target, "找不到 EV-CBL-058"

for k, v in NEW.items():
    c = ws.cell(row=target, column=hdr[k])
    c.value = v
    c.font = Font(name=FN, size=9, bold=False, color="262626")
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

ws.row_dimensions[target].height = est_height(
    [(NEW["问题原文"], 38), (NEW["期望答案要点"], 40), (NEW["重点校验项"], 26),
     (NEW["机器判定规则"], 20), (NEW["备注"], 16)])
wb.save(P)
print(f"[OK] 已修正第 {target} 行 EV-CBL-058，行高 {ws.row_dimensions[target].height:.0f}")
