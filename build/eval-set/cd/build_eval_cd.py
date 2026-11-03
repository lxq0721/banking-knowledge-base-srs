# -*- coding: utf-8 -*-
"""
生成《银行知识库评测集 · CardDemo 实测 100 题 v2.0》
单 sheet，五个分类用「分组标题行 + 一个空行」分隔。
零公式：统计值全部静态 + DataValidation 下拉 + 条件格式。
"""
import os
import math
import collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

from q_cd_1 import G1, G2
from q_cd_2 import G3, G4, G5
from q_cd_base import (CAT_DEV, CAT_BIZ, CAT_TEL, CAT_CFT, CAT_QUAL,
                       T1, T2, T3, T4, T5, T6, T8, T9)

# 默认写到仓库 eval-set/ 下；换路径时用环境变量 EVAL_OUT 覆盖
OUT = os.environ.get("EVAL_OUT") or os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "..", "..", "eval-set", "银行知识库评测集_CardDemo实测100题_v2.0.xlsx"))

FN = "微软雅黑"
C_DARK = "1F3864"
C_GRP = "2E5C8A"
C_HDR = "4472C4"
C_ROW = "FFFFFF"
C_LINE = "BFBFBF"
C_NOTE = "595959"

GROUPS = [
    ("01 开发与迁移", "迁移实施方 · 开发 · 代码走查　｜　保留程序名、字段名、行号，判分看出处是否精确到行", G1),
    ("02 业务分析", "需求分析 · 系统分析 · 数据治理　｜　不点程序名只点业务对象，追问口径边界与不确定处", G2),
    ("03 行员", "柜员 · 网点业务人员　｜　只用报错原文和「我该怎么办」，判分看是否给出可执行动作", G3),
    ("04 冲突与不一致", "知识库实施方 · 质量　｜　同一事实两侧说法不同，判分看是否按「可否同时成立」区分真冲突与详略差异", G4),
    ("05 质量与边界", "测试人员　｜　空输入、诱导、注入、越权、幻觉，判分看越界与编造是否为 0", G5),
]

COLS_RAW = [
    ("编号", 13, "题面"),
    ("分类", 15, "题面"),
    ("所属范围", 14, "题面"),
    ("问题原文", 32, "题面"),
    ("问题类型", 15, "题面"),
    ("难度", 7, "题面"),
    ("需分角色执行", 8, "题面"),
    ("库内是否存在该信息", 9, "判分依据"),
    ("事实基准", 9, "判分依据"),
    ("期望答案要点", 46, "判分依据"),
    ("期望出处", 28, "判分依据"),
    ("期望来源等级", 15, "预期属性"),
    ("期望内容状态", 9, "预期属性"),
    ("重点校验项", 24, "预期属性"),
    ("机器判定规则", 18, "预期属性"),
    ("业务确认人", 8, "管理"),
    ("确认日期", 10, "管理"),
    ("备注", 24, "管理"),
]
# 实测 100% 缩放下表格只占页宽 65%，整体放大到接近满宽
WSCALE = 1.0
COLS = [(n, round(w * WSCALE, 1), g) for n, w, g in COLS_RAW]
NC = len(COLS)
KEYS = ["id", "cat", "scope", "q", "t", "diff", "role", "inkb", "basis",
        "pts", "src", "lvl", "state", "chk", "rule", "", "", "note"]

thin = Side(style="thin", color=C_LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def cell(ws, row, col, text, size=9, bold=False, fill=None, wrap=True,
         color="262626", halign="left", valign="top", span=1, border=True):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=FN, size=size, bold=bold, color=color)
    c.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=wrap)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    for i in range(col, col + span):
        cc = ws.cell(row=row, column=i)
        if border:
            cc.border = BORDER
        if fill:
            cc.fill = PatternFill("solid", fgColor=fill)
    return c


def est_height(pairs, line_h=13.6, pad=11, minh=34):
    lines = 1
    for t, w in pairs:
        if not t:
            continue
        cap = max(4, (w - 1.6) * 0.94)
        n = 0
        for seg in str(t).split("\n"):
            u = sum(2 if ord(ch) > 127 else 1 for ch in seg)
            n += max(1, math.ceil(u / cap))
        lines = max(lines, n)
    return max(minh, lines * line_h + pad)


wb = Workbook()
ws = wb.active
ws.title = "评测集"
ws.sheet_view.showGridLines = False
for i, (_, w, _) in enumerate(COLS, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---------------------------------------------------------------- 行 1 标题
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NC)
c = ws.cell(row=1, column=1,
            value="银行知识库评测集 · CardDemo 实测 100 题")
c.font = Font(name=FN, size=18, bold=True, color=C_DARK)
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 32

# ---------------------------------------------------------------- 行 2 说明
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NC)
c = ws.cell(row=2, column=1,
            value=("样本库：aws-samples/aws-mainframe-modernization-carddemo @59cc6c2（信用卡核心业务，"
                   "46 程序 / 62 copybook / 55 JCL / 21 BMS 屏）+ IBM/zopeneditor-sample @8f98353（多语言补充）"
                   "　｜　v2.0　｜　100 题，五个分类在同一页，分类之间空一行　｜　"
                   "出处全部落在真实代码行，行号经逐条核实"))
c.font = Font(name=FN, size=9, color=C_NOTE)
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws.row_dimensions[2].height = 26

# ---------------------------------------------------------------- 行 3 分组表头
groups, start_col = [], 1
for i, (_, _, g) in enumerate(COLS, 1):
    if not groups or groups[-1][0] != g:
        groups.append([g, i, i])
    else:
        groups[-1][2] = i
for g, c1, c2 in groups:
    nc = cell(ws, 3, c1, g, size=10, bold=True, color="FFFFFF", fill=C_GRP,
              halign="center", valign="center", span=c2 - c1 + 1)
ws.row_dimensions[3].height = 22

# ---------------------------------------------------------------- 行 4 字段名
for i, (name, _, _) in enumerate(COLS, 1):
    cell(ws, 4, i, name, size=9, bold=True, color="FFFFFF", fill=C_HDR,
         halign="center", valign="center")
ws.row_dimensions[4].height = 36

# ---------------------------------------------------------------- 数据行
HDR_ROWS = 4
r = HDR_ROWS + 1
blocks = []          # [(group_name, first_row, last_row)]
for gi, (gname, gdesc, rows) in enumerate(GROUPS):
    if gi > 0:
        # 组间空一行：不设边框、不填底色
        ws.row_dimensions[r].height = 7
        r += 1

    # 分组标题行
    cell(ws, r, 1, f"{gname}　｜　{len(rows)} 题　｜　{gdesc}",
         size=10, bold=True, color="FFFFFF", fill=C_GRP,
         valign="center", span=NC)
    ws.row_dimensions[r].height = 26
    r += 1

    first = r
    for row in rows:
        for i, key in enumerate(KEYS, 1):
            val = row[i - 1] if key else ""
            center = i in (1, 6, 7, 8, 9, 13, 16)
            cell(ws, r, i, val, size=9, fill=C_ROW,
                 halign="center" if center else "left", valign="top")
        ws.row_dimensions[r].height = est_height(
            [(row[9], COLS[9][1]), (row[3], COLS[3][1]), (row[10], COLS[10][1]),
             (row[13], COLS[13][1]), (row[14], COLS[14][1]), (row[17], COLS[17][1])],
            line_h=13.6, pad=11, minh=34)
        r += 1
    blocks.append((gname, first, r - 1))

last_data = r - 1

# ---------------------------------------------------------------- 冻结与筛选
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A{HDR_ROWS}:{get_column_letter(NC)}{last_data}"
ws.sheet_view.zoomScale = 82

# ---------------------------------------------------------------- 下拉与条件格式
def add_dv(formula, rng):
    dv = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True,
                        showDropDown=False)
    dv.error = "请从下拉列表中选择"
    dv.errorTitle = "取值不在范围内"
    ws.add_data_validation(dv)
    dv.add(rng)


red = PatternFill("solid", fgColor="FFC7CE")
yellow = PatternFill("solid", fgColor="FFEB9C")
red_font = Font(name=FN, size=9, color="9C0006", bold=True)

for gname, a, b in blocks:
    add_dv(f"{CAT_DEV},{CAT_BIZ},{CAT_TEL},{CAT_CFT},{CAT_QUAL}", f"B{a}:B{b}")
    add_dv(f"{T1},{T2},{T3},{T4},{T5},{T6},7 跨模态,{T8},{T9}", f"E{a}:E{b}")
    add_dv("简单,中等,困难", f"F{a}:F{b}")
    add_dv("否,是", f"G{a}:G{b}")
    add_dv("是,否,—", f"H{a}:H{b}")
    add_dv("原始资料,开发补全物,推断,—", f"I{a}:I{b}")
    add_dv("代码直出,引自文档,人工修正,模型生成·核验通过,模型生成·核验未通过,—",
           f"L{a}:L{b}")
    add_dv("可取用,已修正,有争议,已废止,—", f"M{a}:M{b}")

    ws.conditional_formatting.add(f"H{a}:H{b}", CellIsRule(
        operator="equal", formula=['"否"'], fill=red, font=red_font))
    ws.conditional_formatting.add(f"L{a}:L{b}", CellIsRule(
        operator="equal", formula=['"模型生成·核验通过"'], fill=yellow))
    ws.conditional_formatting.add(f"M{a}:M{b}", CellIsRule(
        operator="equal", formula=['"有争议"'], fill=yellow))
    ws.conditional_formatting.add(f"M{a}:M{b}", CellIsRule(
        operator="equal", formula=['"已废止"'], fill=yellow))

# ---------------------------------------------------------------- 打印
ws.oddHeader.center.text = "CardDemo 实测 100 题 v2.0　｜　单页五分类"
ws.oddHeader.center.size = 8
ws.oddHeader.center.font = "微软雅黑,Italic"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = 8
# 18 列宽表：缩到 1 页宽便于整体翻阅；需要大字打印时在 Excel 里自行调缩放
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"1:{HDR_ROWS}"
ws.print_options.horizontalCentered = True

# ---------------------------------------------------------------- 末尾统计块
ALL = G1 + G2 + G3 + G4 + G5
r = last_data + 2

cell(ws, r, 1, "配比自检（静态值，非公式）", size=11, bold=True,
     color="FFFFFF", fill=C_DARK, valign="center", span=NC)
ws.row_dimensions[r].height = 24
r += 1

by_cat = collections.Counter(x[1] for x in ALL)
by_type = collections.Counter(x[4] for x in ALL)
by_diff = collections.Counter(x[5] for x in ALL)
no_kb = sum(1 for x in ALL if x[7] == "否")
manual = sum(1 for x in ALL if not str(x[14]).strip())
with_line = sum(1 for x in ALL
                if x[10] and x[10] != "—" and any(ch.isdigit() for ch in str(x[10])))

rows_stat = [
    ("按分类", "　".join(f"{k.split(' ')[0]} {v} 题" for k, v in sorted(by_cat.items())
                       ), f"合计 {len(ALL)} 题"),
    ("按问题类型", "　".join(f"{k.split(' ')[0]}类 {v}" for k, v in sorted(by_type.items())),
     "第 7 类跨模态为 0：两套素材均无图形/界面图像资料，整类裁剪"),
    ("按难度", "　".join(f"{k} {v}" for k, v in
                       sorted(by_diff.items(), key=lambda p: ["简单", "中等", "困难"].index(p[0]))),
     "困难题集中在多跳与冲突，需要跨文件核对"),
    ("库内不存在（幻觉专项）", f"{no_kb} 条",
     f"占 {no_kb / len(ALL) * 100:.1f}%，硬下限 10%"),
    ("可自动判定", f"{len(ALL) - manual} 条",
     f"自动率 {(len(ALL) - manual) / len(ALL) * 100:.1f}%；{manual} 条必须人工（空输入/越界/注入）"),
    ("出处带行号或定位", f"{with_line} 条",
     "无答案题的出处写「—（应无出处）」，不计入此项"),
]
for label, val, note in rows_stat:
    cell(ws, r, 1, label, size=9, bold=True, fill="DEEAF6", span=2, valign="center")
    cell(ws, r, 3, val, size=9, span=9, valign="center")
    cell(ws, r, 12, note, size=9, span=NC - 11, color=C_NOTE, valign="center")
    ws.row_dimensions[r].height = est_height([(val, 220), (note, 165)], minh=24)
    r += 1

r += 1
tip = ("判分口径：机器规则用 A|B|!C|/regex/ 表达，| 分隔，! 为禁词，/…/ 为正则，留空代表只能人工判。"
       "判定顺序不可颠倒——先硬失败（5xx/超时/空响应），再幻觉专项，再级别标注，最后规则匹配；"
       "无规则者落「待人工」，不得默认通过。must_not 是否决项，不进命中率分母。")
cell(ws, r, 1, tip, size=9, color=C_NOTE, span=NC)
ws.row_dimensions[r].height = est_height([(tip, 430)], minh=30)

wb.save(OUT)
print("saved:", OUT)
print("数据行", HDR_ROWS + 1, "→", last_data, "｜分组", len(blocks), "｜末尾行", r)
for g, a, b in blocks:
    print(f"  {g}  {a}-{b}  ({b - a + 1} 行)")
print("分类:", dict(by_cat))
print("类型:", dict(sorted(by_type.items())))
print("难度:", dict(by_diff))
print(f"无答案 {no_kb} 条 ({no_kb/len(ALL)*100:.1f}%)  人工 {manual} 条  自动率 {(len(ALL)-manual)/len(ALL)*100:.1f}%")
