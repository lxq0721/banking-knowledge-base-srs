# -*- coding: utf-8 -*-
"""
生成《银行知识库评测集 · 分视角题型示例 v1.0》
零公式：统计值全部静态 + DataValidation 下拉 + 条件格式。
"""
import os
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

from q_views import VIEWS
from q_views_meta import GAPS, MATRIX, RUBRIC, KEY_FINDINGS

# 默认写到仓库 eval-set/ 下；换路径时用环境变量 EVAL_OUT 覆盖
OUT = os.environ.get("EVAL_OUT") or os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "..", "..", "eval-set", "银行知识库评测集_分视角题型示例_v1.0.xlsx"))

FN = "微软雅黑"
C_DARK = "1F3864"
C_GRP = "2E5C8A"
C_HDR = "4472C4"
C_EG = "FFF9E6"
C_LINE = "BFBFBF"
C_NOTE = "595959"
C_VIEW = ["DEEAF6", "E2EFDA", "FFF2CC", "FCE4D6", "EDE7F6"]

thin = Side(style="thin", color=C_LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def title_cell(ws, row, col, text, span=1, size=14, color=C_DARK, fill=None,
               bold=True, height=None, halign="left"):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=FN, size=size, bold=bold, color=color)
    c.alignment = Alignment(horizontal=halign, vertical="center", wrap_text=True)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    for i in range(col, col + span):
        cc = ws.cell(row=row, column=i)
        cc.border = BORDER
        if fill:
            cc.fill = PatternFill("solid", fgColor=fill)
    if height:
        ws.row_dimensions[row].height = height
    return c


def body_cell(ws, row, col, text, size=9, bold=False, fill=None, wrap=True,
              color="262626", halign="left", valign="top", span=1):
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
        cc.border = BORDER
        if fill:
            cc.fill = PatternFill("solid", fgColor=fill)
    return c


def est_height(texts_widths, line_h=13.0, pad=6, minh=20):
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


wb = Workbook()
wb.remove(wb.active)

# =====================================================================
# Sheet 0  说明
# =====================================================================
ws = wb.create_sheet("00 说明")
ws.sheet_view.showGridLines = False
for i, w in enumerate([7, 20, 92, 22], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws, r, 1, "银行知识库评测集 · 分视角题型示例", span=4, size=18, height=34)
r += 1
title_cell(ws, r, 1,
           "样本库：IBM/zopeneditor-sample 主干 @8f98353 ｜ 版本 v1.0 ｜ 共 68 题 + 22 条缺口方向 ｜ "
           "五个视角分页存放，出处全部落在真实代码行",
           span=4, size=9, color=C_NOTE, bold=False, height=22)

r = 3
title_cell(ws, r, 1, "一、为什么按视角分页", span=4, size=12, fill="DEEAF6", height=24)
r += 1
txt = ("上一版 100 题是按「知识库结构覆盖」铺开的：先按文件分块，再按题型配额，最后填内容。这样出来的题目结构很匀，"
       "但它回答不了「行员会不会这么问」。按提问角色重判后：开发与迁移 78 条、库管理 10 条、质量边界 8 条、业务分析 4 条、行员 0 条。"
       "本文件把同一批事实按五种提问角色重写一遍，放在不同页，便于逐页对照——同一个知识点换个人问，难度和正确答案都不一样。")
body_cell(ws, r, 1, txt, span=4, size=9)
ws.row_dimensions[r].height = est_height([(txt, 116)], minh=30)

r += 2
title_cell(ws, r, 1, "二、五个视角怎么分工", span=4, size=12, fill="DEEAF6", height=24)
r += 1
for i, h in enumerate(["", "视角", "谁来问 / 怎么问 / 判分时看什么", "条数"], 1):
    title_cell(ws, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=22, halign="center")
r += 1
ROLE_TABLE = [
    ("01 开发与迁移", "迁移实施方、开发、代码走查。保留程序名、字段名、行号——他们本来就懂，藏起来反而降低定位精度。判分看出处是否精确到行。", "14"),
    ("02 业务分析", "需求分析、系统分析、数据治理。不点程序名，只点业务对象；追问口径边界（能不能为负、含不含失败、保留多久）。判分看是否点出边界、不确定处是否标注。", "14"),
    ("03 行员", "柜员、网点业务人员。只用报错原文和「我该怎么办」，不出任何技术标识符。答案经常只能答一半，另一半在业务文档里。判分看是否给出可执行动作、查不到是否承认。", "16"),
    ("04 质量与边界", "测试人员。不追求业务价值，专探会不会胡编、越界、被施压改口。判分看幻觉与越界是否为 0。", "12"),
    ("05 库管理", "知识库实施方、运维。问「这条内容以什么身份进库、出处给到什么颗粒度、两边打架怎么办」。判分看元数据与流程有没有被执行。", "12"),
]
for k, (name, desc, n) in enumerate(ROLE_TABLE):
    fill = C_VIEW[k]
    body_cell(ws, r, 1, str(k + 1), size=9, fill=fill, halign="center", valign="center")
    body_cell(ws, r, 2, name, size=9, bold=True, fill=fill, color=C_DARK, valign="center")
    body_cell(ws, r, 3, desc, size=9, fill=fill)
    body_cell(ws, r, 4, n, size=9, fill=fill, halign="center", valign="center")
    ws.row_dimensions[r].height = est_height([(desc, 92)], minh=30)
    r += 1

r += 1
title_cell(ws, r, 1, "三、五条判断（这一版的主要结论）", span=4, size=12, fill="DEEAF6", height=24)
r += 1
for k, (head, desc) in enumerate(KEY_FINDINGS, 1):
    body_cell(ws, r, 1, str(k), size=9, halign="center", valign="center")
    body_cell(ws, r, 2, head, size=9, bold=True, color=C_DARK)
    body_cell(ws, r, 3, desc, size=9)
    body_cell(ws, r, 4, "", size=9)
    ws.row_dimensions[r].height = est_height([(desc, 92), (head, 22)], minh=30)
    r += 1

r += 1
title_cell(ws, r, 1, "四、怎么用这份表", span=4, size=12, fill="DEEAF6", height=24)
r += 1
STEPS = [
    ("第 1 步", "先看 07 页", "同一事实的三种问法摆在一起。看完这一页就明白：问法一换，正确答案的边界就变了，用一套标准判五类题必然失真。"),
    ("第 2 步", "按视角挑题", "要验检索精度就挑 01 页；要验业务口径能力挑 02 页；要给业务方演示「系统能不能回答柜员的问题」挑 03 页；要验幻觉挑 04 页。"),
    ("第 3 步", "查 06 页确认缺口", "凡题目落到 06 页列出的 22 个方向，一律标「库内不存在」进幻觉专项，不要标成「是」再逼模型编。"),
    ("第 4 步", "按 08 页分别定准出", "五个视角的通过率不能合成一个数。开发可卡 90%，业务分析要卡「不确定项标注 100%」，行员不许出现编造的操作步骤。"),
]
for tag, head, desc in STEPS:
    body_cell(ws, r, 1, tag, size=9, bold=True, color=C_DARK, halign="center", valign="center")
    body_cell(ws, r, 2, head, size=9, bold=True, color=C_DARK)
    body_cell(ws, r, 3, desc, size=9, span=2)
    ws.row_dimensions[r].height = est_height([(desc, 116)], minh=28)
    r += 1

r += 1
title_cell(ws, r, 1, "五、页面结构", span=4, size=12, fill="DEEAF6", height=24)
r += 1
for i, h in enumerate(["", "页面", "放什么", "什么时候看"], 1):
    title_cell(ws, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=22, halign="center")
r += 1
PAGES = [
    ("01~05 视角页", "五组题，每组列结构一致（19 列）。比上一版多两列：视角内题型、这个视角特有的缺口。", "按需挑页"),
    ("06 代码里没有的方向", "22 个方向。每行列「会怎么问 / 代码里有什么替代物 / 缺了会怎样 / 要补什么 / 补上后能覆盖什么题型」。", "出题前先过一遍"),
    ("07 同一事实三种问法", "9 个知识点，各自摆出开发问法、业务分析问法、行员问法，并标难度走向与能否自动判。", "解释「为什么要分视角」时"),
    ("08 视角与判分口径", "五个视角的期望等级、机判可行性、判分侧重、准出建议，以及与上一版 100 题的对应关系。", "定准出标准时"),
]
for k, (name, desc, when) in enumerate(PAGES, 1):
    body_cell(ws, r, 1, str(k), size=9, halign="center", valign="center")
    body_cell(ws, r, 2, name, size=9, bold=True, color=C_DARK)
    body_cell(ws, r, 3, desc, size=9)
    body_cell(ws, r, 4, when, size=9, halign="center", valign="center")
    ws.row_dimensions[r].height = est_height([(desc, 92)], minh=28)
    r += 1

r += 2
body_cell(ws, r, 1,
          "说明：本文件用于「看题型与确认口径」，不是正式跑批集。"
          "正式跑批请把这些题并入主评测表（同 17 列结构 + 2 列扩展），或直接用 export_jsonl.py 指定本文件的视角页导出。",
          span=4, size=8, color=C_NOTE)
ws.row_dimensions[r].height = 30
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = 8
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.horizontalCentered = True

# =====================================================================
# Sheet 1~5  五个视角
# =====================================================================
COLS = [
    ("编号", 13, "题面"),
    ("所属范围", 13, "题面"),
    ("问题原文", 30, "题面"),
    ("问题类型", 13, "题面"),
    ("难度", 7, "题面"),
    ("需分角色执行", 8, "题面"),
    ("库内是否存在该信息", 9, "判分依据"),
    ("事实基准", 9, "判分依据"),
    ("期望答案要点", 42, "判分依据"),
    ("期望出处", 23, "判分依据"),
    ("期望来源等级", 16, "预期属性"),
    ("期望内容状态", 9, "预期属性"),
    ("重点校验项", 20, "预期属性"),
    ("机器判定规则", 17, "预期属性"),
    ("业务确认人", 8, "管理"),
    ("确认日期", 9, "管理"),
    ("备注", 18, "管理"),
    ("视角内题型", 12, "视角专用"),
    ("这个视角特有的缺口", 16, "视角专用"),
]

KEYS = ["id", "scope", "q", "t", "diff", "role", "inkb", "basis", "pts", "src",
        "lvl", "state", "chk", "rule", "", "", "note", "sub", "gap"]

HDR_ROW = 2
red_fill = PatternFill("solid", fgColor="FFC7CE")
yellow_fill = PatternFill("solid", fgColor="FFEB9C")
red_font = Font(name=FN, size=9, color="9C0006", bold=True)


def add_dv(ws, formula, rng):
    dv = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True,
                        showDropDown=False)
    dv.error = "请从下拉列表中选择"
    dv.errorTitle = "取值不在范围内"
    ws.add_data_validation(dv)
    dv.add(rng)


all_ids = []
for vi, view in enumerate(VIEWS):
    ws = wb.create_sheet(view["sheet"])
    ws.sheet_view.showGridLines = False
    for i, (name, w, grp) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    groups = {}
    for i, (_, _, g) in enumerate(COLS, 1):
        groups.setdefault(g, []).append(i)
    for g, idxs in groups.items():
        title_cell(ws, 1, idxs[0], g, span=len(idxs), size=10, color="FFFFFF",
                   fill=C_GRP, height=22, halign="center")
    for i, (name, w, _) in enumerate(COLS, 1):
        title_cell(ws, HDR_ROW, i, name, size=9, color="FFFFFF", fill=C_HDR,
                   height=34, halign="center")

    start = HDR_ROW + 1
    for k, row in enumerate(view["rows"]):
        r = start + k
        row = dict(row)
        row.setdefault("scope", view["scope"])
        row.setdefault("role", "否")
        row.setdefault("state", "可取用")
        row.setdefault("basis", "原始资料")
        row.setdefault("gap", "")
        all_ids.append(row["id"])
        for i, key in enumerate(KEYS, 1):
            val = row.get(key, "") if key else ""
            if key == "gap" and not val:
                val = "—"
            center = i in (1, 4, 5, 6, 7, 8, 12, 16, 18)
            body_cell(ws, r, i, val, size=9, fill=C_EG,
                      halign="center" if center else "left",
                      valign="top", wrap=True)
        ws.row_dimensions[r].height = est_height(
            [(row.get("q"), 30), (row.get("pts"), 42),
             (row.get("chk"), 20), (row.get("rule"), 17), (row.get("note"), 18)],
            line_h=13.0, pad=9, minh=34)

    last = start + len(view["rows"]) - 1
    ws.freeze_panes = "D3"
    ws.auto_filter.ref = f"A{HDR_ROW}:S{last}"
    ws.sheet_view.zoomScale = 85

    add_dv(ws, "1 精准命中,2 同义改写+跨语言,3 信息分散（多跳）,4 知识库不存在（幻觉）,5 冲突信息,6 异常与边界,7 跨模态,8 多轮对话,9 元数据过滤", f"D{start}:D{last}")
    add_dv(ws, "简单,中等,困难", f"E{start}:E{last}")
    add_dv(ws, "否,是", f"F{start}:F{last}")
    add_dv(ws, "是,否,—", f"G{start}:G{last}")
    add_dv(ws, "原始资料,开发补全物,推断,—", f"H{start}:H{last}")
    add_dv(ws, "代码直出,引自文档,人工修正,模型生成·核验通过,模型生成·核验未通过,—", f"K{start}:K{last}")
    add_dv(ws, "可取用,已修正,有争议,已废止,—", f"L{start}:L{last}")

    ws.conditional_formatting.add(
        f"G{start}:G{last}",
        CellIsRule(operator="equal", formula=['"否"'], fill=red_fill, font=red_font))
    ws.conditional_formatting.add(
        f"K{start}:K{last}",
        CellIsRule(operator="equal", formula=['"模型生成·核验通过"'], fill=yellow_fill))
    ws.conditional_formatting.add(
        f"L{start}:L{last}",
        CellIsRule(operator="equal", formula=['"有争议"'], fill=yellow_fill))
    ws.conditional_formatting.add(
        f"L{start}:L{last}",
        CellIsRule(operator="equal", formula=['"已废止"'], fill=yellow_fill))

    ws.oddHeader.center.text = f"{view['sheet']} ｜ 提问人：{view['who']}"
    ws.oddHeader.center.size = 8
    ws.oddHeader.center.font = "微软雅黑,Italic"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 8
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    ws.print_options.horizontalCentered = True

# =====================================================================
# Sheet 06  代码里没有的方向
# =====================================================================
ws6 = wb.create_sheet("06 代码里没有的方向")
ws6.sheet_view.showGridLines = False
W6 = [4, 22, 30, 40, 40, 24, 16, 7]
for i, w in enumerate(W6, 1):
    ws6.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws6, r, 1, "代码里没有、但行员和业务迟早会问的方向", span=8, size=15, height=32)
r += 1
title_cell(ws6, r, 1,
           "这份样本是 COBOL 示例集，不是银行核心系统。下面 22 个方向都是能从业务经验想到的，但这份代码里查不到——"
           "不是检索不到，是资料本来就没有。凡题目落到这些方向，一律标「库内不存在」进幻觉专项，不要标成「是」再逼模型编。",
           span=8, size=9, color=C_NOTE, bold=False)
ws6.row_dimensions[r].height = est_height([("x" * 160, 176)], minh=30)
r = 3
for i, h in enumerate(["№", "缺的方向", "会怎么问（一句真实问法）", "代码里有什么替代物",
                       "缺了会怎样（知识库会发生什么）", "要补什么资料", "补上后覆盖的题型", "优先级"], 1):
    title_cell(ws6, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=30, halign="center")
r += 1
for k, g in enumerate(GAPS, 1):
    fill = C_VIEW[1] if g[6] == "高" else (C_VIEW[2] if g[6] == "中" else None)
    body_cell(ws6, r, 1, str(k), size=9, halign="center", valign="center")
    body_cell(ws6, r, 2, g[0], size=9, bold=True, color=C_DARK)
    body_cell(ws6, r, 3, g[1], size=9)
    body_cell(ws6, r, 4, g[2], size=9)
    body_cell(ws6, r, 5, g[3], size=9)
    body_cell(ws6, r, 6, g[4], size=9)
    body_cell(ws6, r, 7, g[5], size=9, halign="center", valign="center")
    body_cell(ws6, r, 8, g[6], size=9, bold=(g[6] == "高"), halign="center", valign="center",
              color="C00000" if g[6] == "高" else "262626")
    ws6.row_dimensions[r].height = est_height(
        [(g[2], 40), (g[3], 40), (g[4], 24), (g[1], 30)], line_h=12.2, pad=8, minh=32)
    r += 1

r += 1
body_cell(ws6, r, 1, "优先级判法：高 = 行员张口就问、且答错会直接误导操作；中 = 业务分析会问、答错影响判断；低 = 主要影响开发与运维，业务侧很少问。",
          span=8, size=8, color=C_NOTE)

ws6.freeze_panes = "C4"
ws6.auto_filter.ref = f"A3:H{r - 2}"
ws6.page_setup.orientation = "landscape"
ws6.page_setup.paperSize = 8
ws6.page_setup.fitToWidth = 1
ws6.page_setup.fitToHeight = 0
ws6.sheet_properties.pageSetUpPr.fitToPage = True
ws6.print_title_rows = "3:3"

# =====================================================================
# Sheet 07  同一事实三种问法
# =====================================================================
ws7 = wb.create_sheet("07 同一事实三种问法")
ws7.sheet_view.showGridLines = False
W7 = [4, 18, 28, 26, 30, 28, 14, 24]
for i, w in enumerate(W7, 1):
    ws7.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws7, r, 1, "同一个知识点，换个人问，答案边界就变了", span=8, size=15, height=32)
r += 1
title_cell(ws7, r, 1,
           "这张表是出题模板：以后每加一个知识点，先按三列各写一遍，再决定收哪几条进评测集。"
           "开发问法答得出就算过；行员问法答得出不算完，还要看有没有接触到业务后果那一层。",
           span=8, size=9, color=C_NOTE, bold=False)
ws7.row_dimensions[r].height = est_height([("x" * 140, 172)], minh=30)
r = 3
for i, h in enumerate(["№", "知识点", "开发问法", "业务分析问法", "行员问法",
                       "答案落在哪", "难度走向", "能否自动判"], 1):
    title_cell(ws7, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=28, halign="center")
r += 1
for k, m in enumerate(MATRIX, 1):
    body_cell(ws7, r, 1, str(k), size=9, halign="center", valign="center")
    body_cell(ws7, r, 2, m[0], size=9, bold=True, color=C_DARK)
    body_cell(ws7, r, 3, m[1], size=9, fill=C_VIEW[0])
    body_cell(ws7, r, 4, m[2], size=9, fill=C_VIEW[1])
    body_cell(ws7, r, 5, m[3], size=9, fill=C_VIEW[2])
    body_cell(ws7, r, 6, m[4], size=9)
    body_cell(ws7, r, 7, m[5], size=9, halign="center")
    body_cell(ws7, r, 8, m[6], size=9)
    ws7.row_dimensions[r].height = est_height(
        [(m[1], 28), (m[2], 26), (m[3], 30), (m[6], 24)], line_h=12.2, pad=8, minh=32)
    r += 1

r += 1
body_cell(ws7, r, 1,
          "读法：同一行里，越往右提问的人越不懂系统，但要求的答案越贴近业务后果。第 8 行（新增后查不到）是典型——"
          "行员问法最短、最容易问出口，答案却最难：要说清记录类型取自交易键最后一位、系统没有校验。",
          span=8, size=8, color=C_NOTE)
ws7.row_dimensions[r].height = 26

ws7.freeze_panes = "C4"
ws7.page_setup.orientation = "landscape"
ws7.page_setup.paperSize = 8
ws7.page_setup.fitToWidth = 1
ws7.page_setup.fitToHeight = 0
ws7.sheet_properties.pageSetUpPr.fitToPage = True
ws7.print_title_rows = "3:3"

# =====================================================================
# Sheet 08  视角与判分口径
# =====================================================================
ws8 = wb.create_sheet("08 视角与判分口径")
ws8.sheet_view.showGridLines = False
W8 = [4, 15, 24, 26, 16, 30, 26, 8, 34]
for i, w in enumerate(W8, 1):
    ws8.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws8, r, 1, "五个视角不能合成一个通过率", span=9, size=15, height=32)
r += 1
title_cell(ws8, r, 1,
           "期望来源等级不同、机判可行性不同、答错的后果也不同。混在一起算总分，等于把「答错一个字段长度」和「编一个柜面操作步骤」当同一件事。",
           span=9, size=9, color=C_NOTE, bold=False, height=22)
r = 3
for i, h in enumerate(["№", "视角", "谁来问", "期望来源等级主要在", "机判可行性",
                       "判分侧重", "准出建议", "条数", "与上一版 100 题的关系"], 1):
    title_cell(ws8, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=30, halign="center")
r += 1
for k, rb in enumerate(RUBRIC, 1):
    fill = C_VIEW[k - 1]
    body_cell(ws8, r, 1, str(k), size=9, fill=fill, halign="center", valign="center")
    body_cell(ws8, r, 2, rb[0], size=9, bold=True, fill=fill, color=C_DARK)
    body_cell(ws8, r, 3, rb[1], size=9, fill=fill)
    body_cell(ws8, r, 4, rb[2], size=9, fill=fill)
    body_cell(ws8, r, 5, rb[3], size=9, fill=fill)
    body_cell(ws8, r, 6, rb[4], size=9, fill=fill)
    body_cell(ws8, r, 7, rb[5], size=9, fill=fill)
    body_cell(ws8, r, 8, rb[6], size=9, fill=fill, halign="center", valign="center")
    body_cell(ws8, r, 9, rb[7], size=9, fill=fill)
    ws8.row_dimensions[r].height = est_height(
        [(rb[1], 24), (rb[3], 30), (rb[4], 26), (rb[7], 34)], line_h=12.2, pad=8, minh=34)
    r += 1

r += 1
body_cell(ws8, r, 1, "合计", size=9, bold=True, color=C_DARK, halign="center")
body_cell(ws8, r, 2, "五类合计", size=9, bold=True, color=C_DARK)
body_cell(ws8, r, 3, "", size=9)
body_cell(ws8, r, 4, "", size=9)
body_cell(ws8, r, 5, "", size=9)
body_cell(ws8, r, 6, "", size=9)
body_cell(ws8, r, 7, "分别卡，不加总", size=9, bold=True)
body_cell(ws8, r, 8, str(sum(int(x[6]) for x in RUBRIC)), size=9, bold=True, halign="center")
body_cell(ws8, r, 9, "上一版 100 题主要覆盖第 1、4、5 类；第 2 类只有 4 条近似，第 3 类完全没有", size=9)
ws8.row_dimensions[r].height = 32

r += 2
title_cell(ws8, r, 1, "为什么行员题不能按通过率卡", span=9, size=11, fill="FFF2CC", height=24)
r += 1
why = ("行员问「客户名最多输几个字」，正确答案不是「17」这一个数，而是「17 字节、超出被截断不报错、"
       "中文大约只能放 8 个字」。后半段代码里没有，只能靠编码知识和前端规则推断。"
       "这样的题按要点命中率判，会出现「答 17 得满分、答 17 并说明截断风险反而因为多说了不确定内容被扣分」的倒挂。"
       "所以行员题的判定要做两件事：一是把「必须答对的事实」和「最好答到的后果」分开计分，"
       "二是把「明确说不知道、需要业务确认」也计为正确——这比编一个像样的操作步骤有价值得多。\n"
       "另外提醒：本文件 68 题全部填了机器判定规则，导出时显示「自动率 100%」，但那只说明每条都能跑规则，"
       "不说明规则判得准。行员页的规则只校验硬事实（17、12、9(7)V99 这类），答到这些词就算命中，"
       "至于有没有触到业务后果，机器判不出来，必须人工复核后半段。所以行员页的实际判定方式是「机判粗筛 + 人工复核」，"
       "不能拿机判通过率当准出依据。")
body_cell(ws8, r, 1, why, span=9, size=9)
ws8.row_dimensions[r].height = est_height([(why, 200)], minh=40)

r += 2
title_cell(ws8, r, 1, "与上一版的差异（要改的话改哪里）", span=9, size=11, fill="DEEAF6", height=24)
r += 1
DIFF = [
    ("列结构", "新增两列：视角内题型、这个视角特有的缺口。原 17 列不动，导出脚本不受影响。"),
    ("编号", "本文件用 EV-DEV/BIZ/TLR/TST/MGT 五个新域，与上一版 EV-CPY/CBL/... 不冲突，可同表合并。"),
    ("行员题", "上一版 0 条，本文件 16 条。其中 1 条（网点可见范围）标为库内不存在，其余多数只能答一半。"),
    ("缺口", "22 个方向单独成页。上一版把这些散落在「无答案题」里，看不出是「资料缺失」还是「检索失败」，这一版分开。"),
    ("判分", "上一版一套判定顺序通吃；这一版按视角分档，行员题允许「不确定」计对，其余视角仍按原判定顺序。"),
]
for k, (a, b) in enumerate(DIFF, 1):
    body_cell(ws8, r, 1, str(k), size=9, halign="center", valign="center")
    body_cell(ws8, r, 2, a, size=9, bold=True, color=C_DARK, span=2)
    body_cell(ws8, r, 4, b, size=9, span=6)
    ws8.row_dimensions[r].height = est_height([(b, 178)], minh=26)
    r += 1

ws8.page_setup.orientation = "landscape"
ws8.page_setup.paperSize = 8
ws8.page_setup.fitToWidth = 1
ws8.page_setup.fitToHeight = 0
ws8.sheet_properties.pageSetUpPr.fitToPage = True

# ---- 自检 ----
assert len(all_ids) == len(set(all_ids)), "编号有重复"
for v in VIEWS:
    for row in v["rows"]:
        assert len(row.get("pts", "").split("\n")) >= 2, f"{row['id']} 要点少于 2 条"
assert len(all_ids) == sum(int(x[6]) for x in RUBRIC), "视角条数与 08 页合计不一致"

wb.properties.title = "银行知识库评测集 · 分视角题型示例"
wb.properties.subject = "IBM zopeneditor-sample 实测"
wb.properties.creator = "评测集"
wb.properties.lastModifiedBy = "评测集"
wb.save(OUT)
print(f"[OK] 已生成 {OUT}")
print(f"     视角页 {len(VIEWS)} 个，共 {len(all_ids)} 题；缺口 {len(GAPS)} 条；对照 {len(MATRIX)} 行；口径 {len(RUBRIC)} 行")
for v in VIEWS:
    print(f"     {v['sheet']}: {len(v['rows'])} 题")
