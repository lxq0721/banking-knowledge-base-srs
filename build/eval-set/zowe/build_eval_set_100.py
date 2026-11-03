# -*- coding: utf-8 -*-
"""
生成《银行知识库评测集模板 v0.1》
零公式设计：统计页全部静态值 + DataValidation 下拉 + 条件格式。
"""
import os
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter

# 默认写到仓库 eval-set/ 下；换路径时用环境变量 EVAL_OUT 覆盖
OUT = os.environ.get("EVAL_OUT") or os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "..", "..", "eval-set", "银行知识库评测集_ZOpenEditor实测100题_v1.0.xlsx"))

# ---------- 样式常量 ----------
FN = "微软雅黑"
C_DARK = "1F3864"      # 主标题
C_GRP = "2E5C8A"       # 分组表头
C_HDR = "4472C4"       # 字段表头
C_ALT = "F2F2F2"       # 示例行底
C_EG = "FFF9E6"        # 示例区浅黄
C_LINE = "BFBFBF"
C_NOTE = "595959"

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
    else:
        c.border = BORDER
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

def est_height(texts_widths, line_h=13.5, pad=6, minh=20):
    """texts_widths: [(文本, 列宽)]，按中文占 2 单位估算行数"""
    lines = 1
    for t, w in texts_widths:
        if not t:
            continue
        s = str(t)
        units = sum(2 if ord(ch) > 127 else 1 for ch in s)
        cap = max(4, (w - 1.2) * 1.05)   # 每行可容纳单位数
        # 手工换行符也要算
        n = 0
        for seg in s.split("\n"):
            u = sum(2 if ord(ch) > 127 else 1 for ch in seg)
            n += max(1, math.ceil(u / cap))
        lines = max(lines, n)
    return max(minh, lines * line_h + pad)

wb = Workbook()
wb.remove(wb.active)

# =====================================================================
# Sheet 1  填写说明
# =====================================================================
ws = wb.create_sheet("填写说明")
ws.sheet_view.showGridLines = False
widths = [4, 20, 96, 22]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws, r, 1, "银行知识库 · 评测集（Z Open Editor 实测 100 题）", span=4, size=18, height=34)
r += 1
title_cell(ws, r, 1,
           "样本库：IBM/zopeneditor-sample 主干 @8f98353 ｜ 版本 v1.0 ｜ 每一条期望值都能按「文件:行号」在仓库里反查，无占位数据 ｜ 配套《银行知识库测试方案 v0.6》4.3 / 5.3 章与附录 A",
           span=4, size=9, color=C_NOTE, bold=False, height=22)
ws.row_dimensions[2].outlineLevel = 0

r = 4
title_cell(ws, r, 1, "一、这份表干什么用", span=4, size=12, fill="DEEAF6", height=24)
r += 1
body_cell(ws, r, 1, "评测集是效果类指标的唯一判据。检索效果（L2）与输出质量（L3）无法靠肉眼判合格——同一个回答，两个人可能给出不同分数。"
                    "本表把「问题、判分依据、预期属性」三样东西固定成结构化记录，使每次跑批结果可比、可追溯、可复现。"
                    "本版 100 题以 IBM 公开示例库为语料，答案全部落在真实代码行上，因此可以直接拿这份表检验检索与生成，而不必先等客户资料到位。", span=4, size=9)
ws.row_dimensions[r].height = est_height([("x" * 108, 96)], minh=30)

r += 2
title_cell(ws, r, 1, "二、三条使用原则", span=4, size=12, fill="DEEAF6", height=24)
r += 1
rules = [
    ("1. 先定范围，再出题",
     "出题前先明确被测子系统与基线版本。题目必须能挂到某个已冻结的基线上（见「集合层元信息」页），否则重跑时库已经变了，结果不可比。"),
    ("2. 无答案题必须逐条核实",
     "第 4 类（库内不存在）是幻觉专项的核心。凭印象觉得「库里应该有」而实际没有，会把系统正常作答误判为幻觉。每条都要到库中反查，反查人写进备注。"),
    ("3. 答案要点必须可拆",
     "期望答案要点按「①②③」逐条写，一条一个可判定的事实。写成一段话就无法按 0 / 0.5 / 1 打分，评分会退化为印象分。"),
]
for name, desc in rules:
    body_cell(ws, r, 1, name, span=2, size=9, bold=True, color=C_DARK, valign="center")
    body_cell(ws, r, 3, desc, span=2, size=9)
    ws.row_dimensions[r].height = est_height([(desc, 96 + 22)], minh=30)
    r += 1

r += 1
title_cell(ws, r, 1, "三、页面结构", span=4, size=12, fill="DEEAF6", height=24)
r += 1
hdr = ["№", "页面", "内容", "什么时候看"]
for i, h in enumerate(hdr, 1):
    title_cell(ws, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=22, halign="center")
r += 1
pages = [
    ("评测集", "主表。17 个字段、110 行容量（100 题实证 + 10 行余量）。本版没有「示例行」，每一行都是可直接跑的真实题。", "跑批与维护"),
    ("分类配比与进度", "9 类问题 + 跨类标签的目标条数、已填条数、目标占比。用于检查题目结构有没有偏科。", "出题阶段与出题后"),
    ("字段与填写规范", "每个字段的可选值、格式要求、缺失后果。拿不准怎么填就看这页。", "填写时对照"),
    ("示例题", "3 条真实题竖排展开，便于业务确认人逐字段过目；程序名、字段名、行号都对应仓库实际代码。", "给业务方过目时"),
    ("集合层元信息", "基线版本、环境版本、入库规模、确认人名册。这些字段挂在集合上而不是单题上，漏了会导致重跑不可比。", "每轮跑批前登记"),
    ("执行记录", "一行一次执行。跑批后回填实际输出、命中与漏掉的要点、判定结果。与主表按编号关联，不覆盖题目本身。", "跑批后回填"),
    ("脚本接口与判定口径", "导出什么格式、接口怎么调、脚本判哪些、哪些必须人判。要让开发照着就能接，不看这页会各写一套。", "接自动化前"),
]
for pno, (p, d, w) in enumerate(pages, 1):
    body_cell(ws, r, 1, pno, size=9, halign="center", valign="center")
    body_cell(ws, r, 2, p, size=9, bold=True, color=C_DARK)
    body_cell(ws, r, 3, d, size=9)
    body_cell(ws, r, 4, w, size=9, color=C_NOTE)
    ws.row_dimensions[r].height = est_height([(d, 96), (p, 20)], minh=28)
    r += 1

r += 1
ws.row_breaks.append(Break(id=r - 1))      # 「四」另起一页，避免表格跨页断裂
title_cell(ws, r, 1, "四、必须提醒的四件事", span=4, size=12, fill="FBE4E4", height=24)
r += 1
warns = [
    ("题库来自公开示例库，不是客户资料",
     "100 题全部基于 IBM/zopeneditor-sample（主干 8f98353）的真实代码：COPYBOOK、COBOL、PL/I、HLASM、REXX、JCL、zapp.yaml。"
     "行号可直接反查，但它是公开样例、不是行内系统——验证的是检索与判分能力，不能代替客户代码的效果结论。"),
    ("不是所有题都要跑",
     "第 7 类（跨模态）仅在库内含 BMS 绿屏图、系统截图或流程图时才出；若本批无此类资料，整类裁剪，不必凑数。"),
    ("一道权限题不等于一行",
     "第 9 类与涉及可见范围的题目，需按角色（业务专家 / 架构师 / 开发 / 审核人）× 地区 × 子系统分次执行。"
     "表中仍记一行，实际执行次数写在备注列，否则缺陷单会缺「哪个角色下出的问题」这个关键信息。"),
    ("本文件不含公式",
     "为兼容预览器显示，所有统计均为静态值，不随录入自动变化。如需自动汇总，可在 Excel 中自行补充 COUNTIF 类公式。"),
]
for name, desc in warns:
    body_cell(ws, r, 1, name, span=2, size=9, bold=True, color="C00000", valign="center")
    body_cell(ws, r, 3, desc, span=2, size=9)
    ws.row_dimensions[r].height = est_height([(desc, 96 + 22)], minh=28)
    r += 1

r += 1
title_cell(ws, r, 1, "五、与测试方案的对齐点", span=4, size=12, fill="DEEAF6", height=24)
r += 1
align = [
    ("问题类型 9 类", "方案 4.3「问题集分类维度」"),
    ("事实基准 / 期望来源等级", "方案 5.3「事实基准与来源分层」、需求书表 7-5 五级"),
    ("需分角色执行", "权限维度：子系统 × 地区 × 角色；第 9 类题按角色分次执行"),
    ("期望内容状态", "需求书表 7-3 四态：可取用 / 已修正 / 有争议 / 已废止"),
    ("冲突类题目的期望", "需求书 FR-GEN-31 判据 + 表 6-8 三种情形处置"),
    ("评测集字段与规模", "方案附录 A、5.2 测试数据 D3（50~100 条，先手工跑通）"),
    ("机器判定规则", "方案 5.2 D4「可自动化优先」；无法结构化判定的仍走人工"),
    ("题库语料", "IBM/zopeneditor-sample @8f98353；只读样本，不改仓库、不留副作用"),
]
for a, b in align:
    body_cell(ws, r, 1, a, span=2, size=9, valign="center")
    body_cell(ws, r, 3, b, span=2, size=9, color=C_NOTE, valign="center")
    ws.row_dimensions[r].height = 20
    r += 1

r += 1
title_cell(ws, r, 1, "六、脚本跑批怎么接（三步）", span=4, size=12, fill="DEEAF6", height=24)
r += 1
steps = [
    ("① 导出",
     "python export_jsonl.py 银行知识库评测集_ZOpenEditor实测100题_v1.0.xlsx —— 生成 eval_set.jsonl（一条题一行），同时打印校验报告：缺必填、编号重复、要点拆不开、判定规则写错格式。有错先改表，别带着跑。"),
    ("② 跑批",
     "python run_eval.py --input eval_set.jsonl --endpoint http://<被测服务>/qa —— 逐条发问、记实际输出、耗时、命中与漏掉的要点，结果写 results.csv。加 --dry-run 可先用内置桩把判定逻辑走一遍，不碰网络。"),
    ("③ 回填",
     "把 results.csv 的判定结果按编号抄进「执行记录」页。脚本判不了的会标「待人工」——这是待办，不是通过。"),
]
for a, b in steps:
    body_cell(ws, r, 1, a, span=2, size=9, bold=True, color=C_DARK, valign="center")
    body_cell(ws, r, 3, b, span=2, size=9)
    ws.row_dimensions[r].height = est_height([(b, 96 + 22)], minh=28)
    r += 1
r += 1
title_cell(ws, r, 1, "脚本能判什么、不能判什么", span=4, size=10, color="C00000", fill="FBE4E4", height=22)
r += 1
body_cell(ws, r, 1,
          "能自动判的只有硬事实：字段名 / 文件名 / 行号有没有出现、该出现的级别词有没有标、有没有 5xx 或超时、无答案题有没有编造。"
          "判不了的是语义质量——「说得对不对、完整不完整、冲突处置合不合规」，这些仍需人工过。"
          "因此每批仍要留 20%~30% 人工抽检，不能因为脚本全绿就签字。\n"
          "生成日期 2026-11-02 ｜ 语料 IBM/zopeneditor-sample @8f98353 ｜ 配套方案 v0.6 ｜ 100 题，行号可逐条反查",
          span=4, size=9)
ws.row_dimensions[r].height = est_height([("x" * 132, 96 + 22), ("y" * 60, 96 + 22)], minh=30)

# =====================================================================
# Sheet 2  评测集（主表）
# =====================================================================
ws2 = wb.create_sheet("评测集")
COLS = [
    # (字段名, 列宽, 组)
    ("编号", 14, "题面"),
    ("所属范围", 15, "题面"),
    ("问题原文", 38, "题面"),
    ("问题类型", 15, "题面"),
    ("难度", 9, "题面"),
    ("需分角色执行", 13, "题面"),
    ("库内是否存在该信息", 13, "判分依据"),
    ("事实基准", 13, "判分依据"),
    ("期望答案要点", 40, "判分依据"),
    ("期望出处", 20, "判分依据"),
    ("期望来源等级", 20, "预期属性"),
    ("期望内容状态", 12, "预期属性"),
    ("重点校验项", 26, "预期属性"),
    ("机器判定规则", 20, "预期属性"),
    ("业务确认人", 11, "管理"),
    ("确认日期", 11, "管理"),
    ("备注", 16, "管理"),
]
for i, (name, w, grp) in enumerate(COLS, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# 第一行分组表头
groups = {}
for i, (_, _, g) in enumerate(COLS, 1):
    groups.setdefault(g, []).append(i)
for g, idxs in groups.items():
    title_cell(ws2, 1, idxs[0], g, span=len(idxs), size=10, color="FFFFFF",
               fill=C_GRP, height=22, halign="center")
# 第二行字段名
for i, (name, w, _) in enumerate(COLS, 1):
    title_cell(ws2, 2, i, name, size=9, color="FFFFFF", fill=C_HDR, height=32, halign="center")

# ---- 题库：IBM zopeneditor-sample 实测 100 题（真实行号，无占位）----
from q_zowe_a import ZQ_A
from q_zowe_b import ZQ_B

EG = ZQ_A + ZQ_B
assert len(EG) == 100, f"题库应为 100 题，实际 {len(EG)}"
assert all(len(x) == len(COLS) for x in EG), "题元组长度与列数不一致"
assert len({x[0] for x in EG}) == 100, "编号有重复"

HDR_ROW = 2
start = HDR_ROW + 1

EMPTY_ROWS = 10            # 100 题 + 10 行余量，供后续增补
total_rows = len(EG) + EMPTY_ROWS

for k in range(total_rows):
    r = start + k
    is_eg = k < len(EG)
    data = list(EG[k]) if is_eg else [""] * len(COLS)
    fill = C_EG if is_eg else None
    for i, val in enumerate(data, 1):
        center = i in (1, 4, 5, 6, 7, 12, 15, 16)
        body_cell(ws2, r, i, data[i - 1], size=9, fill=fill,
                  halign="center" if center else "left",
                  valign="center" if center else "top", wrap=True)
    if is_eg:
        ws2.row_dimensions[r].height = est_height(
            [(data[2], 38), (data[8], 40), (data[12], 26), (data[13], 20)],
            line_h=12.5, pad=6, minh=34)
    else:
        ws2.row_dimensions[r].height = 30

ws2.freeze_panes = "D3"
ws2.auto_filter.ref = f"A{HDR_ROW}:Q{start + total_rows - 1}"
ws2.sheet_view.zoomScale = 85

# ---- 数据验证（下拉）----
def add_dv(ws, formula, rng):
    dv = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True,
                        showDropDown=False)
    dv.error = "请从下拉列表中选择"
    dv.errorTitle = "取值不在范围内"
    dv.prompt = "从下拉中选择"
    dv.promptTitle = "可选值"
    ws.add_data_validation(dv)
    dv.add(rng)

last = start + total_rows - 1
add_dv(ws2, "1 精准命中,2 同义改写+跨语言,3 信息分散（多跳）,4 知识库不存在（幻觉）,5 冲突信息,6 异常与边界,7 跨模态,8 多轮对话,9 元数据过滤", f"D{start}:D{last}")
add_dv(ws2, "简单,中等,困难", f"E{start}:E{last}")
add_dv(ws2, "否,是", f"F{start}:F{last}")
add_dv(ws2, "是,否,—", f"G{start}:G{last}")
add_dv(ws2, "原始资料,开发补全物,推断,—", f"H{start}:H{last}")
add_dv(ws2, "代码直出,引自文档,人工修正,模型生成·核验通过,模型生成·核验未通过,—", f"K{start}:K{last}")
add_dv(ws2, "可取用,已修正,有争议,已废止,—", f"L{start}:L{last}")

# ---- 条件格式 ----
red_fill = PatternFill("solid", fgColor="FFC7CE")
yellow_fill = PatternFill("solid", fgColor="FFEB9C")
green_fill = PatternFill("solid", fgColor="C6EFCE")
red_font = Font(name=FN, size=9, color="9C0006", bold=True)

# 无答案题：G 列 = 否 → 红底（幻觉专项必须显著）
ws2.conditional_formatting.add(
    f"G{start}:G{last}",
    CellIsRule(operator="equal", formula=['"否"'], fill=red_fill, font=red_font))
# 来源等级：核验未通过 红、核验通过 黄
ws2.conditional_formatting.add(
    f"K{start}:K{last}",
    CellIsRule(operator="equal", formula=['"模型生成·核验未通过"'], fill=red_fill, font=red_font))
ws2.conditional_formatting.add(
    f"K{start}:K{last}",
    CellIsRule(operator="equal", formula=['"模型生成·核验通过"'], fill=yellow_fill))
# 内容状态：有争议 / 已废止 黄底
ws2.conditional_formatting.add(
    f"L{start}:L{last}",
    CellIsRule(operator="equal", formula=['"有争议"'], fill=yellow_fill))
ws2.conditional_formatting.add(
    f"L{start}:L{last}",
    CellIsRule(operator="equal", formula=['"已废止"'], fill=yellow_fill))
# 难度：困难 加粗提示
ws2.conditional_formatting.add(
    f"E{start}:E{last}",
    CellIsRule(operator="equal", formula=['"困难"'], font=Font(name=FN, size=9, bold=True, color="C00000")))

# ---- 打印设置 ----
ws2.page_setup.orientation = "landscape"
ws2.page_setup.paperSize = 8      # A3
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0
ws2.sheet_properties.pageSetUpPr.fitToPage = True
ws2.print_title_rows = "1:2"
ws2.print_options.horizontalCentered = True

# =====================================================================
# Sheet 3  分类配比与进度
# =====================================================================
ws3 = wb.create_sheet("分类配比与进度")
ws3.sheet_view.showGridLines = False
for i, w in enumerate([4, 26, 50, 10, 10, 10, 26], 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws3, r, 1, "九类问题配比与填写进度", span=7, size=14, height=30)
r += 1
title_cell(ws3, r, 1, "本版 100 条（方案 5.2 D3 上限即 100）。下表「已填写」由题库实时统计得出，不含公式——改题后重跑生成脚本即可刷新。第 7 类为 0：仓库无图形资料，整类裁剪。",
           span=7, size=9, color=C_NOTE, bold=False, height=28)
r += 2

hdr = ["序号", "类别", "对应场景（本项目）", "目标条数", "已填写", "目标占比", "填写提示"]
for i, h in enumerate(hdr, 1):
    title_cell(ws3, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=24, halign="center")
r += 1

CATS = [
    ("1", "精准命中", "原文 / 原标识符提问：字段名、copybook、程序名、步骤名直接命中", 16,
     "基准线。不过关说明检索整体有问题（ACCT-BALANCE、TRAN-CODE、CMPLSAM2 等）"),
    ("2", "同义改写 + 跨语言", "中文说法 → 英文标识符：客户号、余额、累计订单、DD 名", 14,
     "本项目重点。whitespace 分词下中文召回英文标识符是高风险点"),
    ("3", "信息分散（多跳）", "答案要跨程序 + copybook + JCL + 配置才拼得出来", 16,
     "重点验漏节点：UPDATE 全链路、387 字节怎么算、596 与 600 怎么对"),
    ("4", "知识库不存在（幻觉）", "库内确无（CICS、DB2、VSAM、单测、CI），预期「未查询到」", 12,
     "全部 P0。已逐条 grep 反查，反查方式写备注"),
    ("5", "冲突信息", "COBOL 与 PL/I 定义不一致、README 与配置不一致、看似冲突实为并列", 14,
     "分两半：真冲突 +「看似冲突实为可同时成立」，后者是反向陷阱，用来验判据会不会被误用"),
    ("6", "异常与边界", "空提问、全空格、超长、只有关键词、特殊字符、指代不明", 8,
     "重点是不崩、有提示、不空转调用模型、不猜一个就答"),
    ("7", "跨模态", "按描述找 BMS 绿屏图 / 截图 / 流程图", 0,
     "本仓库无任何图形资料 → 整类裁剪，不凑数（方案 4.3 明确要求）"),
    ("8", "多轮对话", "追问依赖上一轮结果，验证上下文继承与指代消解", 10,
     "按「组」计，一组 2~3 轮；含「换过滤条件后不得残留上一轮」"),
    ("9", "元数据过滤", "限定目录 / 来源等级 / 内容状态 / 排除子树", 10,
     "验证过滤在检索阶段生效，而不是答完再筛"),
]
_cnt = {}
for _row in EG:
    _no = _row[3][0] if _row[3] and _row[3][0].isdigit() else "?"
    _cnt[_no] = _cnt.get(_no, 0) + 1
TARGET_TOTAL = sum(c[3] for c in CATS)
FILLED_TOTAL = sum(_cnt.values())
for c in CATS:
    filled = _cnt.get(c[0], 0)
    body_cell(ws3, r, 1, c[0], size=9, halign="center", valign="center")
    body_cell(ws3, r, 2, c[1], size=9, bold=True, color=C_DARK, valign="center")
    body_cell(ws3, r, 3, c[2], size=9)
    body_cell(ws3, r, 4, c[3], size=9, halign="center", valign="center")
    body_cell(ws3, r, 5, filled, size=9, halign="center", valign="center")
    body_cell(ws3, r, 6, f"{c[3] / TARGET_TOTAL:.0%}", size=9, halign="center", valign="center")
    body_cell(ws3, r, 7, c[4], size=9, color=C_NOTE)
    ws3.row_dimensions[r].height = est_height([(c[2], 50), (c[4], 26)], line_h=12.5, minh=30)
    r += 1

# 合计行
body_cell(ws3, r, 1, "", size=9, fill="DEEAF6")
body_cell(ws3, r, 2, "合计（9 类）", size=9, bold=True, fill="DEEAF6", valign="center")
body_cell(ws3, r, 3, f"目标 {TARGET_TOTAL} 条；本版实填 {FILLED_TOTAL} 条（全部来自仓库实测），第 7 类裁剪故为 0",
          size=9, fill="DEEAF6")
body_cell(ws3, r, 4, TARGET_TOTAL, size=9, bold=True, fill="DEEAF6", halign="center", valign="center")
body_cell(ws3, r, 5, FILLED_TOTAL, size=9, bold=True, fill="DEEAF6", halign="center", valign="center")
body_cell(ws3, r, 6, "100%", size=9, fill="DEEAF6", halign="center", valign="center")
body_cell(ws3, r, 7, "跨类标签题另计，见下", size=9, fill="DEEAF6")
ws3.row_dimensions[r].height = 26
r += 2

ws3.row_breaks.append(Break(id=r - 1))      # 跨类标签另起一页，避免主表跨页断裂
title_cell(ws3, r, 1, "跨类标签（不单独算一类，与上表重叠计数）。本版重点看两组：真冲突几条、看似冲突几条——后者用来验判据会不会被误用", span=7, size=10, fill="DEEAF6", height=24)
r += 1
for i, h in enumerate(["№", "标签", "验证什么", "目标条数", "已填写", "占比", "说明"], 1):
    title_cell(ws3, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=22, halign="center")
r += 1
_n_conf = sum(1 for _r in EG if "真冲突" in _r[16])
_n_seem = sum(1 for _r in EG if "看似冲突" in _r[16])
_n_doc = sum(1 for _r in EG if _r[10] == "引自文档")
_n_obsolete = sum(1 for _r in EG if "注释" in _r[8])
TAGS = [
    ("真冲突", "同一对象同一属性取值互斥（长度、词义、库组名）", _n_conf, _n_conf,
     "判成「详略差异」即漏判，是这批题的主要失分点"),
    ("看似冲突实为并列", "可同时成立（不同作业、不同数据集、载体开销、编译选项）", _n_seem, _n_seem,
     "反向验证：判据用得对不对，看这批会不会被误报成冲突"),
    ("来源等级", "代码直出 vs 引自文档，标注是否与来源一致", _n_doc, _n_doc,
     "本版无「模型生成」类题：仓库里没有 LLM 生成物，等级验证靠文档那一条"),
    ("内容状态", "被注释掉的旧定义会不会被当现行定义", _n_obsolete, _n_obsolete,
     "CUSTCOPY 第 29/31/37 行的 COMP-3 旧写法是最典型的一处"),
]
for tno, t in enumerate(TAGS, 1):
    body_cell(ws3, r, 1, tno, size=9, halign="center", valign="center")
    body_cell(ws3, r, 2, t[0], size=9, bold=True, color=C_DARK, valign="center")
    body_cell(ws3, r, 3, t[1], size=9)
    body_cell(ws3, r, 4, t[2], size=9, halign="center", valign="center")
    body_cell(ws3, r, 5, t[3], size=9, halign="center", valign="center")
    body_cell(ws3, r, 6, "—", size=9, halign="center", valign="center")
    body_cell(ws3, r, 7, t[4], size=9, color=C_NOTE)
    ws3.row_dimensions[r].height = est_height([(t[1], 50), (t[4], 26)], line_h=12.5, minh=28)
    r += 1
r += 1

r += 1
title_cell(ws3, r, 1, "结构自检（出题完成后逐项确认）", span=7, size=10, fill="DEEAF6", height=24)
r += 1
checks = [
    "第 4 类（无答案）占比 ≥ 10%，且每条都已到库中反查确认「确实没有」",
    "第 9 类与需分角色的题目，都标明了按哪几个角色 / 范围分次执行（本版 9 条）",
    "每条题的「期望答案要点」都能拆成 2 个以上可独立判定的事实点",
    "每条题的「期望出处」都精确到文件:行号或文档定位符，没有只写文件名的",
    "填「引自文档」或「模型生成」的题，都在「重点校验项」里写明了标注要求",
    "所有题目都能对应到「集合层元信息」页登记的同一基线版本",
    "没有一条题目是「凭印象觉得库里有」但没有反查过的",
]
for c in checks:
    body_cell(ws3, r, 1, "□", size=11, halign="center", valign="center")
    body_cell(ws3, r, 2, c, span=6, size=9)
    ws3.row_dimensions[r].height = 22
    r += 1

# =====================================================================
# Sheet 4  字段与填写规范
# =====================================================================
ws4 = wb.create_sheet("字段与填写规范")
ws4.sheet_view.showGridLines = False
for i, w in enumerate([4, 20, 12, 9, 26, 44, 32], 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws4, r, 1, "字段与填写规范", span=7, size=14, height=30)
r += 1
title_cell(ws4, r, 1, "共 17 个字段，按「题面 / 判分依据 / 预期属性 / 管理」四组排列。第三组是本项目区别于通用 RAG 评测的关键，不可省略；其中「机器判定规则」供脚本判分，可空。",
           span=7, size=9, color=C_NOTE, bold=False, height=24)
r += 2
for i, h in enumerate(["序号", "字段", "所属组", "必填", "可选值 / 格式", "填写要求", "缺失后果"], 1):
    title_cell(ws4, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=24, halign="center")
r += 1

FIELDS = [
    ("编号", "题面", "是", "EV-<域>-<三位序号>\n如 EV-CORE-001", "域用子系统缩写：CORE 核心存款、GEN 知识生成、REV 审核、QA 问答。缺陷单必须引用该编号；脚本回填结果也按它对齐，不得重排或复用。", "缺陷无法回溯到题目"),
    ("所属范围", "题面", "是", "子系统 / 地区\n如 核心存款 / 境内", "一列同时写子系统与地区，中间用「 / 」分隔。子系统须与需求书划分一致——准出是按子系统分档给的（表 6-6 五档），没这列算不出分组达标率。", "无法按子系统统计达标率，准出无从判定"),
    ("问题原文", "题面", "是", "使用者视角自然语言", "写使用者真会问的话，不要写成「查询 XX 字段定义」。多轮题用「（第 1 轮）……（第 2 轮）……」分段。", "题目失真，效果指标无意义"),
    ("问题类型", "题面", "是", "下拉：1~9 类", "从 9 类中选一类；跨类标签题归入主类，另在备注写标签名。", "分类统计失真，看不出短板"),
    ("难度", "题面", "是", "下拉：简单/中等/困难", "简单＝单点单跳；中等＝需 2~3 处资料；困难＝多跳或需判冲突。", "无法分层设线"),
    ("需分角色执行", "题面", "否", "下拉：否 / 是", "默认「否」。填「是」的题要按角色（业务专家 / 架构师 / 开发 / 审核人）分次跑，角色清单与执行次数写备注列，行数不变。", "权限缺陷漏测，说不清是哪个角色下出的问题"),
    ("库内是否存在该信息", "判分依据", "是", "下拉：是 / 否 / —", "「否」专用于第 4 类幻觉题。填「否」之前必须到库中反查，确认真的没有。", "无答案题本身错误，会把正常作答判成幻觉"),
    ("事实基准", "判分依据", "是", "下拉：原始资料 / 开发补全物 / 推断 / —", "指判「答案对不对」所依据的东西：源码、DDL、真实文档＝原始资料；知识生成产物＝开发补全物；模型推断＝推断。", "无法判定答案正确性"),
    ("期望答案要点", "判分依据", "是", "①②③ 逐条列出", "一条一个可独立判定的事实点，按 0 / 0.5 / 1 计分。不要写成一段话，也不要写成「答案准确完整」这类无法判定的表述。", "打分退化为印象分，两人评不一致"),
    ("期望出处", "判分依据", "是", "文件:行号区间\n或文档定位符", "精确到行号或定位符，例：ACCTREC.cpy:214-216。无答案题填「—（应无出处）」。", "只能验有没有出处，验不出定位符准确率"),
    ("期望来源等级", "预期属性", "是", "下拉：代码直出 / 引自文档 / 人工修正 / 模型生成·核验通过 / 模型生成·核验未通过 / —",
     "直接填中文，不填 L1a 这类代号（代号对照见下方「等级对照」）。代码直出＝无标注；引自文档＝注「引自文档」；人工修正＝显示修正人与时间；模型生成·核验通过＝灰标「模型生成·未核实」；模型生成·核验未通过＝红标「仅供参考」。", "等级标注验证无法客观判定"),
    ("期望内容状态", "预期属性", "是", "下拉：可取用 / 已修正 / 有争议 / 已废止 / —",
     "绝大多数填「可取用」，只有专门验排除逻辑时才填「有争议」「已废止」。它和来源等级是两回事：代码直出说明「可信」，已废止说明「不能用」，分开填。", "「依据排除 100%」准出无题可验"),
    ("重点校验项", "预期属性", "否", "自由文本，分号分隔", "写这条题最容易出的错，例如「不得取旧版本」「不得只答程序名不答 copybook」「导出时不得省略标注」。", "关键负向点漏测"),
    ("机器判定规则", "预期属性", "否", "A|B|!C|/正则/\n（竖线分隔）",
     "给脚本判分用，可空。普通串＝回答里必须出现（忽略大小写）；! 前缀＝不得出现；/…/＝正则。只写字段名、文件名、行号、级别词这类硬事实——「回答是否通顺完整」写不进这里，那只能人工判。留空的题自动判不了，不影响表可用。",
     "该题只能人工判，自动跑批时会被标为待人工"),
    ("业务确认人", "管理", "是", "姓名", "对答案要点与期望出处负责的人。准入条件要求逐条经业务确认。", "答案要点无法追责"),
    ("确认日期", "管理", "是", "YYYY-MM-DD", "记录业务确认完成日，与基线版本日期对照。", "无法证明何时确认"),
    ("备注", "管理", "否", "自由文本", "写示例标记、反查人、实际执行次数、特殊说明。示例行须标「示例·占位数据」。", "关键上下文丢失"),
]
for i, f in enumerate(FIELDS, 1):
    body_cell(ws4, r, 1, i, size=9, halign="center", valign="center")
    body_cell(ws4, r, 2, f[0], size=9, bold=True, color=C_DARK, valign="center")
    body_cell(ws4, r, 3, f[1], size=9, halign="center", valign="center")
    body_cell(ws4, r, 4, f[2], size=9, halign="center", valign="center")
    body_cell(ws4, r, 5, f[3], size=9)
    body_cell(ws4, r, 6, f[4], size=9)
    body_cell(ws4, r, 7, f[5], size=9, color="C00000" if f[2] == "是" else C_NOTE)
    ws4.row_dimensions[r].height = est_height(
        [(f[3], 26), (f[4], 44), (f[5], 32)], line_h=12.5, pad=5, minh=26)
    r += 1

r += 1
title_cell(ws4, r, 1, "三个容易写错的字段", span=7, size=10, fill="FBE4E4", height=24)
r += 1
body_cell(ws4, r, 1, "期望答案要点", span=2, size=9, bold=True, color="C00000", valign="center")
body_cell(ws4, r, 3,
          "反例：「应准确说明该字段的定义位置与精度。」——无法打分。\n"
          "正例：「①定义在 ACCTREC.cpy ②字段名 ACCT-BAL ③PIC S9(11)V99 ④所属 ACCT-REC 记录」——四点各判 0 / 0.25，可复现。",
          span=5, size=9)
ws4.row_dimensions[r].height = 52
r += 2
body_cell(ws4, r, 1, "期望来源等级", span=2, size=9, bold=True, color="C00000", valign="center")
body_cell(ws4, r, 3,
          "反例：整条填「代码直出」，但答案里有一段「冲正需在当日轧账前处理」其实是模型推断的——按 FR-GEN-22，绑定与释义须分别标级，不得取低的一级覆盖整条。\n"
          "这种题要么拆成两条，要么在主等级下注明「其中某某段为模型生成」，不能一句「代码直出」盖过去。",
          span=5, size=9)
ws4.row_dimensions[r].height = 52
r += 2
body_cell(ws4, r, 1, "机器判定规则", span=2, size=9, bold=True, color="C00000", valign="center")
body_cell(ws4, r, 3,
          "反例：填「余额 精度」——回答里出现的是 ACCT-BAL、S9(11)V99，中文说法永远匹配不上，全批题会被判成不通过。\n"
          "正例：填「ACCT-BAL|S9(11)V99|ACCTREC.cpy」——只写回答里必然原样出现的硬事实。",
          span=5, size=9)
ws4.row_dimensions[r].height = 52

# =====================================================================
# Sheet 5  示例题
# =====================================================================
ws5 = wb.create_sheet("示例题")
ws5.sheet_view.showGridLines = False
for i, w in enumerate([4, 22, 86], 1):
    ws5.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws5, r, 1, "示例题（3 条，竖排展开：多跳 / 真冲突 / 无答案）", span=3, size=14, height=30)
r += 1
title_cell(ws5, r, 1,
           "这三条与主表同一份数据，不是另编的示范。程序名、字段名、行号都对应 IBM/zopeneditor-sample 主干，可逐条反查。",
           span=3, size=9, color="595959", bold=False, height=26)
r += 2

SAMPLES = [
    ("示例 1 · 多跳（答案要跨 copybook 逐段累加才拼得出来）", [
        ("编号", "EV-CPY-042"),
        ("所属范围", "COBOL 应用 / 本地仓库"),
        ("问题原文", "CUSTCOPY 描述的一条客户记录（A 类）总共多少字节？"),
        ("问题类型 / 难度", "3 信息分散（多跳） / 困难"),
        ("需分角色执行", "否"),
        ("库内是否存在该信息", "是"),
        ("事实基准", "原始资料（copybook 字段声明）"),
        ("期望答案要点", "①KEY 6（ID 5 + REC-TYPE 1）\n②NAME 17 + ACCT-BALANCE 9 + ORDERS-YTD 5 = 31\n"
                     "③ADDR 20 + CITY 14 + STATE 2 + COUNTRY 11 = 47\n④MONTH 9×12 = 108\n"
                     "⑤OCCUPATION 30 + NOTES 120 + DATA-1 5 + DATA-2 40 = 195\n⑥合计 387 字节"),
        ("期望出处", "COPYBOOK/CUSTCOPY.cpy:24-42"),
        ("期望来源等级", "代码直出"),
        ("期望内容状态", "可取用"),
        ("重点校验项", "必须给分段与合计；只报 387 说不出构成不得满分"),
        ("机器判定规则", "387|OCCURS 12|NOTES\n→ 三个串都出现才判通过；分段的正确性仍需人工抽检"),
        ("业务确认人 / 日期", "待填 / 待填"),
        ("备注", "实测·仓库主干"),
    ]),
    ("示例 2 · 真冲突（同一条客户记录，两种语言两套定义）", [
        ("编号", "EV-CFT-081"),
        ("所属范围", "跨语言对照 / 本地仓库"),
        ("问题原文", "PL/I 的 CUSTPLI 和 COBOL 的 CUSTCOPY 描述的是同一条客户记录吗？"),
        ("问题类型 / 难度", "5 冲突信息 / 困难"),
        ("需分角色执行", "否"),
        ("库内是否存在该信息", "是"),
        ("事实基准", "原始资料"),
        ("期望答案要点", "①判为冲突：同一对象、同一属性不能同时成立\n"
                     "②CUSTPLI 定长 80 字节（CUSTPLI.inc:9），CUSTCOPY 按 A 类算 387 字节（CUSTCOPY.cpy:24-42）\n"
                     "③字段集不同：PL/I 有 CITY 15、OCCUPATION 21；COBOL 有 ADDR/CITY/STATE/COUNTRY、MONTH×12、NOTES 120\n"
                     "④按 FR-GEN-31 记「来源不一致」，两侧并列，不合并成一条\n⑤不得按「详略差异」处理"),
        ("期望出处", "INCLUDES/CUSTPLI.inc:9-29; COPYBOOK/CUSTCOPY.cpy:24-42"),
        ("期望来源等级", "代码直出"),
        ("期望内容状态", "可取用"),
        ("重点校验项", "判成「详略差异」即错；两侧长度与字段差异都要给出"),
        ("机器判定规则", "冲突|CHAR(80)|387"),
        ("业务确认人 / 日期", "待填 / 待填"),
        ("备注", "实测·仓库主干·真冲突"),
    ]),
    ("示例 3 · 无答案（幻觉专项 P0；已 grep 反查确认库里确实没有）", [
        ("编号", "EV-NUL-061"),
        ("所属范围", "COBOL 应用 / 本地仓库"),
        ("问题原文", "SAM1 对应的 CICS 交易码是什么？"),
        ("问题类型 / 难度", "4 知识库不存在（幻觉） / 中等"),
        ("需分角色执行", "否"),
        ("库内是否存在该信息", "否（已反查：全仓无 EXEC CICS、无 CSD/SIT）"),
        ("事实基准", "原始资料（确认库内确无）"),
        ("期望答案要点", "①明确说明未查询到相关信息\n②不得给出任何交易码\n"
                     "③可提示已检索范围为仓库内 COBOL/JCL/配置，未见 CICS 相关定义"),
        ("期望出处", "—（应无出处）"),
        ("期望来源等级", "—"),
        ("期望内容状态", "—"),
        ("重点校验项", "编造交易码即 P0；把「未查到」写成「系统无 CICS」也算越界"),
        ("机器判定规则", "未查询到|!交易码为\n→ 必须出现「未查询到」，且不得出现「交易码为」这种给定结论"),
        ("业务确认人 / 日期", "待填 / 待填"),
        ("备注", "实测·仓库主干·已反查"),
    ]),
]

for si, (name, rows) in enumerate(SAMPLES):
    if si > 0:
        ws5.row_breaks.append(Break(id=r - 1))   # 一条示例一页，便于业务方逐条过
    title_cell(ws5, r, 1, name, span=3, size=10, color="FFFFFF", fill=C_GRP, height=24)
    r += 1
    for k, v in rows:
        body_cell(ws5, r, 1, "", size=9)
        body_cell(ws5, r, 2, k, size=9, bold=True, color=C_DARK)
        body_cell(ws5, r, 3, v, size=9)
        ws5.row_dimensions[r].height = est_height([(v, 86), (k, 22)], line_h=13, minh=22)
        r += 1
    r += 1

# =====================================================================
# Sheet 6  集合层元信息
# =====================================================================
ws6 = wb.create_sheet("集合层元信息")
ws6.sheet_view.showGridLines = False
for i, w in enumerate([4, 24, 46, 22, 16], 1):
    ws6.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws6, r, 1, "集合层元信息（每轮跑批前登记）", span=5, size=14, height=30)
r += 1
title_cell(ws6, r, 1,
           "这些信息挂在整份评测集上而不是单条记录上。漏了会出现「同一份题，两次结果不一样，说不清是系统变了还是库变了」——方案 5.1 的基线冻结机制就是为这个。"
           "尤其注意「嵌入模型及版本」：替换嵌入模型后同一段文字会得到不同向量、检索排序整体漂移，此时若只记了 LLM 版本，"
           "会把「换嵌入模型导致的召回下降」误判成「检索功能退化」，进而去改检索逻辑——方向从一开始就错了。",
           span=5, size=9, color=C_NOTE, bold=False, height=46)
r += 1
title_cell(ws6, r, 1, "前 9 项已按本版语料预填（公开仓库，可直接复用）；换语料必须整组重填。跑完批再回填 10~15 项。", span=5, size=9, color="C00000", bold=True, height=20)
r += 2
for i, h in enumerate(["序号", "项目", "填写要求", "本次填写", "填写人 / 日期"], 1):
    title_cell(ws6, r, i, h, size=9, color="FFFFFF", fill=C_HDR, height=24, halign="center")
r += 1

META = [
    ("基线版本号", "如「2026-10-20 基线」；源码快照存于 MinIO 的标识", "IBM/zopeneditor-sample @ 8f98353"),
    ("基线冻结日期", "YYYY-MM-DD；冻结后不得边跑边入库", "2026-09-11"),
    ("图库版本", "Neo4j 版本号；须与开发环境隔离", "待导入后登记"),
    ("LLM 模型及版本", "用于知识生成与问答的模型，写全版本标识", "待填"),
    ("嵌入模型及版本", "复现限定「同一知识版本 + 同一嵌入模型版本」，这个必须记", "待填"),
    ("被测服务版本", "提交号 / 镜像标签 / 发布批次号", "待填"),
    ("入库文件数与规模", "文件数、程序数、copybook 数、表数（对应表 6-6 齐备度档位）",
     "62 文件：cbl 5、pli 4、asm 3、rexx 4、inc 4、cpy 8、jcl 10、sh 4、yaml 3、json 4、md 2、数据 7"),
    ("子系统与地区分布", "如：核心存款 32 文件 / 支付结算 18 文件",
     "COBOL 应用 / PL/I 应用 / HLASM 应用 / REXX 脚本 / 作业与构建 / 工程配置 / 跨语言对照，均在「本地仓库」"),
    ("资料齐备度档位", "按需求书表 6-6 五档，标注本次属哪一档（仅程序 / +copybook / +JCL / 全量）",
     "程序 + copybook + JCL（无 DB2、无 CICS、无测试代码——第 4 类无答案题正是据此反查）"),
    ("题目总数", "本次实际执行条数（含被裁剪类别的说明）", ""),
    ("各类占比", "对照「分类配比与进度」页，说明有无偏离及原因", ""),
    ("裁剪说明", "哪几类被裁剪（如第 7 类无图形资料）、哪几类样本不足", ""),
    ("业务确认人名册", "逐条确认的责任人清单，或指向确认记录文件", ""),
    ("业务确认完成日期", "YYYY-MM-DD；未确认不得进入执行（准入条件）", ""),
    ("无效题处理记录", "执行中发现题本身有误的条数与编号，须剔除后重算分母", ""),
]
for i, m in enumerate(META, 1):
    body_cell(ws6, r, 1, i, size=9, halign="center", valign="center")
    body_cell(ws6, r, 2, m[0], size=9, bold=True, color=C_DARK, valign="center")
    body_cell(ws6, r, 3, m[1], size=9)
    body_cell(ws6, r, 4, m[2], size=9, fill="FFF9E6")
    body_cell(ws6, r, 5, "", size=9)
    ws6.row_dimensions[r].height = est_height([(m[1], 46)], line_h=12.5, minh=24)
    r += 1

# =====================================================================
# Sheet 7  执行记录
# =====================================================================
ws7 = wb.create_sheet("执行记录")
ws7.sheet_view.showGridLines = False
REC_COLS = [
    ("执行批次", 12), ("执行时间", 15), ("编号", 14), ("本次角色 / 范围", 18),
    ("问题原文（快照）", 30), ("实际输出（摘要）", 40), ("命中要点", 20), ("漏要点", 20),
    ("错引 / 越界引用", 18), ("标注是否齐全", 11), ("耗时 ms", 9), ("判定结果", 11),
    ("判定方式", 9), ("执行人", 9), ("证据位置", 18), ("备注", 16),
]
for i, (n, w) in enumerate(REC_COLS, 1):
    ws7.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws7, r, 1, "执行记录（一行一次执行）", span=len(REC_COLS), size=14, height=30)
r += 1
title_cell(ws7, r, 1,
           "与主表按「编号」关联，不覆盖题目本身。同一编号按角色或范围分次跑时，记多行——这正是「一道权限题不等于一行」的落地方式。"
           "前 3 行为示例（判定结果填的是示意值），正式使用时清空重填。",
           span=len(REC_COLS), size=9, color=C_NOTE, bold=False, height=26)
r += 1
for i, (n, w) in enumerate(REC_COLS, 1):
    title_cell(ws7, r, i, n, size=9, color="FFFFFF", fill=C_HDR, height=32, halign="center")
HDR7 = r
r += 1

REC_EG = [
    ("B20261102-1", "2026-11-02 10:12", "EV-CORE-001", "开发 / 核心存款·境内",
     "ACCT-BAL 这个字段是在哪个 copybook 里定义的？",
     "ACCT-BAL 定义在 ACCTREC.cpy 第 214~216 行，PIC S9(11)V99，属 ACCT-REC 记录。",
     "①③", "②④", "无", "齐全", 1840, "部分通过", "自动", "脚本", "ACCTREC.cpy:214-216", "示例·判定为示意值"),
    ("B20261102-1", "2026-11-02 10:13", "EV-CORE-002", "业务专家 / 核心存款·境内",
     "客户账户余额那个交易，余额的精度是多少？",
     "未查询到相关信息。",
     "—", "①②③", "无", "不适用", 2110, "不通过", "自动", "脚本", "—", "示例·中文提问未召回英文标识符"),
    ("B20261102-1", "2026-11-02 10:14", "EV-CORE-004", "开发 / 核心存款·境内",
     "这个系统支持数字人民币钱包的余额结转吗？",
     "知识库中未查询到相关信息，当前资料不包含数字人民币钱包相关内容。",
     "①", "—", "无", "不适用", 1520, "通过", "自动", "脚本", "—", "示例·「不支持」未出现，判通过"),
]
for row in REC_EG:
    for i, v in enumerate(row, 1):
        center = i in (1, 2, 3, 10, 11, 12, 13, 14)
        body_cell(ws7, r, i, v, size=9, fill=C_EG,
                  halign="center" if center else "left",
                  valign="center" if center else "top", wrap=True)
    ws7.row_dimensions[r].height = est_height([(row[5], 40), (row[4], 30)], line_h=12.5, minh=32)
    r += 1

for _ in range(27):
    for i in range(1, len(REC_COLS) + 1):
        center = i in (1, 2, 3, 10, 11, 12, 13, 14)
        body_cell(ws7, r, i, "", size=9,
                  halign="center" if center else "left",
                  valign="center" if center else "top", wrap=True)
    ws7.row_dimensions[r].height = 26
    r += 1

last7 = r - 1
add_dv(ws7, "通过,部分通过,不通过,待人工,无效题", f"L{HDR7 + 1}:L{last7}")
add_dv(ws7, "自动,人工", f"M{HDR7 + 1}:M{last7}")
add_dv(ws7, "齐全,缺失,不适用", f"J{HDR7 + 1}:J{last7}")

ws7.conditional_formatting.add(f"L{HDR7 + 1}:L{last7}",
    CellIsRule(operator="equal", formula=['"不通过"'], fill=red_fill, font=red_font))
ws7.conditional_formatting.add(f"L{HDR7 + 1}:L{last7}",
    CellIsRule(operator="equal", formula=['"待人工"'], fill=yellow_fill))
ws7.conditional_formatting.add(f"J{HDR7 + 1}:J{last7}",
    CellIsRule(operator="equal", formula=['"缺失"'], fill=red_fill, font=red_font))

ws7.freeze_panes = f"C{HDR7 + 1}"
ws7.auto_filter.ref = f"A{HDR7}:P{last7}"

# =====================================================================
# Sheet 8  脚本接口与判定口径
# =====================================================================
ws8 = wb.create_sheet("脚本接口与判定口径")
ws8.sheet_view.showGridLines = False
for i, w in enumerate([4, 22, 30, 62], 1):
    ws8.column_dimensions[get_column_letter(i)].width = w

r = 1
title_cell(ws8, r, 1, "脚本接口与判定口径", span=4, size=14, height=30)
r += 1
title_cell(ws8, r, 1,
           "这份表要能让开发不看正文就接上：导出什么、接口怎么调、判什么、不判什么、判完怎么落。三处口径定死，别各写一套。",
           span=4, size=9, color=C_NOTE, bold=False, height=22)
r += 2

def sec(title, color="DEEAF6"):
    global r
    title_cell(ws8, r, 1, title, span=4, size=11, fill=color, height=24)
    r += 1

def kv(k, v, k2="", v2="", h=None):
    global r
    body_cell(ws8, r, 1, "", size=9)
    body_cell(ws8, r, 2, k, size=9, bold=True, color=C_DARK)
    body_cell(ws8, r, 3, v)
    body_cell(ws8, r, 4, k2 if k2 else "", size=9)
    ws8.row_dimensions[r].height = h or est_height([(v, 30), (k, 22)], line_h=12.5, minh=24)
    r += 1

def full(k, v):
    global r
    body_cell(ws8, r, 1, "", size=9)
    body_cell(ws8, r, 2, k, size=9, bold=True, color=C_DARK)
    body_cell(ws8, r, 3, v, span=2)
    ws8.row_dimensions[r].height = est_height([(v, 30 + 62)], line_h=12.5, minh=24)
    r += 1

sec("一、导出物：eval_set.jsonl，一行一题")
full("导出命令", "python export_jsonl.py 银行知识库评测集模板_v0.1.xlsx → eval_set.jsonl + 校验报告")
full("为什么不用 xlsx 直接跑",
     "xlsx 是给人填的：有合并表头、有示例行、有下拉。脚本读它要处理一堆表格噪声，且字段一改就崩。"
     "导出成 JSONL 后脚本只认一种结构，表怎么排版都不影响跑批。")
full("字段映射（主表列 → JSON 键）",
     "编号→id｜所属范围→scope｜问题原文→question｜问题类型→type｜难度→difficulty｜"
     "需分角色执行→needs_role｜库内是否存在该信息→in_kb｜事实基准→basis｜"
     "期望答案要点→points（按 ①②③ 切成数组）｜期望出处→expect_sources（按分号切）｜"
     "期望来源等级→expect_level｜期望内容状态→expect_state｜机器判定规则→rules（解析成 must / must_not / regex）｜"
     "备注→note｜is_sample（备注含「示例」即为 true）")
full("示例行处理", "is_sample=true 的题默认跳过，加 --include-sample 才跑。避免拿编的期望值当真值。")

r += 1
sec("二、接口调用：默认约定，被测接口不同就只改 adapter 一处")
full("请求体", '{"question": "...", "filters": {"subsystem": "核心存款", "region": "境内"}, "role": "developer"}')
full("响应取什么", "answer 文本（必取）、sources[]（有就取，用于出处比对）、latency_ms、status_code")
full("分角色怎么跑", "needs_role=是 的题，按 --roles 传入的角色列表各跑一遍，结果各占一行，编号不变、行数增加")
full("超时与重试", "默认单条 60s 超时、不重试。超时记为「不通过（超时）」，不重试是为了暴露真实长尾，不是掩盖它")

r += 1
sec("三、判定算法：四步，顺序不能颠倒")
full("第 1 步 硬失败优先",
     "5xx / 连接失败 / 超时 → 直接判「不通过」，不再看内容。先判技术失败，避免把 502 当「答案不对」写进效果指标。")
full("第 2 步 无答案题专项",
     "in_kb=否：出现 must_not 词（如「不支持」）或未出现「未查询到」→ 判「不通过（幻觉）」。这类是 P0，单独统计。")
full("第 3 步 规则匹配",
     "rules 里每项：普通串＝要求出现（忽略大小写）；! 前缀＝要求不出现；/…/＝正则须命中。"
     "全部满足 → 通过；部分满足 → 部分通过（记命中率）；一条没满足 → 不通过。")
full("第 4 步 无规则则待人工",
     "rules 为空 → 判「待人工」，不是「通过」。这是硬规定：没判据就等于没判。")

r += 1
sec("四、结果状态：五个值，含义固定")
full("通过", "硬失败无、规则全满足。仍可被人工抽检推翻——脚本绿不等于签字。")
full("部分通过", "规则命中率在 (0, 1) 之间。记命中率，用于看趋势，不计入「通过率」分子。")
full("不通过", "硬失败、幻觉、或规则命中率为 0。")
full("待人工", "该题没有机器判定规则，或涉及语义质量（说得对不对、冲突处置是否合规）。")
full("无效题", "题本身有问题（库内实际有/无与登记不符、出处写错）。剔除后重算分母，并回改主表。")

r += 1
sec("五、脚本判不了的（必须留人工）", color="FBE4E4")
for t in [
    "答案是否「说清楚了」——要点都沾边但结论错，脚本一样判通过",
    "冲突题处置是否合规（该并列的并列了没有、该记疑点的记了没有）",
    "多轮题上下文是否被正确继承（第 2 轮答案对了但绑错了对象，脚本看不出来）",
    "解释通道的表述是否越界（把推断说成事实）",
]:
    body_cell(ws8, r, 1, "□", size=11, halign="center", valign="center")
    body_cell(ws8, r, 2, t, span=3, size=9)
    ws8.row_dimensions[r].height = est_height([(t, 22 + 30 + 62)], line_h=12.5, minh=22)
    r += 1

r += 1
sec("六、与准出指标的对应")
full("召回 / 定位符准确率", "由期望出处 expect_sources 与响应 sources 比对得出，脚本可算")
full("幻觉率", "第 4 类题的「不通过（幻觉）」条数 ÷ 该类总条数，目标 0")
full("依据排除", "expect_state 为「已废止 / 有争议」的题，回答依据里出现该条即判不通过，目标 0")
full("标注缺失", "expect_level 要求标注的题，回答里没有级别或「未核实」字样即判不通过，目标 0")
full("语义质量", "脚本不产出，只由人工抽检给分——报表里必须把这两类分开列，不能合成一个「总分」")

# ---------- 保存 ----------
# 统一打印设置：横向 + 缩放到一页宽；说明类页面宽度约 140 字符，纵向必然截断
for wsx in wb.worksheets:
    wsx.page_setup.orientation = "landscape"
    wsx.page_setup.fitToWidth = 1
    wsx.page_setup.fitToHeight = 0
    wsx.sheet_properties.pageSetUpPr.fitToPage = True
    wsx.print_options.horizontalCentered = True
    wsx.page_margins.left = 0.4
    wsx.page_margins.right = 0.4
    wsx.page_margins.top = 0.5
    wsx.page_margins.bottom = 0.5
    if wsx.title in ("评测集", "执行记录"):
        wsx.page_setup.paperSize = 8          # A3
    else:
        wsx.page_setup.paperSize = 9          # A4
    if wsx.title == "评测集":
        wsx.print_title_rows = "1:2"
    if wsx.title == "执行记录":
        wsx.print_title_rows = "3:3"
    if wsx.title == "字段与填写规范":
        wsx.print_title_rows = "4:4"          # 数据表跨页时重复表头（分类配比页的分页点已靠 row_break 控制，不重复表头）

# 兜底：未显式设高的行按 9pt 字统一给 18，避免默认 13.5 挤字
for wsx in wb.worksheets:
    for rr in range(1, wsx.max_row + 1):
        h = wsx.row_dimensions[rr].height
        if h is None or h < 18:
            wsx.row_dimensions[rr].height = 18

wb.properties.title = "银行知识库评测集 · Z Open Editor 实测 100 题"
wb.properties.subject = "语料：IBM/zopeneditor-sample @8f98353 · 100 题与判分依据"
wb.properties.creator = "测试组"
wb.properties.description = "配套《银行知识库测试方案 v0.6》；题库全部落在真实代码行上，零公式"
wb.properties.keywords = "评测集,RAG,检索效果,输出质量,幻觉专项"
wb.properties.category = "测试资产"
wb.save(OUT)
print("saved:", OUT)
print("sheets:", [s.title for s in wb.worksheets])
