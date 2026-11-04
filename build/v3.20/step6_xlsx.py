# -*- coding: utf-8 -*-
"""v3.20 第六步：验收标准清单补三条需求（15 条 AC）。

行结构：1~3 行标题、4 行表头、5 行起数据，一条 AC 一行、需求描述逐行重复。
序号与 AC 编号全局连续，插行后需整体重排。

插行必须走 Excel COM（openpyxl 不搬样式），写入值也用 COM 批量赋值以避免逐格调用。
FR-STO-31 与 FR-STO-32 同属 8.3.4 图谱建立、文档里前后相邻，按一个块插在
FR-STO-27 之后，块内各条按文档顺序回填。
"""
import os
import shutil

import openpyxl
import win32com.client as win32
from docx import Document

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCX = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20.docx")
SRC = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19_验收标准清单.xlsx")
DST = os.path.join(ROOT, "release", "v3.20", "核心系统迁移知识库需求规格书_v3.20_验收标准清单.xlsx")

AC = {
    "FR-STO-31": [
        "解析含同一程序内两段代码分别访问两个文件的样例（§1500 读 AUTHFILE、§1600 写 ACCTFILE）后，"
        "图谱中该程序的两个段落节点各带一条「读写」边指向对应文件；以「只在程序节点上有一条读写边、"
        "两个文件都挂在它下面」为反例验收。",
        "查询「ACCTFILE 被哪一段代码写过」，返回段落编号与行号区间，且区间覆盖实际 WRITE 语句所在行；"
        "只返回程序名、不给段落与行号的判为不通过。",
        "程序级「谁用过 ACCTFILE」由段落边沿「包含」向上聚合得出，与直接查段落级边得到的程序集合一致；"
        "以「程序级另存一份边、与段落级边不一致」为反例验收。",
        "JCL 含两个 EXEC 步骤、各自访问不同数据集时，两条「读写」边分别挂在对应作业步骤节点上；"
        "以「两条边都挂作业节点、分不出是哪个步骤访问」为反例验收。",
        "只能定位到程序整体的资料，允许边落程序节点且证据位置注明「仅到程序」；以「证据位置留空」"
        "或「填一个与该程序内任何语句都不对应的行号」为反例验收。",
    ],
    "FR-STO-32": [
        "以文件节点为起点提问「ACCTFILE 被哪些程序读写」，返回访问它的程序与段落集合；"
        "返回空集或只返回该文件自身的判为不通过。",
        "以程序节点为起点分别提问「它调用了谁」与「谁调用了它」，前者返回下游程序集合、后者返回上游"
        "程序集合，两者互不相同；以「两个方向返回同一集合」为反例验收。",
        "「包含」只沿出边：由程序取到其段落，不得由段落反查程序后再当作链路扩展对象；"
        "以「闭包结果中出现同程序下与该问题无关的段落」为反例验收。",
        "「归属」不参与闭包扩展：以「闭包结果中出现同一应用系统下与该问题无关的对象」为反例验收。",
        "反向得到的路径在应答中标注方向并说明「由被访问方反查得到」；以「把反向路径展示为"
        "正向调用顺序」为反例验收。",
    ],
    "FR-QA-24": [
        "提问中直接出现程序名、文件名或字段名时，检索入口为该对象节点，候选集以该节点的闭包为范围；"
        "以「命中结果中不含该对象节点、只按措辞取到无关段落」为反例验收。",
        "同名对象存在于两个应用系统时，并列呈现两侧结果并各标所属应用系统；以「只给其中一个应用系统"
        "的结果、不说明存在同名对象」为反例验收。",
        "提问中的对象名在库中不存在时，回答说明未匹配到该对象并退回相似度检索；以「静默改按措辞取"
        "候选、不说明未匹配」为反例验收。",
        "锚定后仍受 FR-STO-21 的知识版本与应用系统限定：以「锚定绕过应用系统限定、取到其他应用系统的"
        "同名字段」为反例验收。",
        "锚定只决定入口，事实类结论仍只引 L1a 或 L1b：以「锚定后把 L3 生成内容作为事实结论的依据」"
        "为反例验收。",
    ],
}

# 需求编号 -> (标题, 章节, 功能组, 优先级)
META = {
    "FR-STO-31": ("边的节点粒度", "8 入库建图", "8.3.4 图谱建立", "必须"),
    "FR-STO-32": ("可达闭包的遍历方向", "8 入库建图", "8.3.4 图谱建立", "必须"),
    "FR-QA-24": ("问题中的对象名锚定", "10 问答应答", "10.3.4 来源约束", "必须"),
}

# 插入块：插在哪个需求之后、块内按文档顺序排列的编号
BLOCKS = [
    ("FR-STO-27", ["FR-STO-31", "FR-STO-32"]),
    ("FR-QA-22", ["FR-QA-24"]),
]


def desc_of(code):
    """从需求书取描述段原文（去掉「描述：」前缀）。"""
    doc = Document(DOCX)
    ps = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    i = next(k for k, t in enumerate(ps) if t.startswith(code))
    for t in ps[i + 1:i + 5]:
        if t.startswith("描述："):
            return t[len("描述："):]
        if t.startswith("FR-"):
            break
    raise KeyError(code)


def main():
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    shutil.copyfile(SRC, DST)

    wb = openpyxl.load_workbook(DST)
    ws = wb["验收清单"]
    last = ws.max_row
    rows_by = {}
    for r in range(5, last + 1):
        rows_by.setdefault(ws.cell(r, 2).value, []).append(r)
    n_old = last - 4
    wb.close()

    # 升序算出每块的最终起始行（每块加上它之前所有块的行数）
    plan, off = [], 0
    for after, codes in BLOCKS:
        at = max(rows_by[after]) + 1
        n = sum(len(AC[c]) for c in codes)
        plan.append((at, codes, n, at + off))
        off += n
    print("插入计划:", [(p[0], p[3], p[1]) for p in plan])

    # 块内各条的最终行号
    final = {}
    for _, codes, _, start in plan:
        cur = start
        for c in codes:
            final[c] = cur
            cur += len(AC[c])
    print("各条起始行:", final)

    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    book = xl.Workbooks.Open(DST)
    try:
        sh = book.Worksheets("验收清单")
        # 行号大的先插，避免漂移
        for at, codes, n, _ in sorted(plan, reverse=True):
            sh.Rows("%d:%d" % (at, at + n - 1)).Insert(-4121)
            print("  在 %d 行插入 %d 行 -> %s" % (at, n, codes))

        total = n_old + sum(len(v) for v in AC.values())
        sh.Range("A5:A%d" % (5 + total - 1)).Value2 = [[i + 1] for i in range(total)]
        sh.Range("H5:H%d" % (5 + total - 1)).Value2 = [["AC-%d" % (i + 1)] for i in range(total)]

        for code, (title, ch, grp, prio) in META.items():
            desc = desc_of(code)
            target = list(range(final[code], final[code] + len(AC[code])))
            for r in target:
                assert not sh.Cells(r, 2).Value, ("第 %d 行非空，插入点算错" % r)
            for k, r in enumerate(target):
                sh.Cells(r, 2).Value = code
                sh.Cells(r, 3).Value = title
                sh.Cells(r, 4).Value = ch
                sh.Cells(r, 5).Value = grp
                sh.Cells(r, 6).Value = prio
                sh.Cells(r, 7).Value = desc
                sh.Cells(r, 9).Value = AC[code][k]
                sh.Cells(r, 10).Value = "未执行"
            print("  回填", code, len(target), "行")

        # 标题行
        n_req = 141 + len(META)
        sh.Cells(2, 1).Value = ("对应文档：核心系统迁移知识库需求规格书_v3.20.docx\u3000｜\u3000"
                                "本清单是验收标准的唯一出处，需求书正文不重复列示")
        sh.Cells(3, 1).Value = ("共 %d 条需求 / %d 条验收标准\u3000｜\u3000"
                                "优先级：必须 111 · 应该 32 · 可选 1\u3000｜\u3000"
                                "交付批次：P1 108 · P2 35 · P3 1" % (n_req, total))
        print("标题行:", sh.Cells(3, 1).Value)

        book.Save()
    finally:
        book.Close(False)
        xl.Quit()
    print("saved", DST)


if __name__ == "__main__":
    main()
