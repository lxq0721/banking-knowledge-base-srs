# -*- coding: utf-8 -*-
"""v3.19 第五步：验收标准清单补三条需求。

行结构：1~3 行标题、4 行表头、5 行起数据，一条 AC 一行、需求描述逐行重复。
序号与 AC 编号全局连续，插行后需整体重排。

插行必须走 Excel COM（openpyxl 不搬样式），写入值也用 COM 批量赋值以避免逐格调用。
"""
import os
import shutil

import openpyxl
import win32com.client as win32
from docx import Document

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCX = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19.docx")
SRC = os.path.join(ROOT, "release", "v3.18", "核心系统迁移知识库需求规格书_v3.18_验收标准清单.xlsx")
DST = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19_验收标准清单.xlsx")

AC = {
    "FR-GEN-34": [
        "导入含交易定义（CSD/PCT）与主程序、子程序的资料并生成后，打开任一交易说明，其中列出的调用链"
        "与源码 CALL 边一致：链上每个程序都能在源码中找到调用它的语句，源码中存在的下游调用也都出现在链中。",
        "构造「交易码 → 主程序 → 两级子程序」的样例，交易说明的调用链包含三级程序；删除中间子程序的"
        "调用语句后重新生成，链中不再出现该子程序。",
        "交易说明中列出的表与链上各程序的 READ/WRITE 语句一致；以「说明中出现某表、但全链无该表的"
        "读写语句」为反例验收。",
        "资源定义不可得时生成的交易说明，其交易码与入口程序的对应关系标注来源为程序内的交易码分支，"
        "并标注「资源定义不可得」；以「直接采用人工整理的交易清单作为绑定依据」为反例验收。",
    ],
    "FR-STO-30": [
        "构造跨 4 个程序的链路（改动入口程序的某字段会影响下游 3 个程序），提问「改这个程序影响谁」，"
        "应答覆盖全部 4 个程序；只给出其中 2~3 个的判为不通过。",
        "链路中某对象在检索库中没有切片时，应答显式列出该对象并标注缺切片；以「省略该对象后按剩余部分"
        "作答、不作任何标注」为反例验收。",
        "同一问题连续提问两次，覆盖的对象集合相同；以「两次覆盖的对象不相同」为反例验收。",
    ],
    "FR-QA-23": [
        "命中某程序源码切片时，应答中的解释取自该程序已生成的知识文档对应节并给出节号；"
        "以「应答中出现知识文档中不存在的新表述」为反例验收。",
        "命中段落级源码切片时，解释取自所属程序的流程或业务规则节，不返回其他程序的文档。",
        "命中字段或表时解释取自数据实体说明，命中交易时取自交易说明；以「命中字段却返回程序文档」"
        "为反例验收。",
        "该对象尚未成文时，应答只给源码原文并标注「该对象未成文」；以「未成文对象的应答中出现"
        "模型补写的解释段」为反例验收。",
    ],
}

# 需求编号 -> (标题, 章节, 功能组, 优先级, 插在哪个需求之后)
META = {
    "FR-GEN-34": ("交易与交易链路说明", "6 知识生成", "6.3.3 对象产物", "必须", "FR-GEN-12"),
    "FR-STO-30": ("关系链路问题的检索覆盖", "8 入库建图", "8.3.3 向量化与索引", "必须", "FR-STO-21"),
    "FR-QA-23": ("命中源码时的知识取用", "10 问答应答", "10.3.4 来源约束", "必须", "FR-QA-08"),
}


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

    # 读原表：插入点与总行数
    wb = openpyxl.load_workbook(DST)
    ws = wb["验收清单"]
    last = ws.max_row
    rows_by = {}
    for r in range(5, last + 1):
        rows_by.setdefault(ws.cell(r, 2).value, []).append(r)
    n_old = last - 4
    wb.close()

    # 从后往前确定插入点（行号大的先插，避免漂移）
    plan = []
    for code, (_, _, _, _, after) in META.items():
        plan.append((max(rows_by[after]) + 1, code))
    # 升序算出插入完成后的最终起始行（每块加上它之前所有块的插入行数）
    final, off = {}, 0
    for at, code in sorted(plan):
        final[code] = at + off
        off += len(AC[code])
    plan.sort(reverse=True)
    print("插入计划（行号大的先执行）:", [(p, c) for p, c in plan])
    print("最终起始行:", final)

    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    book = xl.Workbooks.Open(DST)
    try:
        sh = book.Worksheets("验收清单")
        for at, code in plan:
            n = len(AC[code])
            sh.Rows("%d:%d" % (at, at + n - 1)).Insert(-4121)
            print("  在 %d 行插入 %d 行 -> %s" % (at, n, code))

        # 重排序号与 AC 编号，并按需求回填描述（插入行的 G 列随样式为空）
        total = n_old + sum(len(v) for v in AC.values())
        seq = []
        for r in range(5, 5 + total):
            seq.append((r - 4, "AC-%d" % (r - 4)))
        # 只写 A 列序号与 H 列 AC 编号；B 列是需求编号，不能被覆盖
        sh.Range("A5:A%d" % (5 + total - 1)).Value2 = [[x[0]] for x in seq]
        sh.Range("H5:H%d" % (5 + total - 1)).Value2 = [[x[1]] for x in seq]

        # 回填新行的六列内容（按最终行号定位，插入行是空白行、无需求编号）
        for code, (title, ch, grp, prio, _) in META.items():
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

        book.Save()
    finally:
        book.Close(False)
        xl.Quit()
    print("saved", DST)


if __name__ == "__main__":
    main()
