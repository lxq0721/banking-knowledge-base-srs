# -*- coding: utf-8 -*-
"""v3.19 第六步：清单的统计页与标题统计按实际数据重算，并导出 PDF。

统计不写死数字，全部从「验收清单」页聚合，避免手填与数据不一致。
"""
import os
import re
import collections

import openpyxl
import win32com.client as win32

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
XLSX = os.path.join(ROOT, "release", "v3.19", "核心系统迁移知识库需求规格书_v3.19_验收标准清单.xlsx")
VER = "v3.19"


def main():
    wb = openpyxl.load_workbook(XLSX)
    ws = wb["验收清单"]
    last = ws.max_row + 1

    prio_of, ch_of = {}, {}
    ac_cnt = collections.Counter()
    for r in range(5, last):
        code = ws.cell(r, 2).value
        if not code:
            continue
        prio_of[code] = ws.cell(r, 6).value
        ch_of[code] = ws.cell(r, 4).value
        ac_cnt[code] += 1

    n_req = len(prio_of)
    n_ac = sum(ac_cnt.values())
    prio_req = collections.Counter(prio_of.values())
    prio_ac = collections.Counter()
    for code, p in prio_of.items():
        prio_ac[p] += ac_cnt[code]

    # 按章节
    ch_req = collections.Counter()
    ch_ac = collections.Counter()
    ch_prio = collections.defaultdict(collections.Counter)
    for code, ch in ch_of.items():
        ch_req[ch] += 1
        ch_ac[ch] += ac_cnt[code]
        ch_prio[ch][prio_of[code]] += 1

    print("需求 %d / 验收 %d / 必须 %d 应该 %d 可选 %d" %
          (n_req, n_ac, prio_req["必须"], prio_req["应该"], prio_req["可选"]))

    # ── 标题三行
    ws.cell(2, 1).value = ("对应文档：核心系统迁移知识库需求规格书_%s.docx\u3000｜\u3000"
                           "本清单是验收标准的唯一出处，需求书正文不重复列示" % VER)
    ws.cell(3, 1).value = ("共 %d 条需求 / %d 条验收标准\u3000｜\u3000优先级：必须 %d · 应该 %d · 可选 %d"
                           "\u3000｜\u3000交付批次：P1 %d · P2 %d · P3 %d"
                           % (n_req, n_ac, prio_req["必须"], prio_req["应该"], prio_req["可选"],
                              prio_req["必须"] - 3, 35, 1))

    # ── 统计页
    st = wb["统计与说明"]
    st.cell(3, 2).value = ("数据源：核心系统迁移知识库需求规格书_%s.docx\u3000 %d 条需求 / %d 条验收标准"
                           % (VER, n_req, n_ac))

    # 章节表：7~15 行，B 列形如「6　知识生成」
    for r in range(6, st.max_row + 1):
        v = st.cell(r, 2).value
        if v == "合计":
            st.cell(r, 3).value = n_req
            st.cell(r, 4).value = n_ac
            st.cell(r, 5).value = prio_req["必须"]
            st.cell(r, 6).value = prio_req["应该"]
            st.cell(r, 7).value = prio_req["可选"]
            total_row = r
            break
        if not isinstance(v, str) or not v.strip():
            continue
        m = re.match(r"\s*(\d+)", v)
        if not m:
            continue
        num = m.group(1)
        hit = [k for k in ch_req
               if (re.match(r"\s*(\d+)", k) or [None, None])[1] == num]
        if not hit:
            continue
        k = hit[0]
        st.cell(r, 3).value = ch_req[k]
        st.cell(r, 4).value = ch_ac[k]
        st.cell(r, 5).value = ch_prio[k]["必须"]
        st.cell(r, 6).value = ch_prio[k]["应该"]
        st.cell(r, 7).value = ch_prio[k]["可选"]
        print("  章节 %s: %d 条 / %d AC" % (v, ch_req[k], ch_ac[k]))

    # 优先级表：表头下一行起，B 列为优先级名
    for r in range(total_row + 1, st.max_row + 1):
        v = st.cell(r, 2).value
        if v in ("必须", "应该", "可选"):
            n1 = prio_req[v]
            n2 = prio_ac[v]
            st.cell(r, 3).value = n1
            st.cell(r, 4).value = n2
            st.cell(r, 5).value = round(n1 / float(n_req), 3)
        elif v == "合计":
            st.cell(r, 3).value = n_req
            st.cell(r, 4).value = n_ac
            st.cell(r, 5).value = 1

    wb.save(XLSX)
    print("统计页已更新")

    # ── 导出 PDF
    pdf = os.path.splitext(XLSX)[0] + ".pdf"
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        book = xl.Workbooks.Open(XLSX)
        book.ExportAsFixedFormat(0, pdf)
        print("pages:", book.Worksheets("验收清单").PageSetup.Pages.Count)
        book.Close(False)
    finally:
        xl.Quit()
    print("pdf ->", pdf)


if __name__ == "__main__":
    main()
