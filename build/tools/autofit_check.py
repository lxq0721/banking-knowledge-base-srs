# -*- coding: utf-8 -*-
"""清单行高实测：Excel AutoFit 后回读，与设定值比较；导出 PDF"""
import os
import sys
import openpyxl
import win32com.client as win32

XLSX = os.path.abspath(sys.argv[1])
PDF = os.path.abspath(sys.argv[2])

wb = openpyxl.load_workbook(XLSX)
ws = wb['验收清单']
last = ws.max_row
before = {r: ws.row_dimensions[r].height for r in range(5, last + 1)}

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
xl = excel.Workbooks.Open(XLSX)
s = xl.Worksheets('验收清单')
s.Rows(f'5:{last}').AutoFit()
after = {}
for r in range(5, last + 1):
    after[r] = round(s.Rows(r).RowHeight, 2)
xl.Save()
s.PageSetup.Orientation = 2      # landscape
s.PageSetup.Zoom = False
s.PageSetup.FitToPagesWide = 1
s.PageSetup.FitToPagesTall = False
xl.ExportAsFixedFormat(0, PDF)
pages = xl.Worksheets('验收清单').PageSetup.Pages.Count
xl.Close(SaveChanges=True)
excel.Quit()

short = [(r, before[r], after[r]) for r in after
         if before[r] is not None and after[r] > before[r] + 0.6]
print('行数', last - 4)
print('AutoFit 后高于设定值的行数', len(short))
for x in short[:15]:
    print('  行', x[0], '设定', x[1], '需要', x[2])
print('最小/最大行高', min(after.values()), max(after.values()))
print('PDF 页数', pages)
