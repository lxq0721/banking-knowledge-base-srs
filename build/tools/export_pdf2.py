# -*- coding: utf-8 -*-
"""导出 PDF（带重试）：更新域与目录后导出"""
import os
import sys
import time
import win32com.client as win32

src = os.path.abspath(sys.argv[1])
dst = os.path.abspath(sys.argv[2])


def retry(fn, n=12, wait=3):
    last = None
    for i in range(n):
        try:
            return fn()
        except Exception as e:
            last = e
            time.sleep(wait)
    raise last


word = win32.DispatchEx('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = retry(lambda: word.Documents.Open(src, ReadOnly=False, AddToRecentFiles=False))
    retry(lambda: doc.Fields.Update())
    try:
        retry(lambda: doc.TablesOfContents(1).Update())
    except Exception:
        pass
    retry(lambda: doc.Repaginate())
    pages = retry(lambda: doc.ComputeStatistics(2))
    retry(lambda: doc.ExportAsFixedFormat(dst, 17))
    print('pages', pages)
    doc.Saved = True
    retry(lambda: doc.Close(SaveChanges=0))
finally:
    try:
        word.Quit()
    except Exception:
        pass
print('ok')
