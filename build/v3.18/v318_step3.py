# -*- coding: utf-8 -*-
"""v3.18 step3：Word 更新域与目录、导出 PDF、清文档属性、落到 release/"""
import os
import sys
import io
import time
import win32com.client as win32
import docx

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = os.path.abspath('_v318_step2.docx')
OUT_DIR = os.path.abspath('../../release/v3.18')
os.makedirs(OUT_DIR, exist_ok=True)
DST = os.path.join(OUT_DIR, '核心系统迁移知识库需求规格书_v3.18.docx')
PDF = os.path.join(OUT_DIR, '核心系统迁移知识库需求规格书_v3.18.pdf')


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
    doc = retry(lambda: word.Documents.Open(SRC, ReadOnly=False, AddToRecentFiles=False))
    retry(lambda: doc.Fields.Update())
    try:
        retry(lambda: doc.TablesOfContents(1).Update())
    except Exception as e:
        print('TOC 更新跳过:', e)
    retry(lambda: doc.Repaginate())
    pages = retry(lambda: doc.ComputeStatistics(2))
    retry(lambda: doc.Save())
    retry(lambda: doc.ExportAsFixedFormat(PDF, 17, False, 0, 0, 0, 0, 0, True, True, 1, True, True, False))
    retry(lambda: doc.Close(SaveChanges=0))
    print('更新域完成，页数', pages, '| PDF', PDF)
finally:
    try:
        word.Quit()
    except Exception:
        pass

d = docx.Document(SRC)
cp = d.core_properties
cp.title = '核心系统迁移知识库需求规格书'
cp.subject = ''
cp.comments = ''
cp.keywords = ''
cp.category = ''
cp.author = ''
cp.last_modified_by = ''
d.save(DST)
print('saved', DST)

d2 = docx.Document(DST)
print('封面版本', repr(d2.paragraphs[3].text))
print('属性', repr(d2.core_properties.title), repr(d2.core_properties.author))
print('大小', round(os.path.getsize(DST) / 1024 / 1024, 2), 'MB')
