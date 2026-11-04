# -*- coding: utf-8 -*-
"""重写 Office 文件的 docProps/core.xml。

用途：交付前统一作者、标题、说明字段。不要用正则就地替换——openpyxl 写出的
core.xml 在元素上带局部 xmlns:dc 声明，正则整段替换会把声明一起吃掉，后续
dc:* 元素全部失名，openpyxl 回读直接报 Namespace prefix not defined；带属性
的元素（如 <dc:description xmlns:dc=...>）也匹配不到，会插入重复元素。

做法：读成字段 dict（键名一律不带命名空间前缀），根元素统一声明 cp/dc/
dcterms/dcmitype/xsi，按 ECMA-376 顺序整体重写，再回读验证。

用法：
    py -3 build/tools/fix_core.py <file> [--title T] [--description D]
                                  [--subject S] [--keywords K]
                                  [--creator C] [--last-by L] [--category G]
                                  [--created ISO] [--modified ISO]
"""
import argparse
import datetime
import os
import re
import shutil
import sys
import tempfile
import zipfile

NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

# 内部键名（不带前缀） → XML 元素名。顺序即 ECMA-376 规定顺序，不要打乱
ORDER = [("title", "dc:title"), ("subject", "dc:subject"), ("creator", "dc:creator"),
         ("keywords", "cp:keywords"), ("description", "dc:description"),
         ("lastModifiedBy", "cp:lastModifiedBy"), ("revision", "cp:revision"),
         ("lastPrinted", "cp:lastPrinted"), ("created", "dcterms:created"),
         ("modified", "dcterms:modified"), ("category", "cp:category"),
         ("contentStatus", "cp:contentStatus"), ("version", "cp:version")]

# 命令行参数名 → 内部键名
ARG = {"title": "title", "subject": "subject", "creator": "creator",
       "keywords": "keywords", "description": "description",
       "last_by": "lastModifiedBy", "revision": "revision",
       "last_printed": "lastPrinted", "category": "category",
       "created": "created", "modified": "modified"}


def read_core(zpath):
    z = zipfile.ZipFile(zpath)
    if "docProps/core.xml" not in z.namelist():
        return {}
    s = z.read("docProps/core.xml").decode("utf-8")
    out = {}
    for key, tag in ORDER:
        m = re.search(r"<%s\b[^>]*>(.*?)</%s>" % (tag, tag), s, re.S)
        if m:
            out[key] = m.group(1).strip()
    return out


def esc(v):
    return v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_core(f):
    head = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties'
            ' xmlns:cp="%s" xmlns:dc="%s" xmlns:dcterms="%s"'
            ' xmlns:dcmitype="%s" xmlns:xsi="%s">'
            % (NS["cp"], NS["dc"], NS["dcterms"], NS["dcmitype"], NS["xsi"]))
    body = []
    for key, tag in ORDER:
        v = f.get(key)
        if not v:
            continue
        if key in ("created", "modified"):
            body.append('<%s xsi:type="dcterms:W3CDTF">%s</%s>' % (tag, esc(v), tag))
        else:
            body.append("<%s>%s</%s>" % (tag, esc(str(v)), tag))
    return head + "".join(body) + "</cp:coreProperties>"


def rewrite(zpath, updates):
    """整个 zip 重打包，只换 docProps/core.xml。"""
    cur = read_core(zpath)
    for k, v in updates.items():
        if v is not None:
            cur[k] = v
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    cur.setdefault("created", now)
    if updates.get("modified"):
        cur["modified"] = updates["modified"]
    elif not cur.get("modified"):
        cur["modified"] = now

    xml = build_core(cur)
    d = os.path.dirname(os.path.abspath(zpath))
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(zpath)[1], dir=d)
    os.close(fd)
    try:
        with zipfile.ZipFile(zpath) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "docProps/core.xml":
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        shutil.move(tmp, zpath)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return cur


def verify(zpath):
    """回读验证：OOXML 能打开、字段值对得上。"""
    ext = os.path.splitext(zpath)[1].lower()
    if ext == ".xlsx":
        import openpyxl
        cp = openpyxl.load_workbook(zpath).properties
        return {"creator": cp.creator, "lastModifiedBy": cp.lastModifiedBy,
                "title": cp.title, "description": cp.description,
                "created": str(cp.created), "modified": str(cp.modified)}
    elif ext == ".docx":
        import docx
        cp = docx.Document(zpath).core_properties
        return {"creator": cp.author, "lastModifiedBy": cp.last_modified_by,
                "title": cp.title, "description": cp.comments,
                "created": str(cp.created), "modified": str(cp.modified)}
    return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    for k in ARG:
        ap.add_argument("--" + k.replace("_", "-"), dest=k)
    a = ap.parse_args()
    upd = {}
    for argname, field in ARG.items():
        v = getattr(a, argname, None)
        if v is not None:
            upd[field] = v
    if not upd:
        print("没有任何要改的字段")
        return 1
    for f in a.files:
        old = read_core(f)
        rewrite(f, upd)
        got = verify(f)
        print("==", os.path.basename(f))
        for k in ("creator", "lastModifiedBy", "title", "description", "created"):
            o, n = old.get(k, "-"), got.get(k, "-")
            print("   %-16s %s" % (k, ("%s  →  %s" % (o, n)) if o != n else o))
    return 0


if __name__ == "__main__":
    sys.exit(main())
