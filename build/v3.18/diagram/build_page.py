# -*- coding: utf-8 -*-
"""把三张图 + 场景对照表 + 缺口处置表拼成一份可交付的 HTML（依据需求书 v3.18）。"""
import io, os

D = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(D, "..", ".."))
if not os.path.isdir(os.path.join(ROOT, "prototype")):
    ROOT = os.path.abspath(os.path.join(D, "..", "..", ".."))
OUT = os.path.join(ROOT, "prototype", "资料管理流程图_v1.2.html")


def svg(name):
    return io.open(os.path.join(D, name), encoding="utf-8").read()


SCEN = [
    ("S1", "首次整包；之后只换一个程序文件",
     "该文件路径相同、指纹不同 → 同路径变更",
     "新版本入库（留历史）→ 重解析该文件 → 重算：本文件、引用它的程序、四类汇总",
     "闭合"),
    ("S2", "整包重传，包里只有部分文件变了",
     "未变的 → 重复提交；变的 → 同路径变更；新的 → 新增",
     "只有变更与新增的往后走，未变的不重解析、不重算",
     "闭合，且不会全量重跑"),
    ("S3", "整包重传，其中有文件被删掉",
     "操作人声明整包范围后，与上一版比对得出「上次有、本次无」",
     "按单份停用处置：知识文档转「已废止」、切片下线、清单标「本次未提交」",
     "v3.18 闭合（表 6-9 第三情形）"),
    ("S4", "同一文件改名并改了内容、或换格式（doc→pdf）",
     "路径与指纹均未命中，但文件名主干相同 → 疑似同一份",
     "预检列候选，由人指定接续：接续按「同路径变更」，否则按「新增」",
     "v3.18 闭合（表 4-5 第四情形）"),
    ("S5", "A、B 两系统分别提交；之后只更新 B，且改到两者相接处",
     "B 侧判同路径变更；A 侧那份资料没有变化",
     "只重算 B 系统下引用它的程序；A 侧按应用系统边界不重算，改动侧登记疑点",
     "待确认（表 A-3）：边界待委托方定"),
    ("S6", "先传程序，隔一批再补 copybook",
     "后一批只有 copybook 是「新增」；程序不在本批范围内",
     "按 FR-GEN-32 重算上一批的程序——「生成单元」已定义，重算不限于本批范围",
     "v3.18 闭合（术语表「生成单元」）"),
    ("S7", "撤销批次后重新上传同一批资料",
     "撤销时该批资料的知识层已转「已废止」；重传判「新增」",
     "重新走全链；旧知识已下线，不会与新记录并存",
     "v3.18 闭合（FR-GEN-33）"),
    ("S8", "打包方式变了（多套一层目录 / 换成压缩包前缀）",
     "所有路径都不同 → 判「疑似同一份」或「新增」",
     "预检中逐份确认是否接续；确认后可批量接续，不必整系统重建",
     "v3.18 闭合（表 4-5）"),
]

CSS = """
*{box-sizing:border-box}
body{margin:0;background:#F5F6F8;font-family:'Microsoft YaHei','PingFang SC',sans-serif;color:#1F1F1F}
.page{max-width:1620px;margin:0 auto;background:#fff;padding:44px 48px 60px}
h1{font-size:26px;margin:0 0 8px;letter-spacing:.5px}
.sub{font-size:13.5px;color:#595959;margin:0 0 4px}
.meta{font-size:12.5px;color:#8C8C8C;margin:0 0 28px}
h2{font-size:18px;margin:40px 0 6px;padding-left:10px;border-left:4px solid #2F5597}
.cap{font-size:12.6px;color:#595959;margin:0 0 14px}
.fig{background:#fff;border:1px solid #E3E6EA;border-radius:6px;padding:10px;margin-bottom:10px;overflow:auto}
.fig svg{display:block;width:100%;height:auto}
table{width:100%;border-collapse:collapse;font-size:13px;margin:8px 0 6px}
th{background:#F2F4F7;font-weight:600;text-align:left;padding:9px 10px;border:1px solid #DCE0E5;color:#333}
td{padding:9px 10px;border:1px solid #E3E6EA;vertical-align:top;line-height:1.55}
tr:nth-child(even) td{background:#FAFBFC}
td.n{font-weight:600;white-space:nowrap;color:#2F5597}
th:first-child{white-space:nowrap}
td.g{font-weight:600;color:#1E7B45;white-space:nowrap;text-align:center}
.tag{display:inline-block;font-size:11.5px;padding:1px 7px;border-radius:9px;border:1px solid #1E7B45;color:#1E7B45;margin-left:4px}
.note{font-size:13px;color:#4A4A4A;line-height:1.85;margin:12px 0 0}
.note b{color:#1F1F1F}
hr{border:0;border-top:1px solid #E3E6EA;margin:44px 0 0}
@media print{body{background:#fff}.page{max-width:none;padding:0}.fig{border:0;padding:0}}
"""


def table(head, rows, kinds):
    out = ["<table><thead><tr>"]
    for h in head:
        out.append("<th>%s</th>" % h)
    out.append("</tr></thead><tbody>")
    for r in rows:
        out.append("<tr>")
        for i, c in enumerate(r):
            cls = kinds[i]
            out.append('<td%s>%s</td>' % (' class="%s"' % cls if cls else "", c))
        out.append("</tr>")
    out.append("</tbody></table>")
    return "".join(out)


html = []
html.append('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">')
html.append("<title>资料管理链路 · 流程图与场景说明</title><style>%s</style></head><body>" % CSS)
html.append('<div class="page">')
html.append("<h1>资料管理链路：流程图与场景说明</h1>")
html.append('<p class="sub">依据《核心系统迁移知识库需求规格书》v3.18（138 条需求 / 485 条 AC）　·　'
            '本页 v1.2，图 3-3 层序已按链路先后重排、原「缺口处置表」并为一节附注</p>')
html.append('<p class="meta">三张图分别回答：链路怎么走（图 3-1）、一次提交怎么判、重算到哪（图 3-2）、'
            '撤销清到哪一层（图 3-3）。绿框为 v3.18 补入的处置规则；红虚线仅剩跨应用系统一处，挂在表 A-3 待确认。</p>')

html.append('<div class="fig">%s</div>' % svg("_fig1.svg"))
html.append('<p class="cap">图 3-1　主干八步，右侧为资料再次提交、撤销与失效的四条路径；逐文件判定展开见 图 3-2。</p>')

html.append('<div class="fig">%s</div>' % svg("_fig3.svg"))
html.append('<p class="cap">图 3-2　判定单位是文件，不是批次——同一套出口同时覆盖单份替换、整包重传、多系统分别提交三种形态。</p>')

html.append('<div class="fig">%s</div>' % svg("_fig2.svg"))
html.append('<p class="cap">图 3-3　四层按链路先后排列：原件 → 解析产物 → 知识产物 → 检索索引。撤销清掉了知识产物的上游与下游，唯独它本身原先不失效。</p>')

html.append("<h2>表 1　八种提交形态与系统实际处理</h2>")
html.append('<p class="cap">前两行是最常见的两类，后面按同一判据推出；「结论」一列标出在 v3.18 下是否闭环。</p>')
html.append(table(["编号", "提交形态", "系统怎么判", "会做什么", "结论"],
                  SCEN, ["n", "", "", "", "g"]))

html.append('<p class="note"><b>读图要点：</b>图 3-1 的第 02 步仍是唯一的分岔点，'
            '后面所有复杂度都来自它——一份资料进来先被归到四种情形之一，只有「变更」「新增」与经人接续的'
            '「疑似同一份」往后走。图 3-2 把这四种情形与两种失效触发摊开成一个逐文件判定表；'
            '图 3-3 说明为什么撤批之后问答不能继续拿到结论：知识产物删掉无法复现，只能另出新版本、旧版留档，'
            '所以它在撤销时必须显式失效。</p>')

html.append("<h2>附　本版改动与审查留痕</h2>")
html.append('<p class="note"><b>①②③⑥ · 失效与撤销：</b>撤销批清的是导入记录、解析产物与索引条目，'
            '知识文档与术语表不在清除范围（①）；只有整批撤销、没有单份撤销（②）；'
            '表 7-4 的「已废止」由源文件变更进入，却没写谁触发（③）；'
            '整包重传时「上次有、本次无」不落任何情形（⑥）。四条由 FR-GEN-33 与表 6-9 一并收口。</p>')
html.append('<p class="note"><b>④⑤⑦⑧ · 判定与重算：</b>判据只认路径与内容指纹、不认文档标识，'
            '改名换格式都判成新增（④），由表 4-5 第四情形接住；汇总产物只列三类、漏了项目知识总览（⑤），'
            '由 FR-GEN-32 补入；「生成单元」未定义导致重算范围说不清（⑧），术语表已补。'
            '⑦ 跨应用系统的引用方是否重算，取决于委托方系统边界如何划分，不写成条目，挂表 A-3，确认前按应用系统边界处理。</p>')
html.append('<p class="note"><b>改动量：</b>新增 1 条需求（FR-GEN-33，必须）与 1 张三情形表（表 6-9）；'
            '表 4-5 加一行、FR-GEN-32 加一项、表 7-4 改一处进入条件、术语表加「生成单元」、表 A-3 加一项。'
            '合计 137 → 138 条，AC 481 → 485。编号 ①～⑧ 与附录 C 的记录一一对应。</p>')

html.append("</div></body></html>")

io.open(OUT, "w", encoding="utf-8", newline="\n").write("\n".join(html))
print("written", OUT, os.path.getsize(OUT))
