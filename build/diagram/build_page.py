# -*- coding: utf-8 -*-
"""把五张图 + 三张对照表拼成一份可传阅的说明页（讨论稿，非需求书正文）。

页面只做穿针引线：图内已有的标题、副题、结论不再复述，每图只留一句「看哪里」。
"""
import io, os

D = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(D, "..", ".."))
OUT = os.path.join(ROOT, "prototype", "切片与召回示意图_v1.4.html")


def svg(fn):
    """SVG 不落盘、不入库：直接调用生成函数取字符串。"""
    import gen_recall
    return getattr(gen_recall, fn)()


CSS = """
*{box-sizing:border-box}
body{margin:0;background:#F5F6F8;font-family:'Microsoft YaHei','PingFang SC',sans-serif;color:#1F1F1F}
.page{max-width:1620px;margin:0 auto;background:#fff;padding:44px 48px 56px}
h1{font-size:27px;margin:0 0 10px;letter-spacing:.5px}
.sub{font-size:13.5px;color:#595959;margin:0 0 6px}
.meta{font-size:12.5px;color:#8C8C8C;margin:0 0 30px;line-height:1.7}
h2{font-size:17px;margin:40px 0 10px;padding-left:10px;border-left:4px solid #2F5597}
.fig{background:#fff;border:1px solid #E3E6EA;border-radius:6px;padding:10px;margin-bottom:8px;overflow:auto}
.fig svg{display:block;width:100%;height:auto}

/* 图后一句「看哪里」 */
.hint{font-size:13px;color:#595959;line-height:1.75;margin:0 0 4px}
.hint b{color:#2F5597;font-weight:600}

body p.lead{font-size:13.4px;line-height:1.95;color:#333;margin:0 0 8px}
.lead b{color:#1F1F1F}
ol.steps{margin:4px 0 10px;padding-left:24px;font-size:13.4px;line-height:1.95;color:#333}
ol.steps li{margin:2px 0}
ol.steps b{color:#1F1F1F}

table{width:100%;border-collapse:collapse;font-size:13px;margin:4px 0 6px}
th{background:#F2F4F7;font-weight:600;text-align:left;padding:9px 10px;border:1px solid #DCE0E5;color:#333}
td{padding:9px 10px;border:1px solid #E3E6EA;vertical-align:top;line-height:1.55}
tr:nth-child(even) td{background:#FAFBFC}
td.g{font-weight:600;color:#1E7B45;white-space:nowrap;text-align:center}
td.o{font-weight:600;color:#C55A11;white-space:nowrap;text-align:center}
td.r{font-weight:600;color:#C00000;white-space:nowrap;text-align:center}
td.n{font-weight:600;white-space:nowrap;color:#2F5597}
.cap{font-size:12.6px;color:#595959;margin:0 0 10px}
.nw{white-space:nowrap}
@media print{body{background:#fff}.page{max-width:none;padding:0}.fig{border:0;padding:0}}
"""

SCEN = [
    ("同一判定散落三处", "三处写法一致、互不引用，拿到的是三段残缺描述；FR-GEN-20 只管「取值不同」的冲突",
     "未落条目，见附录 A 表 A-3", "空着", "r"),
    ("联机交易无整链产物", "只有 CSD 的「交易码 → 程序」这一跳，往下逐程序翻第 7 节",
     "FR-GEN-34（v3.19）", "已落", "g"),
    ("链上关系不进切片", "五项元数据里没有「属于哪条链的哪一环」，检索只能按相似度取",
     "FR-STO-30（v3.19）", "已落", "g"),
    ("相似度 ≠ 完备", "漏掉的那段措辞本就不像问题，排序永远把它排在尾部",
     "同 FR-STO-30；方向另见 FR-STO-32", "已落", "g"),
    ("命中代码取不到知识", "只有「知识 → 代码」的溯源，反向取用无条目",
     "FR-QA-23（v3.19）", "已落", "g"),
    ("边挂在哪一级节点", "表 8-5 列了节点与关系，没规定一条边挂哪级，一段程序只能答到程序",
     "FR-STO-31（v3.20）", "已落", "g"),
    ("问题里有对象名却只按措辞找", "检索入口只有相似度，明写着 ACCTFILE 也可能命中别处",
     "FR-QA-24（v3.20）", "已落", "g"),
]

MAPS = [
    ("知识 → 代码（溯源）", "FR-GEN-17 结论带出处、FR-QA-07 引用可定位、8.1 三库互指、详情三栏联动",
     "够", "已有", "g"),
    ("代码 → 知识（取用）", "FR-QA-23：按切片元数据的对象编号取该对象已生成的知识节",
     "够（对象编号 + 定位）", "已有", "g"),
    ("程序级", "程序与其知识文档 1:1，十节全在", "够", "已有", "g"),
    ("段落级", "FR-QA-23 明确落到所属程序的流程与业务规则节；FR-STO-31 让段落成为边的端点",
     "够（段落属同一对象编号）", "已有", "g"),
    ("字段 · 表级", "字段并入数据实体说明、表按工作表切；FR-QA-23 按对象类型落点",
     "够（访问矩阵与字段骨架）", "已有", "g"),
    ("交易码级", "FR-GEN-34 的「交易与交易链路说明」成为落点，表 8-4 已并入切片口径",
     "够（资源定义的绑定）", "已有", "g"),
]

MEANS = [
    ("线 A　关系建模", "对象与边齐了，跨程序的问题才答得全",
     "FR-GEN-34、FR-STO-30、FR-STO-32", "能。评测集里 3 多跳那 29 条直接验"),
    ("线 B　边的粒度", "答到段落与行，而不是整个程序",
     "FR-STO-31（表 8-5 补作业步骤节点）", "能。构造同程序两段落连两文件，看边落哪一级"),
    ("线 C　入口与反查", "入口确定，解释不现场编",
     "FR-QA-24、FR-QA-23", "能。同一问题两次提问，入口与措辞都一致"),
    ("未落　同一判定归并", "三段残缺描述合成一份",
     "（无）规则可判部分可入 L1a，语义部分须人判", "半能。语义相同写法不同的不能自动合并"),
]


def table(head, rows, kinds=None, widths=None):
    """kinds 为 None 时，取每行末尾一个元素作 class。"""
    out = ["<table>"]
    if widths:
        out.append("<colgroup>" + "".join('<col style="width:%s">' % w for w in widths) + "</colgroup>")
    out.append("<thead><tr>")
    for x in head:
        out.append("<th>%s</th>" % x)
    out.append("</tr></thead><tbody>")
    for r in rows:
        cells = list(r) if kinds else list(r[:-1])
        if kinds:
            classes = kinds
        else:
            classes = [""] * (len(cells) - 1) + [r[-1]]
        out.append("<tr>")
        for i, c in enumerate(cells):
            cls = classes[i]
            out.append('<td%s>%s</td>' % (' class="%s"' % cls if cls else "", c))
        out.append("</tr>")
    out.append("</tbody></table>")
    return "".join(out)


h = []
h.append('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">')
h.append("<title>切片与召回：COBOL 的业务流会不会被切碎</title><style>%s</style></head><body>" % CSS)
h.append('<div class="page">')

h.append("<h1>切片与召回：COBOL 的业务流会不会被切碎</h1>")
h.append('<p class="sub">依据《核心系统迁移知识库需求规格书》v3.20　·　讨论稿，非需求书正文　·　'
         '图号独立编号，不占用需求书图序</p>')
h.append('<p class="meta">图 1 是总览，图 2~图 5 展开原有的四个漏点，图 6、图 7 展开 v3.20 新补的两处'
         '——「图按什么建」与「边挂在哪一级」。图 2、图 6、图 7 的程序名与行号为讨论用简化示例。</p>')

# ---------------- 图 1 主图 ----------------
h.append('<div class="fig">%s</div>' % svg("fig0"))
h.append('<p class="hint"><b>看哪里：</b>五个环节从左到右排成一条链，灰的环节不出问题；'
         '每个蓝环节下方的虚线框是这一环的漏点，卡片下沿标出该处现在由哪条需求承接——'
         '绿色是已落条目，灰色是仍空着的那一处。</p>')

# ---------------- 图 2 ----------------
h.append("<h2>知识文档：内容在源头就是散的</h2>")
h.append('<div class="fig">%s</div>' % svg("fig1"))
h.append('<p class="hint"><b>看哪里：</b>上半三个程序靠 CALL 与 READ/WRITE 串成一条链；'
         '下半这条链被切成八片、只召回三片，漏掉的 04、05 恰在中间。</p>')

# ---------------- 图 3 ----------------
h.append("<h2>知识文档：同样是链，一类写成了一份，一类只到第一跳</h2>")
h.append('<div class="fig">%s</div>' % svg("fig3"))
h.append('<p class="hint"><b>看哪里：</b>左右对照——左边作业说明一处成文、有出处可引；'
         '右边交易只到「交易码 → 程序」这一跳，往下要逐个程序翻第 7 节。</p>')

# ---------------- 图 4 ----------------
h.append("<h2>切片与召回：三层怎么切，一次问答怎么走</h2>")
h.append('<div class="fig">%s</div>' % svg("fig2"))
h.append('<p class="hint"><b>看哪里：</b>上半看三层各切到什么粒度、各解决什么；'
         '下半左右两栏是同一道题的两条走法，差别落在第 3 步和第 5 步。</p>')

# ---------------- 图 5 ----------------
h.append("<h2>检索与作答：命中的是代码时，知识取不回来</h2>")
h.append('<div class="fig">%s</div>' % svg("fig4"))
h.append('<p class="hint"><b>看哪里：</b>上行是已有的正向溯源，下行是缺的反向取用；'
         '中间两列是两侧对应的键。下沿四张卡说明映射不是 1:1，取用要按对象类型落点。</p>')

# ---------------- 图 6 ----------------
h.append("<h2>图谱：按对象建，不按文件建</h2>")
h.append('<div class="fig">%s</div>' % svg("fig5"))
h.append('<p class="hint"><b>看哪里：</b>左上是节点与边——同一程序下两个段落各连各的文件，'
         '不是一条程序级边连两个文件；左下是切片，靠「对象编号 + 定位」两步挂回节点。'
         '右侧四条是这张图长什么样的边界，最后一条最容易被忽略：<b>生成知识不产生节点</b>。</p>')

# ---------------- 图 7 ----------------
h.append("<h2>图谱：从一行代码到一条链，两个方向都要能走</h2>")
h.append('<div class="fig">%s</div>' % svg("fig6"))
h.append('<p class="hint"><b>看哪里：</b>上部是同一程序的两段代码连两个文件；中部是两个问句走的两条路；'
         '下部是<b>逆着走</b>——从被写的文件回到写它的那一段，再回到调用链上游。'
         '右下三个开关，就是这条路上缺的三处口径。</p>')

# ---------------- 表 1 ----------------
h.append("<h2>表 1　七处缺口：六处已落条目，一处仍空着</h2>")
h.append('<p class="cap">v3.19 补了三处，v3.20 又补了三处；剩下的一处不是漏写，是判据未定——'
         '语义相同、写法不同的两段算不算同一条规则，得人来判。</p>')
h.append(table(["缺口", "原状", "承接条目", "状态"], SCEN,
               widths=["20%", "42%", "24%", "14%"]))

# ---------------- 表 2 ----------------
h.append("<h2>表 2　映射方向与粒度：六行现在都有落点</h2>")
h.append('<p class="cap">这张表在 v3.18 时有三行标「缺规则」，v3.20 的 FR-QA-23 与 FR-STO-31 把它们补齐了。</p>')
h.append(table(["方向 / 粒度", "需求书现状", "数据是否够", "结论"], MAPS,
               widths=["18%", "46%", "22%", "14%"]))

# ---------------- 表 3 ----------------
h.append("<h2>表 3　三条线与一处未落：各自能验到什么程度</h2>")
h.append('<p class="cap">前三条互相依赖：线 A 把边建出来，线 B 决定边挂哪一级，线 C 决定从哪儿进、取哪一节。</p>')
h.append(table(["线", "解决什么", "落在哪些条目", "能否自验"], MEANS,
               ["n", "", "", ""], widths=["18%", "26%", "28%", "28%"]))

# ---------------- 合起来 ----------------
h.append("<h2>合起来看</h2>")
h.append('<p class="lead">原来的四个漏点串成一条因果：<b>内容在源头是散的</b>（图 2）→ '
         '<b>有些链干脆没有产物</b>（图 3）→ <b>切片装不下链的关系</b>，召回只按相似度取（图 4）→ '
         '<b>有产物也可能取不回来</b>（图 5）。表现出来就是答案少一段，而且不报错。</p>')
h.append('<p class="lead">v3.20 补的两处是另一类问题：它们不是某一环漏了，而是<b>整条链的取数方式没定</b>——'
         '图按什么建（图 6）、边挂在哪一级、从哪个方向走（图 7）。'
         '这类问题按字面实现也能过验收，测试照着条目写用例一条都不会挂，只有真拿一个跨程序的问题去问才暴露。</p>')
h.append('<p class="lead">所以顺序不是按严重程度排，是按依赖排：</p>')
h.append('<ol class="steps">'
         '<li><b>先把边建对</b>（线 A + 线 B）——FR-GEN-34 补交易链、FR-STO-31 把边挂到段落。'
         '边建错了，后面调方向、调闭包都是在错的图上使劲。</li>'
         '<li><b>再定怎么走</b>——FR-STO-32 给每类关系定可遍历方向，读写与调用支持反向；'
         'FR-STO-30 把候选集从「相似度前 k 条」换成闭包。</li>'
         '<li><b>最后定从哪儿进</b>——FR-QA-24 让对象名先锚定入口，FR-QA-23 让命中源码时取已有知识节'
         '而不是现场编。</li>'
         '</ol>')
h.append('<p class="lead">反过来先调切片粒度或 top-k，是在错的地方使劲：那两处调的是「找到没找到」，'
         '而这里的问题是<b>「找全没找全」</b>——漏掉的那一段，措辞本来就跟问题不像，把 k 调大带进来的是噪声。</p>')
h.append('<p class="lead">还有一件事图给不了：<b>顺序</b>。图只会说这几个对象相关，不会说先走哪个；'
         '先后只能来自交易说明的调用链、JCL 的 STEP 次序，或调用点的行号。'
         '这也是 FR-GEN-34 必须成文的原因——链路成文了，才有「一条链」这个可引用的东西。</p>')

h.append("</div></body></html>")

import re
_html = "\n".join(h)
# 只处理正文：先把 SVG 挖出占位，避免把 span 塞进 <text> 里破坏结构
_svgs = []


def _stash(m):
    _svgs.append(m.group(0))
    return "@@SVG%d@@" % (len(_svgs) - 1)


_html = re.sub(r"<svg[\s\S]*?</svg>", _stash, _html)
_html = re.sub(r"(FR-[A-Z]{3}-\d{2})", r'<span class="nw">\1</span>', _html)
for _i, _s in enumerate(_svgs):
    _html = _html.replace("@@SVG%d@@" % _i, _s, 1)
io.open(OUT, "w", encoding="utf-8", newline="\n").write(_html)
print("written", OUT, os.path.getsize(OUT))
