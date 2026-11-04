# -*- coding: utf-8 -*-
"""切片与召回示意图：五张 SVG。

图 1 总览：一条链五个环节上的漏点。
图 2~图 5 是对其中四个漏点的放大：图 2、图 3 属「知识文档」环节，
图 4 属「切片 · 向量」与「检索 · 答案」，图 5 属「检索 · 答案」。
"""
import io, os

D = os.path.dirname(os.path.abspath(__file__))
F = "Microsoft YaHei,PingFang SC,Hiragino Sans GB,sans-serif"

INK = "#1F1F1F"
SUB = "#595959"
MUTE = "#8C8C8C"
BLUE = "#2F5597"
BLUE_BG = "#EDF3FA"
ORG = "#C55A11"
ORG_BG = "#FDF3E7"
GRN = "#1E7B45"
GRN_BG = "#EAF5EE"
RED = "#C00000"
RED_BG = "#FDECEC"
LINE = "#E3E6EA"


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def t(x, y, s, size=14, fill=INK, w="normal", anchor="start", ls="0", op=None):
    o = ' opacity="%s"' % op if op else ""
    return ('<text x="%.1f" y="%.1f" font-family="%s" font-size="%s" fill="%s" '
            'font-weight="%s" text-anchor="%s" letter-spacing="%s"%s>%s</text>'
            % (x, y, F, size, fill, w, anchor, ls, o, esc(s)))


def r(x, y, w, h, fill="#FFFFFF", stroke="none", sw=0, rx=6, dash=None):
    d = ' stroke-dasharray="%s"' % dash if dash else ""
    return ('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="%s" fill="%s" '
            'stroke="%s" stroke-width="%s"%s/>' % (x, y, w, h, rx, fill, stroke, sw, d))


def ln(x1, y1, x2, y2, stroke="#404040", sw=1.6, dash=None, marker=None):
    d = ' stroke-dasharray="%s"' % dash if dash else ""
    m = ' marker-end="url(#%s)"' % marker if marker else ""
    return ('<path d="M %.1f %.1f L %.1f %.1f" fill="none" stroke="%s" stroke-width="%s"%s%s/>'
            % (x1, y1, x2, y2, stroke, sw, d, m))


def poly(pts, stroke="#404040", sw=1.6, dash=None, marker=None, fill="none"):
    d = ' stroke-dasharray="%s"' % dash if dash else ""
    m = ' marker-end="url(#%s)"' % marker if marker else ""
    p = " ".join("%.1f %.1f" % (a, b) for a, b in pts)
    return ('<path d="M %s" fill="%s" stroke="%s" stroke-width="%s"%s%s/>'
            % (p, fill, stroke, sw, d, m))


def circ(cx, cy, rr, fill="#FFFFFF", stroke=BLUE, sw=1.8, dash=None):
    d = ' stroke-dasharray="%s"' % dash if dash else ""
    return ('<circle cx="%.1f" cy="%.1f" r="%.1f" fill="%s" stroke="%s" stroke-width="%s"%s/>'
            % (cx, cy, rr, fill, stroke, sw, d))


def tri(cx, cy, fill, w=11, h=7):
    return ('<path d="M %.1f %.1f L %.1f %.1f L %.1f %.1f z" fill="%s"/>'
            % (cx - w / 2.0, cy, cx + w / 2.0, cy, cx, cy + h, fill))


def defs(tag):
    out = ['<defs>']
    for name, color in (("g", "#404040"), ("b", BLUE), ("r", RED), ("o", ORG)):
        out.append('<marker id="m%s%s" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" '
                   'markerHeight="7" orient="auto-start-reverse">'
                   '<path d="M 0 0 L 10 5 L 0 10 z" fill="%s"/></marker>' % (tag, name, color))
    out.append('</defs>')
    return "".join(out)


def wrap(W, H, body):
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
            % (W, H, W, H) + r(0, 0, W, H, "#FFFFFF", rx=6) + body + '</svg>')


# ---------------------------------------------------------------- 图 1 总览
def fig0():
    """总览：一条链上的五个漏点，并标出各自由哪条需求承接。"""
    W, H = 1500, 692
    b = [defs("z")]
    b.append(t(64, 52, "图 1　一条链上的五个漏点：两处已落条目，一处仍空着", 23, INK, "bold"))
    b.append(t(64, 80, "从代码到答案是一条链，每一环都可能漏，且漏法互不相同——修上游那一处，补不了下游那一处。"
                       "卡片下沿标出该处现在由哪条需求承接", 13.6, SUB))

    stage = [("代码", "原文", "无漏", "#7F7F7F", "#F7F7F7"),
             ("结构事实", "解析产物", "无漏", "#7F7F7F", "#F7F7F7"),
             ("知识文档", "成文产物", "2 处漏点", BLUE, BLUE_BG),
             ("切片 · 向量", "检索单元", "1 处漏点", BLUE, BLUE_BG),
             ("检索 · 答案", "召回与作答", "2 处漏点", BLUE, BLUE_BG)]
    bw, gap, by, bh = 252, 30, 112, 64
    x0 = (W - (bw * 5 + gap * 4)) / 2.0
    for i, (name, sub, flag, col, bg) in enumerate(stage):
        x = x0 + i * (bw + gap)
        b.append(r(x, by, bw, bh, bg, col, 1.5, rx=5))
        b.append(t(x + 18, by + 30, name, 15.5, INK, "bold"))
        b.append(t(x + 18, by + 51, sub, 12.2, SUB))
        b.append(t(x + bw - 16, by + 30, flag, 11.4, col, "bold", "end"))
        if i < 4:
            b.append(ln(x + bw + 3, by + bh / 2, x + bw + gap - 5, by + bh / 2, "#A6A6A6", 1.5, None, "mzg"))

    # 漏点卡片：最后一项是承接状态（文字, 颜色）
    cw, ch = 252, 110
    cards = {
        0: [("原文永远在", ["解析出错也只错在它的下游"], "#7F7F7F", "#F7F7F7", None, None)],
        1: [("按文件分工抽取", ["程序出调用、copybook 出字段", "JCL 出作业链"], "#7F7F7F", "#F7F7F7", None, None)],
        2: [("同一判定散落三处", ["A 算出、B 判它、C 另有一处校验", "三处彼此不引用"], RED, RED_BG, "图 2",
             ("未落条目 · 见附录 A 表 A-3", MUTE)),
            ("联机交易无整链产物", ["CSD 只给「交易码 → 程序」这一跳", "再往下没有成文产物"], ORG, ORG_BG, "图 3",
             ("v3.19 已补 · FR-GEN-34", GRN))],
        3: [("链上关系不进切片", ["五项元数据里没有", "「属于哪条链的哪一环」"], RED, RED_BG, "图 4",
             ("v3.19 已补 · FR-STO-30、32", GRN))],
        4: [("相似度 ≠ 完备", ["漏的那段措辞本来就不像问题", "调大 k 带进来的是噪声"], RED, RED_BG, "图 4",
             ("与上一处同条 · 不单独立项", GRN)),
            ("命中代码取不到知识", ["只有「知识 → 代码」的溯源", "没有反向取用的规则"], ORG, ORG_BG, "图 5",
             ("v3.19 已补 · FR-QA-23", GRN))],
    }
    for i, lst in cards.items():
        x = x0 + i * (bw + gap)
        for k, (head, lines, col, bg, ref, status) in enumerate(lst):
            y = 206 + k * (ch + 12)
            b.append(r(x, y, cw, ch, bg, col, 1.4, rx=5, dash="5 3" if col != "#7F7F7F" else None))
            b.append(t(x + 16, y + 26, head, 13.6, col if col != "#7F7F7F" else SUB, "bold"))
            for m, ln_ in enumerate(lines[:2]):
                b.append(t(x + 16, y + 49 + m * 18, ln_, 11.6, SUB))
            if ref:
                b.append(r(x + cw - 62, y + 8, 46, 18, "#FFFFFF", col, 1.0, rx=3))
                b.append(t(x + cw - 39, y + 21, ref, 10.8, col, "bold", "middle"))
            if status:
                b.append(ln(x + 16, y + ch - 26, x + cw - 16, y + ch - 26, LINE, 1))
                b.append(t(x + 16, y + ch - 9, status[0], 10.6, status[1], "bold"))

    # 结论
    b.append(ln(64, 468, 1436, 468, LINE, 1))
    b.append(t(64, 498, "五处漏点里两处已落条目、一处仍空着；而且「调一个参数」从来不是选项——", 14, INK, "bold"))
    b.append(t(64, 524, "成文环节的漏要补产物，切片与召回环节的漏要改取数方式，反查环节的漏要加取用规则。三者互不替代。",
               13, "#404040"))
    b.append(t(64, 550, "但漏 2、3、4 共用的是同一套「对象 + 关系」建模：先有对象与边，才谈得上按闭包取数、按链路成文。",
               13, "#404040"))
    b.append(t(64, 584, "另两处不在上面五处之列——「边挂在哪一级节点」与「问题里有对象名却只按措辞找」。它们不是某一环漏了，",
               13, "#404040"))
    b.append(t(64, 606, "而是整条链的取数方式没定，v3.20 分别由 FR-STO-31 与 FR-QA-24 承接，见新图 6、图 7。",
               13, "#404040"))
    b.append(t(64, 652, "注：本图为整体结构示意；各环节的具体机制见后续各图。节点名与产物名取自需求书，编号见图内标注。",
               12.2, MUTE))
    return wrap(W, H, "".join(b))


# ---------------------------------------------------------------- 图 5 双向映射
def fig4():
    """知识与代码的双向映射：两个方向都已有条目，落点按对象类型不同。"""
    W, H = 1500, 704
    b = [defs("e")]
    b.append(t(64, 52, "图 5　检索与作答：两个方向都通了，只是落点不同", 23, INK, "bold"))
    b.append(t(64, 80, "两侧本就一一对应——知识文档由代码生成，映射键就写在切片的元数据里。"
                       "v3.19 的 FR-QA-23 补上了反向这条：命中代码后按对象编号取该对象已生成的知识节", 13.6, SUB))

    # 上：正向（已有）
    b.append(ln(120, 176, 1362, 176, GRN, 1.8, None, "eG"))
    b.append(r(560, 164, 400, 24, "#FFFFFF", rx=3))
    b.append(t(760, 181, "① 正向：结论 → 出处 → 源码行号", 13, GRN, "bold", "middle"))
    b.append(t(1436, 181, "已有", 12.4, GRN, "bold", "end"))

    # 左栏：知识文档
    lx, lw, ty, th = 64, 430, 206, 312
    b.append(r(lx, ty, lw, th, "#F7FBF8", GRN, 1.5, rx=5))
    b.append(t(lx + 22, ty + 30, "知识文档 PGMAA", 15.5, INK, "bold"))
    b.append(t(lx + 22, ty + 52, "十节结构，一节即一个切片", 12.2, SUB))
    secs = [("§4 处理流程", "Mermaid，节点取自段落名", "索引"),
            ("§5 业务规则", "含出处与置信度（表 6-3）", "解释"),
            ("§7 依赖关系", "调用、被调用、共享数据", "索引"),
            ("§10 知识来源", "生成方式、核验、版本", "必填")]
    sy = ty + 74
    for name, note, ch_ in secs:
        b.append(r(lx + 22, sy, lw - 44, 44, "#FFFFFF", "#CFE3D6", 1.2, rx=4))
        b.append(t(lx + 38, sy + 20, name, 13.2, INK, "bold"))
        b.append(t(lx + 38, sy + 37, note, 11.4, SUB))
        b.append(t(lx + lw - 38, sy + 20, ch_, 10.6, GRN, "normal", "end"))
        sy += 52
    b.append(t(lx + 22, ty + th - 14, "切片单元＝节，定位＝节号与段落序号（表 8-4）", 11.6, MUTE))

    # 中栏：映射键
    mx, mw = 534, 290
    b.append(t(mx + mw / 2, ty + 26, "对上两侧的，只有这两项", 12.6, INK, "bold", "middle"))
    b.append(r(mx, ty + 44, mw, 214, "#FFFFFF", BLUE, 1.5, rx=5))
    b.append(t(mx + 20, ty + 74, "对象编号", 12.6, BLUE, "bold"))
    b.append(t(mx + 20, ty + 96, "PGMAA（程序）", 12.6, INK))
    b.append(t(mx + 20, ty + 116, "3000-CALC（段落）", 12.6, INK))
    b.append(ln(mx + 20, ty + 132, mx + mw - 20, ty + 132, LINE, 1))
    b.append(t(mx + 20, ty + 158, "定位信息", 12.6, BLUE, "bold"))
    b.append(t(mx + 20, ty + 180, "§4 · 段落序号 2", 12.6, INK))
    b.append(t(mx + 20, ty + 200, "行号区间 1180–1246", 12.6, INK))
    b.append(t(mx + 20, ty + 228, "五项元数据里其余三项（来源类型、级别、版本）", 11.4, MUTE))
    b.append(t(mx + 20, ty + 246, "不承担定位，无法用来对应两侧对象", 11.4, MUTE))

    # 右栏：源码
    rx2, rw2 = 864, 572
    b.append(r(rx2, ty, rw2, 300, "#FBF7F4", ORG, 1.5, rx=5))
    b.append(t(rx2 + 22, ty + 30, "源码与结构事实", 15.5, INK, "bold"))
    b.append(t(rx2 + 22, ty + 52, "段落按段切、字段并入数据实体说明、表按工作表切", 12.2, SUB))
    items = [("PGMAA · 3000-CALC", "行号 1180–1246　·　段落切片", "段落", "#FFFFFF", "#D9C2AE"),
             ("ACCTFILE · FLAG-POSTPONE", "copybook 01 级记录　·　并入数据实体说明", "字段", "#FFFFFF", "#D9C2AE"),
             ("ACCTFILE", "DB2 表 / 物理数据集　·　按工作表切片", "表", "#FFFFFF", "#D9C2AE")]
    sy = ty + 74
    for name, note, kind, bg, stroke in items:
        b.append(r(rx2 + 22, sy, rw2 - 44, 52, bg, stroke, 1.2, rx=4))
        b.append(t(rx2 + 38, sy + 22, name, 13.2, INK, "bold"))
        b.append(t(rx2 + 38, sy + 41, note, 11.4, SUB))
        b.append(t(rx2 + rw2 - 38, sy + 22, kind, 10.8, ORG, "normal", "end"))
        sy += 60
    b.append(t(rx2 + 22, ty + 286, "同一程序下所有段落共享一个对象编号 PGMAA，这是多对一的来源", 11.6, MUTE))

    # 下：反向（缺）
    b.append(ln(1362, 528, 120, 528, BLUE, 1.8, None, "eb"))
    b.append(r(556, 516, 412, 24, "#FFFFFF", rx=3))
    b.append(t(762, 533, "② 反向：命中源码切片 → 该对象的哪一节", 13, BLUE, "bold", "middle"))
    b.append(t(1436, 533, "v3.19 已补 · FR-QA-23", 12.4, GRN, "bold", "end"))

    # 映射粒度
    b.append(t(64, 574, "映射不是 1:1，四种对应关系各不相同：", 13.2, INK, "bold"))
    gr = [("程序", "→ 该程序的知识文档", "1:1，十节全在", GRN),
          ("段落", "→ 所属程序的 §4 / §5", "多对一，段落无独立文档", ORG),
          ("字段 · 表", "→ 数据实体说明", "跨对象，不在本程序文档里", ORG),
          ("交易码", "→ 交易链路说明", "已有，FR-GEN-34", GRN)]
    gw, ggap = 332, 20
    gx0 = (W - (gw * 4 + ggap * 3)) / 2.0
    for i, (k1, k2, k3, col) in enumerate(gr):
        x = gx0 + i * (gw + ggap)
        b.append(r(x, 588, gw, 60, "#FFFFFF", col, 1.3, rx=4))
        b.append(t(x + 16, 610, k1, 12.8, col, "bold"))
        b.append(t(x + 16 + len(k1) * 13, 610, k2, 12.2, INK))
        b.append(t(x + 16, 634, k3, 11.4, SUB))
    return wrap(W, H, "".join(b))


# ---------------------------------------------------------------- 图 2
def fig1():
    W, H = 1500, 852
    b = [defs("a")]
    b.append(t(64, 52, "图 2　知识文档：一段业务流的三个落点，被切成三片", 23, INK, "bold"))
    b.append(t(64, 80, "上为代码里真实的样子，下为切片进检索库后的样子。示意结构，程序名与字段为讨论用简化命名", 13.6, SUB))

    # 图例
    lx = 64
    for color, txt in ((BLUE, "召回命中"), (RED, "答案必需但未召回"), (MUTE, "未召回且与本问无关")):
        b.append(r(lx, 100, 22, 16, color, "none", 0, rx=3))
        b.append(t(lx + 32, 113, txt, 12.4, SUB))
        lx += 32 + len(txt) * 12.4 + 26
    b.append(r(lx, 100, 22, 16, ORG, "none", 0, rx=3))
    b.append(t(lx + 32, 113, "同一条判定的实现位置", 12.4, SUB))
    lx += 32 + 10 * 12.4 + 34
    b.append(ln(lx, 108, lx + 46, 108, "#BFBFBF", 1.4, "4 3"))
    b.append(t(lx + 54, 113, "原调用关系（切片层不保存）", 12.4, SUB))

    # 上栏
    b.append(t(64, 164, "① 逻辑视图　一段业务流的真实形态", 16.5, INK, "bold"))
    b.append(t(64, 188, "判定分散在 3 个程序里，靠 CALL 与 READ/WRITE 串成一条链；链是业务流，不是文本相似度", 13.2, SUB))

    ny, nh = 206, 158
    prog = [
        (140, 250, "PGMAA", "入口程序", "交易入口，算出压单标志",
         [("1000-INIT", 0), ("2000-READ", 0), ("3000-CALC", "1")]),
        (470, 250, "PGMBB", "子程序", "冲正与压单判定",
         [("2100-CHECK", "2"), ("2200-POST", 0)]),
        (800, 250, "PGMCC", "子程序", "更新与校验",
         [("1500-VALID", "3"), ("1600-UPDATE", 0)]),
    ]
    for x, w, name, kind, desc, paras in prog:
        b.append(r(x, ny, w, nh, "#FFFFFF", "#BFBFBF", 1.5))
        b.append(t(x + 18, ny + 30, name, 15.5, INK, "bold"))
        b.append(t(x + 94, ny + 30, kind, 12.2, MUTE))
        b.append(t(x + 18, ny + 52, desc, 12.4, SUB))
        b.append(ln(x + 16, ny + 62, x + w - 16, ny + 62, "#EAEAEA", 1.0))
        for k, (pname, mark) in enumerate(paras):
            py = ny + 86 + k * 23
            hot = bool(mark)
            b.append(circ(x + 22, py - 4, 3.4, ORG if hot else "#BFBFBF", "none", 0))
            b.append(t(x + 34, py, pname, 13.4, ORG if hot else INK, "bold" if hot else "normal"))
            if hot:
                b.append(circ(x + w - 26, py - 4, 9, ORG, "none", 0))
                b.append(t(x + w - 26, py + 0.5, mark, 11.5, "#FFFFFF", "bold", "middle"))

    # 表
    dx, dw = 1180, 200
    b.append(r(dx, ny, dw, nh, "#FFFFFF", "#BFBFBF", 1.5))
    b.append(t(dx + 18, ny + 30, "ACCTFILE", 15.5, INK, "bold"))
    b.append(t(dx + 18, ny + 52, "账户主文件", 12.4, SUB))
    b.append(ln(dx + 16, ny + 62, dx + dw - 16, ny + 62, "#EAEAEA", 1.0))
    b.append(circ(dx + 22, ny + 82, 3.4, "#BFBFBF", "none", 0))
    b.append(t(dx + 34, ny + 86, "字段定义", 13.4, INK))
    b.append(circ(dx + 22, ny + 105, 3.4, "#BFBFBF", "none", 0))
    b.append(t(dx + 34, ny + 109, "FLAG-POSTPONE", 13.4, INK))

    cy = ny + 86
    for x1, x2, lab in ((390, 468, "CALL"), (720, 798, "CALL"), (1050, 1178, "READ / WRITE")):
        mx = (x1 + x2) / 2.0
        b.append(ln(x1, cy, x2 - 2, cy, "#404040", 1.6, None, "mag"))
        hw = len(lab) * 3.5 + 6
        b.append(r(mx - hw, cy - 24, hw * 2, 18, "#FFFFFF", rx=3))
        b.append(t(mx, cy - 10.5, lab, 11.6, SUB, "normal", "middle"))

    b.append(t(64, 392, "①→②→③ 是同一条判定：A 算出标志、B 判该标志、C 另有一处校验。三处不在一份文件里，也不在一个段落里", 12.8, ORG))

    # 下栏
    b.append(t(64, 436, "② 切片视图　进检索库后的样子", 16.5, INK, "bold"))
    b.append(t(64, 460, "切片只带五项元数据（来源类型 / 级别 / 对象编号 / 定位 / 版本）；段落之间的调用与读写关系不进切片，也不进向量", 13.2, SUB))

    cells = [
        ("切片 01", "1000-INIT", "PGMAA · §1000", "未召回：与本问无关", MUTE, "#FFFFFF", None),
        ("切片 02", "2000-READ", "PGMAA · §2000", "未召回：与本问无关", MUTE, "#FFFFFF", None),
        ("切片 03", "3000-CALC", "PGMAA · §3000", "命中：措辞与问题相近", BLUE, BLUE_BG, None),
        ("切片 04", "2100-CHECK", "PGMBB · §2100", "漏：措辞与问题不像", RED, RED_BG, "4 3"),
        ("切片 05", "2200-POST", "PGMBB · §2200", "漏：措辞与问题不像", RED, RED_BG, "4 3"),
        ("切片 06", "1500-VALID", "PGMCC · §1500", "命中：措辞与问题相近", BLUE, BLUE_BG, None),
        ("切片 07", "1600-UPDATE", "PGMCC · §1600", "未召回：与本问无关", MUTE, "#FFFFFF", None),
        ("切片 08", "字段定义", "ACCTFILE · 工作表 1", "命中：表结构", BLUE, BLUE_BG, None),
    ]
    gw, gh = 280, 110
    for i, (no, name, loc, stat, color, bg, dash) in enumerate(cells):
        gx = 140 + (i % 4) * 320
        gy = 484 + (i // 4) * 136
        b.append(r(gx, gy, gw, gh, bg, color, 1.5, dash=dash))
        b.append(t(gx + 16, gy + 24, no, 11.6, MUTE))
        b.append(t(gx + 16, gy + 52, name, 15, color if color != MUTE else INK, "bold"))
        b.append(t(gx + 16, gy + 74, loc, 12.2, SUB))
        b.append(t(gx + 16, gy + 94, stat, 12.4, color))
        if color != MUTE:
            b.append(circ(gx + gw - 22, gy + 20, 5, color, "none", 0))

    # 结论
    b.append(r(64, 742, 1372, 70, "#FAFBFC", "#E3E6EA", 1, rx=5))
    b.append(t(84, 768, "为什么把 k 从 5 调到 20 也补不回来", 14, INK, "bold"))
    b.append(t(84, 792, "相似度回答的是「这段文字像不像问题」，不回答「这条链还缺哪一段」。缺的那一段措辞本来就不像问题，"
                        "排序永远在尾部；调大 k 带进来的是措辞相近的噪声，不是逻辑上必需的那一段。", 12.8, "#404040"))

    b.append(t(64, 838, "注：①②③ 只表示同一条判定分三处实现，不代表执行顺序；灰色虚线仅示逻辑关系存在，切片层并不保存这条关系。", 12.2, MUTE))
    return wrap(W, H, "".join(b))


# ---------------------------------------------------------------- 图 2
def fig2():
    W, H = 1500, 1096
    b = [defs("b")]
    b.append(t(64, 52, "图 4　切片与召回：三层分工，与两种召回路径", 23, INK, "bold"))
    b.append(t(64, 80, "层不同，切法就不同：结构层不切、成文层按节切、证据层按段切。问答题走第三条路径，不是把前两条叠加", 13.6, SUB))

    b.append(t(64, 128, "① 三层各管一段", 16.5, INK, "bold"))

    rows = [
        ("结构层", "图谱",
         ["程序 / 段落 / 字段 / 文件 / 表 / 术语", "边只取调用、读写、包含；不得由模型推断"],
         ("不切", "图只存编号与关系，不存正文"),
         ["该去哪儿找、影响面有多大", "事实类结论的出处由这里给定"],
         "FR-STO-26 · 表 8-2", BLUE, BLUE_BG),
        ("成文层", "知识文档",
         ["十节结构，一节一语义单元", "第 4 节的处理流程本身就是一整节"],
         ("按节切", "节内不再切，超嵌入上限才在句末二次切"),
         ["单对象说全、不碎", "程序内跨段落流程不被切开"],
         "表 8-4 · FR-STO-11 · FR-GEN-10", GRN, GRN_BG),
        ("证据层", "源码原文",
         ["段落原文，逐段成片", "元数据缺一不进检索库"],
         ("按段落", "只到段落边界，不做语义切分"),
         ["回指原文、给出处与行号", "默认不作回答依据，标「原始证据」"],
         "FR-STO-03 · FR-STO-02", ORG, ORG_BG),
    ]
    ry = 152
    for name, sub, c1, c2, c3, ref, color, bg in rows:
        h = 108
        b.append(r(64, ry, 1372, h, "#FFFFFF", color, 1.5, rx=5))
        b.append(r(64, ry, 8, h, color, "none", 0, rx=3))
        b.append(t(90, ry + 34, name, 18, INK, "bold"))
        b.append(t(90, ry + 58, sub, 13, SUB))
        b.append(t(90, ry + 84, ref, 11.6, MUTE))

        b.append(ln(250, ry + 14, 250, ry + h - 14, LINE, 1))
        b.append(t(270, ry + 26, "存什么", 11.6, MUTE, "bold"))
        for i, s in enumerate(c1):
            b.append(t(270, ry + 50 + i * 24, s, 13.2, INK))

        b.append(ln(700, ry + 14, 700, ry + h - 14, LINE, 1))
        b.append(t(720, ry + 26, "怎么切", 11.6, MUTE, "bold"))
        b.append(t(720, ry + 56, c2[0], 16, color, "bold"))
        b.append(t(720, ry + 80, c2[1], 12.2, SUB))

        b.append(ln(1040, ry + 14, 1040, ry + h - 14, LINE, 1))
        b.append(t(1060, ry + 26, "解决什么", 11.6, MUTE, "bold"))
        b.append(t(1060, ry + 52, c3[0], 14, INK, "bold"))
        b.append(t(1060, ry + 78, c3[1], 12.6, SUB))
        ry += h + 14

    b.append(t(64, 536, "注意：三层不是三道工序，是同一个对象的三份表示——图上说「还有 B、C 两个程序」，"
                        "切片才知道 B、C 里写了什么。没有成文层，图只告诉你位置；没有图，切片不知道链有多长。", 12.8, SUB))

    # 下半：两条路径
    b.append(t(64, 592, "② 同一道题，两种召回路径", 16.5, INK, "bold"))
    b.append(t(64, 616, "问：这笔交易在什么情况下会被压单？（答案的 ①②③ 分在三个程序里）", 13.2, SUB))

    b.append(r(64, 636, 620, 348, "#FAFBFC", "#D9D9D9", 1.2, rx=5))
    b.append(t(88, 664, "现在：纯向量 top-k", 15, INK, "bold"))
    b.append(t(88, 686, "一次相似度排序决定全部上下文", 12.4, SUB))
    left = [("问题", "「压单」这个词出现在哪", None),
            ("向量检索，取最相似的 5 片", "命中 03、06，各取一片", None),
            ("按相似度顺序拼接 context", "顺序 = 相似度，不是逻辑先后", None),
            ("作答", "拿到 ①③，缺 ②", RED)]
    ly = 702
    for i, (txt, note, col) in enumerate(left):
        h = 54
        b.append(r(88, ly, 572, h, RED_BG if col else "#FFFFFF", col or "#D9D9D9", 1.4, rx=4))
        b.append(circ(110, ly + 27, 10, col or "#D9D9D9", "none", 0))
        b.append(t(110, ly + 31, str(i + 1), 11.5, "#FFFFFF", "bold", "middle"))
        b.append(t(132, ly + 24, txt, 13.4, INK, "bold" if col else "normal"))
        b.append(t(132, ly + 44, note, 12.0, col or SUB))
        if i < 3:
            b.append(tri(374, ly + h + 3, "#404040"))
        ly += h + 14

    b.append(r(760, 636, 676, 348, "#F7FBF8", GRN, 1.4, rx=5))
    b.append(t(784, 664, "改成：图先定范围，再在范围内检索", 15, GRN, "bold"))
    b.append(t(784, 686, "把一次多跳拆成多次单跳，每次单跳都是确定的", 12.4, SUB))
    right = [("问题", "同上"),
             ("判类型", "单点取值 / 关系链路 / 解释"),
             ("图查询算对象闭包", "程序集合 + 表集合，按需查询不限 2 跳"),
             ("以闭包限定范围做切片检索", "只在闭包内的对象里排相似度"),
             ("按闭包校验：链上对象是否都有产物", "缺则补检，不靠调大 k"),
             ("作答", "①②③ 齐，路径只画参与结论的节点")]
    ry2 = 700
    for i, (txt, note) in enumerate(right):
        h = 38
        b.append(r(784, ry2, 628, h, "#FFFFFF", "#BFDCC9", 1.2, rx=4))
        b.append(circ(806, ry2 + 19, 9, GRN, "none", 0))
        b.append(t(806, ry2 + 23, str(i + 1), 11, "#FFFFFF", "bold", "middle"))
        b.append(t(826, ry2 + 24, txt, 13.4, INK))
        b.append(t(1398, ry2 + 24, note, 11.8, SUB, "normal", "end"))
        if i < 5:
            b.append(tri(1098, ry2 + h + 1, GRN, 10, 6))
        ry2 += h + 8

    b.append(t(64, 1026, "两条路径的差别不在多了一次查询，而在「谁决定取哪几片」：左图由相似度决定，右图由结构决定，相似度只在结构内排序。", 13, INK, "bold"))
    b.append(t(64, 1056, "图查询口径见 FR-QA-08（事实类结论只来自预定义图查询或源码原文）、FR-GRP-02 的约束段（按需查询不受 2 跳限制）；"
                        "校验是否要在需求书里写成一条要求，目前没有，是待定项。", 12.4, MUTE))
    return wrap(W, H, "".join(b))


# ---------------------------------------------------------------- 图 3
def fig3():
    W, H = 1500, 702
    b = [defs("c")]
    b.append(t(64, 52, "图 3　知识文档：有的链写成了一份，有的只到第一跳", 23, INK, "bold"))
    b.append(t(64, 80, "同一个「跨对象成链」的需求，两类对象一个已覆盖、一个未覆盖；差别不在资料多少，在有没有一份成文链路", 13.6, SUB))

    # 左：批处理
    lx, lw = 64, 620
    b.append(r(lx, 116, lw, 420, "#F7FBF8", GRN, 1.5, rx=5))
    b.append(t(lx + 22, 146, "批处理：已覆盖", 15.5, GRN, "bold"))
    b.append(t(lx + 148, 146, "FR-GEN-11", 12.2, GRN, "bold"))
    b.append(t(lx + 22, 170, "作业与批处理说明：作业步与 DD 清单、执行顺序、共享数据集", 12.6, SUB))

    steps = [("作业 BATCH01", "夜间批，日均 1 次", GRN),
             ("作业步 STEP10　→　PGM01", "DD：ACCTIN　·　读账户主文件", INK),
             ("作业步 STEP20　→　PGM02", "DD：ACCTOUT　·　写结果文件", INK),
             ("共享数据集 ACCTFILE", "两步都访问，顺序由作业步定", INK)]
    sy = 196
    for i, (head, note, col) in enumerate(steps):
        b.append(r(lx + 40, sy, lw - 80, 62, "#FFFFFF", "#BFBFBF", 1.3, rx=4))
        b.append(t(lx + 60, sy + 26, head, 13.8, col, "bold"))
        b.append(t(lx + 60, sy + 47, note, 12.2, SUB))
        if i < 3:
            b.append(ln(lx + lw / 2, sy + 63, lx + lw / 2, sy + 78, "#404040", 1.6, None, "mcg"))
        sy += 80
    b.append(t(lx + 22, 522, "一处成文，问答问到「这两步谁先谁后」时，有唯一出处可引", 12.2, GRN))

    # 右：联机交易
    rx_, rw_ = 816, 620
    b.append(r(rx_, 116, rw_, 420, "#FDF9F6", ORG, 1.5, rx=5))
    b.append(t(rx_ + 22, 146, "联机交易：只有绑定", 15.5, ORG, "bold"))
    b.append(t(rx_ + 184, 146, "无对应条目", 12.2, ORG, "bold"))
    b.append(t(rx_ + 22, 170, "交易码到程序有绑定，程序到程序、程序到表没有成文链路", 12.6, SUB))

    tsteps = [("交易码 BALQ", "来源：CSD 资源定义，唯一确定", BLUE, None),
              ("入口程序 PGMAA", "来源：同上，绑定到程序名", BLUE, None),
              ("子程序 PGMBB", "来源：仅 PGMAA 第 7 节「依赖关系」", ORG, "5 4"),
              ("表 ACCTFILE", "来源：各自程序的表访问，逐节分散", ORG, "5 4")]
    sy = 196
    for i, (head, note, col, dash) in enumerate(tsteps):
        b.append(r(rx_ + 60, sy, rw_ - 120, 62, "#FFFFFF", col, 1.4, rx=4, dash=dash))
        b.append(t(rx_ + 80, sy + 26, head, 13.8, INK, "bold"))
        b.append(t(rx_ + 80, sy + 47, note, 12.2, col))
        if i < 3:
            b.append(ln(rx_ + rw_ / 2, sy + 63, rx_ + rw_ / 2, sy + 78,
                        "#404040" if dash is None else ORG, 1.6,
                        None if dash is None else "5 4", "mcg"))
        sy += 80
    b.append(t(rx_ + 22, 522, "链路能拼出来，但要人逐程序翻第 7 节；没有一份成文产物，问答无从引", 12.2, ORG))

    # 中间缺口
    b.append(ln(750, 140, 750, 512, RED, 1.6, "6 5"))
    b.append(r(710, 292, 80, 34, "#FFFFFF", RED, 1.2, rx=4))
    b.append(t(750, 314, "缺 口", 14, RED, "bold", "middle"))
    b.append(r(708, 336, 84, 36, "#FFFFFF", "none", 0, rx=3))
    b.append(t(750, 350, "同一形态", 11.4, RED, "normal", "middle"))
    b.append(t(750, 366, "未写成条目", 11.4, RED, "normal", "middle"))

    # 补法
    b.append(r(64, 566, 1372, 96, GRN_BG, GRN, 1.4, rx=5, dash="5 3"))
    b.append(t(84, 594, "补法：照 FR-GEN-11 的形态，补一份「交易链路说明」", 14.5, GRN, "bold"))
    b.append(t(84, 618, "节点取程序名与表名，边只取调用边与读写边——都是结构事实，不违 FR-STO-26 的「边不得由模型推断」；"
                        "级别 L1a、索引通道，免人工放行。", 12.8, "#404040"))
    b.append(t(84, 642, "与作业说明的差别只在入口：作业从 JCL 进，交易从 CSD 进。两者都是「一进多出」，用同一套节点与边即可。", 12.8, "#404040"))
    return wrap(W, H, "".join(b))


# ---------------------------------------------------------------- 图 6 图谱建模
def fig5():
    """图谱按对象建：节点是对象、边是关系、切片靠两个字段挂回节点。"""
    W, H = 1500, 706
    b = [defs("g5")]
    b.append(t(64, 52, "图 6　图谱按对象建：节点不是文件，切片也不在图里", 23, INK, "bold"))
    b.append(t(64, 80, "文件只是 14 类节点之一；切片不进图，它靠「对象编号 + 定位」两个字段挂回节点——"
                       "不是外键，是两步", 13.6, SUB))

    # ── ① 对象层
    b.append(t(64, 118, "① 节点 = 对象", 15, BLUE, "bold"))
    b.append(t(206, 118, "表 8-5 的 14 类：应用系统、程序、段落、记录、字段、文件、表、作业、作业步骤、"
                         "交易、DD 语句、术语、知识条目、经验条目、占位对象", 11.4, SUB))

    def node(x, y, w, h, name, note, col=INK, bg="#FFFFFF", size=12.4):
        b.append(r(x, y, w, h, bg, col, 1.4, rx=4))
        b.append(t(x + 12, y + 19, name, size, col, "bold"))
        b.append(t(x + 12, y + 37, note, 10.6, SUB))

    node(90, 140, 170, 52, "PGMCC", "程序")
    node(90, 206, 170, 48, "§1500 VALID", "段落 · 行 1120–1180")
    node(90, 274, 170, 48, "§1600 UPDATE", "段落 · 行 1190–1246")
    node(560, 206, 160, 48, "AUTHFILE", "文件 · 授权文件", GRN)
    node(560, 274, 160, 48, "ACCTFILE", "文件 · 账户主文件", GRN)
    node(770, 274, 190, 48, "FLAG-POSTPONE", "字段 · 缓记标志", GRN, size=11.6)

    # 包含：PGMCC → §1500（垂直）；PGMCC → §1600（左侧绕开 §1500）
    b.append(ln(175, 192, 175, 206, "#8C8C8C", 1.4, None, "mgg5"))
    b.append(poly([(90, 166), (62, 166), (62, 298), (90, 298)], "#8C8C8C", 1.4, None, "mgg5"))
    # 读写：段落 → 文件
    b.append(ln(260, 230, 560, 230, BLUE, 1.6, None, "mbg5"))
    b.append(ln(260, 298, 560, 298, BLUE, 1.6, None, "mbg5"))
    # 包含：文件 → 字段
    b.append(ln(720, 298, 770, 298, "#8C8C8C", 1.4, None, "mgg5"))

    b.append(t(300, 222, "读写", 11, BLUE, "bold"))
    b.append(t(300, 290, "读写", 11, BLUE, "bold"))
    b.append(t(72, 250, "包含", 10.2, MUTE))
    b.append(t(728, 290, "包含", 10.2, MUTE))
    b.append(t(392, 196, "每段各连各的文件", 11.2, ORG))
    b.append(t(392, 344, "不是「程序 → 两个文件」一条边", 11.2, ORG))

    # ── ② 切片层
    b.append(t(64, 384, "② 切片不在图里：它有自己的元数据，靠两个字段挂回节点", 15, BLUE, "bold"))
    b.append(r(90, 404, 380, 88, "#FFFFFF", LINE, 1.2, rx=5))
    b.append(t(104, 426, "切片 #04", 12.6, INK, "bold"))
    b.append(t(104, 448, "对象编号：PGMCC", 11.2, SUB))
    b.append(t(104, 468, "定位：§1600 · 行 1190–1246", 11.2, SUB))

    b.append(r(500, 404, 380, 88, "#FFFFFF", LINE, 1.2, rx=5))
    b.append(t(514, 426, "切片 #08", 12.6, INK, "bold"))
    b.append(t(514, 448, "对象编号：ACCTFILE", 11.2, SUB))
    b.append(t(514, 468, "定位：数据实体说明 · 第 2 节", 11.2, SUB))

    b.append(r(90, 508, 790, 88, BLUE_BG, BLUE, 1.2, rx=5, dash="5 3"))
    b.append(t(110, 532, "切片 → 节点走两步，不是一步 join：", 13.2, BLUE, "bold"))
    b.append(t(110, 556, "① 元数据「对象编号」找到图上的节点（程序名 / 交易码）", 11.8, "#404040"))
    b.append(t(110, 578, "② 元数据「定位信息」跳到具体位置（行号区间 / 节号 + 段落序号）", 11.8, "#404040"))

    # ── 右栏：四重受限
    b.append(t(1020, 118, "图能长什么样，由四重限制定死", 15, RED, "bold"))
    limits = [
        ("类型受限", "FR-STO-23", "14 类节点、9 类关系。", "清单外不得入库。", 140),
        ("身份受限", "FR-STO-24", "唯一键 = 类型 + 应用系统 + 对象名。", "同一对象出现一百次只是一个节点。", 244),
        ("来源受限", "FR-STO-26", "边只来自结构事实与人工确认条目。", "模型不得推断调用、读写或业务相关性。", 348),
        ("范围受限", "占位节点", "范围外被引用到的落占位节点，", "纳入范围后按同一唯一键回填，不另建。", 452),
    ]
    for head, tag, l1, l2, y in limits:
        b.append(r(1020, y, 416, 92, "#FFFFFF", RED, 1.2, rx=5))
        b.append(t(1036, y + 24, head, 13, RED, "bold"))
        b.append(t(1420, y + 24, tag, 10.8, MUTE, "normal", "end"))
        b.append(t(1036, y + 50, l1, 11.4, "#404040"))
        b.append(t(1036, y + 72, l2, 11.4, "#404040"))

    b.append(r(1020, 566, 416, 112, GRN_BG, GRN, 1.4, rx=5))
    b.append(t(1036, 592, "所以：生成知识不产生节点", 13.4, GRN, "bold"))
    b.append(t(1036, 616, "图里的「知识条目」是「某对象有一条知识」的引用点，", 11.4, "#404040"))
    b.append(t(1036, 638, "不是生成出来的一篇文档。图的形状由解析范围决定，", 11.4, "#404040"))
    b.append(t(1036, 660, "不由生成了多少解释决定。", 11.4, "#404040"))
    return wrap(W, H, "".join(b))


# ---------------------------------------------------------------- 图 7 一行到一条链
def fig6():
    """从一行代码到一条链：边挂段落、两个方向都能走、入口先锚定对象名。"""
    W, H = 1500, 890
    b = [defs("g6")]
    b.append(t(64, 52, "图 7　从一行代码到一条链：边挂段落，两个方向都要能走", 23, INK, "bold"))
    b.append(t(64, 80, "答案常常就在链上某一行的某一句里。要取到它，先得让这一行在图上被「挂对位置」——"
                       "挂到段落，不是挂到程序", 13.6, SUB))

    # ── 上：程序 / 段落 / 文件
    b.append(r(90, 116, 180, 52, BLUE_BG, BLUE, 1.6, rx=5))
    b.append(t(110, 138, "PGMCC", 14.5, INK, "bold"))
    b.append(t(110, 156, "程序 · 一个文件对象", 11, SUB))

    b.append(r(90, 190, 430, 58, "#FFFFFF", BLUE, 1.4, rx=5))
    b.append(t(106, 212, "§1500 VALID", 13.4, INK, "bold"))
    b.append(t(106, 234, "行 1120–1180\u3000｜\u3000READ AUTHFILE 取授权码并校验", 11.2, SUB))

    b.append(r(90, 258, 430, 58, "#FFFFFF", BLUE, 1.4, rx=5))
    b.append(t(106, 280, "§1600 UPDATE", 13.4, INK, "bold"))
    b.append(t(106, 302, "行 1190–1246\u3000｜\u3000WRITE ACCTFILE 更新账户状态", 11.2, SUB))

    b.append(ln(175, 168, 175, 190, "#8C8C8C", 1.4, None, "mgg6"))
    b.append(poly([(90, 142), (66, 142), (66, 287), (90, 287)], "#8C8C8C", 1.4, None, "mgg6"))
    b.append(t(72, 176, "包含", 10.2, MUTE))

    b.append(r(700, 190, 170, 46, GRN_BG, GRN, 1.4, rx=5))
    b.append(t(716, 210, "AUTHFILE", 12.6, GRN, "bold"))
    b.append(t(716, 228, "授权文件", 10.6, SUB))
    b.append(r(700, 258, 170, 46, GRN_BG, GRN, 1.4, rx=5))
    b.append(t(716, 278, "ACCTFILE", 12.6, GRN, "bold"))
    b.append(t(716, 296, "账户主文件", 10.6, SUB))

    b.append(ln(520, 219, 700, 213, BLUE, 1.8, None, "mbg6"))
    b.append(ln(520, 287, 700, 281, BLUE, 1.8, None, "mbg6"))
    b.append(t(556, 204, "读写（段落级）", 10.6, BLUE, "bold"))
    b.append(t(556, 320, "读写（段落级）", 10.6, BLUE, "bold"))

    b.append(r(800, 116, 500, 44, ORG_BG, ORG, 1.2, rx=4, dash="4 3"))
    b.append(t(816, 144, "若只在 PGMCC 上落一条边，这两段的差别就被抹平", 11.4, ORG))

    # ── 中：两个提问
    b.append(t(64, 366, "同一份图，两个问题，走的是两条不同的路", 15, BLUE, "bold"))

    b.append(r(90, 382, 620, 84, "#FFFFFF", BLUE, 1.3, rx=5))
    b.append(t(106, 404, "问 A｜AUTHFILE 的授权码在哪一段校验？", 12.8, INK, "bold"))
    b.append(t(106, 428, "入口按对象名锚定 → 闭包只到 §1500 → 命中行 1142–1158", 11.4, SUB))
    b.append(t(106, 450, "答得全：就是这一段，出处可跳行", 11.2, GRN, "bold"))

    b.append(r(740, 382, 660, 84, "#FFFFFF", ORG, 1.3, rx=5))
    b.append(t(756, 404, "问 B｜ACCTFILE 被哪一段写过？", 12.8, INK, "bold"))
    b.append(t(756, 428, "锚定 ACCTFILE → 逆「读写」边回到 §1600 → 再逆「调用」边到上游", 11.4, SUB))
    b.append(t(756, 450, "只沿出边走一步都走不出去——这是 FR-STO-32 要定的方向", 11.2, ORG, "bold"))

    # ── 下：逆查三跳（路径方向由数据指向入口程序）
    b.append(t(64, 508, "逆着走：从被访问的数据回到改它的那一段，再回到调用链的上游", 15, ORG, "bold"))
    chain = [("ACCTFILE", "被写的文件", 90), ("PGMCC", "§1600 写它", 330),
             ("PGMBB", "调用 PGMCC", 570), ("PGMAA", "入口程序", 810)]
    for name, note, x in chain:
        b.append(r(x, 528, 190, 52, ORG_BG, ORG, 1.4, rx=5))
        b.append(t(x + 14, 550, name, 12.6, ORG, "bold"))
        b.append(t(x + 14, 570, note, 10.6, SUB))
    for x in (280, 520, 760):
        b.append(ln(x, 554, x + 50, 554, ORG, 1.6, "5 4", "mog6"))
    b.append(t(96, 606, "反向遍历 · FR-STO-32：读写与调用都支持反查；反向得到的路径要标方向，"
                        "不得展示成正向调用顺序", 11.4, ORG))

    # ── 右下：三张卡
    b.append(t(1060, 508, "这条路上三个开关", 15, BLUE, "bold"))
    knobs = [("入口", "FR-QA-24", "问题里有对象名，先按唯一键取节点", "不让相似度决定入口", 524),
             ("范围", "FR-STO-30", "候选集是闭包，闭包里没切片", "的显式标注，不以条数替代覆盖", 600),
             ("方向", "FR-STO-32", "逐类关系定可遍历方向", "读写与调用支持反向", 676)]
    for head, tag, l1, l2, y in knobs:
        b.append(r(1060, y, 376, 72, "#FFFFFF", BLUE, 1.2, rx=5))
        b.append(t(1074, y + 20, head, 12.4, BLUE, "bold"))
        b.append(t(1420, y + 20, tag, 10.4, MUTE, "normal", "end"))
        b.append(t(1074, y + 42, l1, 10.8, "#404040"))
        b.append(t(1074, y + 60, l2, 10.8, "#404040"))

    # ── 收口
    b.append(ln(64, 748, 1436, 748, LINE, 1))
    b.append(r(64, 764, 1372, 100, GRN_BG, GRN, 1.4, rx=5, dash="5 3"))
    b.append(t(84, 792, "图给「哪几段相关」，成文给「先后顺序」，行号给「回到哪一行」——"
                        "三者各答一件事，谁也替不了谁。", 13.6, GRN, "bold"))
    b.append(t(84, 818, "图只会说 PGMAA、PGMBB、PGMCC 与 ACCTFILE 同属一条链，它不告诉你先走哪个；"
                        "顺序只能来自交易说明的调用链或调用点的行号。", 12, "#404040"))
    b.append(t(84, 840, "也正因如此，图缺一条边就等于答案缺一段——而且不报错，只少说一句。", 12, "#404040"))
    return wrap(W, H, "".join(b))


if __name__ == "__main__":
    for name, fn in (("_r0.svg", fig0), ("_r1.svg", fig1), ("_r2.svg", fig2),
                     ("_r3.svg", fig3), ("_r4.svg", fig4), ("_r5.svg", fig5), ("_r6.svg", fig6)):
        s = fn()
        io.open(os.path.join(D, name), "w", encoding="utf-8", newline="\n").write(s)
        print(name, len(s))
