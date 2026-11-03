# -*- coding: utf-8 -*-
"""
生成「资料管理全流程」专业流程图（图 3-1）与「动作的四层落点」分层图（图 3-2）。
输出：_fig1.svg / _fig2.svg / _fig1.html / _fig2.html（供 Edge 无头截图）
"""
import io, os

OUT = os.path.dirname(os.path.abspath(__file__))
FONT = "'Microsoft YaHei','PingFang SC','Hiragino Sans GB',sans-serif"

C_SYS_F, C_SYS_L = "#EDF3FA", "#2F5597"
C_HUM_F, C_HUM_L = "#FDF3E7", "#C55A11"
C_BR_F,  C_BR_L  = "#FFFFFF", "#8C8C8C"
C_RED,   C_RED_F = "#C00000", "#FDECEC"
C_OK,    C_OK_F  = "#1E7B45", "#EAF5EE"
C_TXT,   C_MUT   = "#1F1F1F", "#595959"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def wu(ch):
    return 1.0 if ord(ch) > 127 else 0.53


def wrap(text, maxu):
    """按显示宽度折行；优先在标点、空格、箭头处断，避免把词拆开。"""
    lines, cur, cu, last_ok = [], "", 0.0, -1
    for ch in text:
        cur += ch
        cu += wu(ch)
        if ch in "、，。；：）→ ":
            last_ok = len(cur)
        if cu > maxu:
            if last_ok > 0 and sum(wu(c) for c in cur[:last_ok]) > maxu * 0.55:
                lines.append(cur[:last_ok])
                cur = cur[last_ok:]
            else:
                lines.append(cur[:-1])
                cur = ch
            cu = sum(wu(c) for c in cur)
            last_ok = -1
    if cur:
        lines.append(cur)
    return lines


def txt(x, y, s, size=14, color=C_TXT, weight="normal", anchor="start", ls=0):
    return ('<text x="%.1f" y="%.1f" font-family=%s font-size="%s" fill="%s" '
            'font-weight="%s" text-anchor="%s" letter-spacing="%s">%s</text>'
            % (x, y, FONT, size, color, weight, anchor, ls, esc(s)))


def rect(x, y, w, h, fill, stroke, rx=6, lw=1.6, dash=None, opacity=None):
    d = ' stroke-dasharray="%s"' % dash if dash else ""
    o = ' fill-opacity="%s"' % opacity if opacity else ""
    return ('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="%s" '
            'fill="%s"%s stroke="%s" stroke-width="%s"%s/>'
            % (x, y, w, h, rx, fill, o, stroke, lw, d))


DEFS = """<defs>
<marker id="ag" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
  <path d="M 0 0 L 10 5 L 0 10 z" fill="#404040"/></marker>
<marker id="ab" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
  <path d="M 0 0 L 10 5 L 0 10 z" fill="#2F5597"/></marker>
<marker id="ar" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
  <path d="M 0 0 L 10 5 L 0 10 z" fill="#C00000"/></marker>
</defs>"""


def path(d, color="#404040", mk="ag", dash=None, lw=1.6):
    da = ' stroke-dasharray="%s"' % dash if dash else ""
    return ('<path d="%s" fill="none" stroke="%s" stroke-width="%s"%s marker-end="url(#%s)"/>'
            % (d, color, lw, da, mk))


def line(d, color="#404040", dash=None, lw=1.6):
    da = ' stroke-dasharray="%s"' % dash if dash else ""
    return '<path d="%s" fill="none" stroke="%s" stroke-width="%s"%s/>' % (d, color, lw, da)


def label(x, y, s, size=12.5, color=C_MUT):
    w = sum(wu(c) for c in s) * size * 0.98 + 10
    return (rect(x - w / 2, y - size + 1, w, size + 6, "#FFFFFF", "none", rx=3, lw=0)
            + txt(x, y, s, size, color, anchor="middle"))


# ---------------------------------------------------------------- 图 3-1

def fig1():
    W = 1500
    BX, BW = 400, 680           # 主流程列
    CX = BX + BW / 2
    RX, RW = 1140, 330          # 分支列
    RCX = RX + RW / 2

    s = [DEFS]
    s.append(rect(0, 0, W, 1, "#FFFFFF", "none", lw=0))
    s.append(txt(64, 54, "图 3-1　资料管理全流程（导入 · 覆盖 · 更新 · 撤销 · 失效）",
                 24, C_TXT, "bold"))
    s.append(txt(64, 84, "依据《核心系统迁移知识库需求规格书》v3.18；绿框为 v3.18 补入的处置规则，框内灰字为需求条目编号；逐文件判定见 图 3-2",
                 14, C_MUT))

    # 图例
    lx = 64
    ly = 104
    s.append(rect(lx, ly, 26, 18, C_SYS_F, C_SYS_L, rx=3, lw=1.3))
    s.append(txt(lx + 32, ly + 14, "系统处理", 12.5, C_MUT))
    lx += 130
    s.append(rect(lx, ly, 26, 18, C_HUM_F, C_HUM_L, rx=3, lw=1.3))
    s.append(txt(lx + 32, ly + 14, "人工动作", 12.5, C_MUT))
    lx += 130
    s.append(rect(lx, ly, 26, 18, C_BR_F, C_BR_L, rx=3, lw=1.3))
    s.append(txt(lx + 32, ly + 14, "判定结果 / 终止", 12.5, C_MUT))
    lx += 170
    s.append(rect(lx, ly, 26, 18, C_OK_F, C_OK, rx=3, lw=1.4))
    s.append(txt(lx + 32, ly + 14, "v3.18 补入的规则", 12.5, C_MUT))
    lx += 180
    s.append(rect(lx, ly, 26, 18, C_RED_F, C_RED, rx=4, lw=1.3, dash="3 2"))
    s.append(txt(lx + 32, ly + 14, "待确认口径（表 A-3）", 12.5, C_MUT))

    nodes = [
        dict(n="01", t="提交批次", a="人工", r="FR-IMP-21",
             d=["手工提交，系统不扫描目录、不自动建立批次；提交即触发预检"]),
        dict(n="02", t="预检与去重判定", a="系统", r="FR-IMP-10、FR-IMP-15、FR-IMP-18、FR-IMP-20、表 4-5",
             d=["按路径、内容指纹与文件名主干判定：重复提交 / 同路径变更 / 新增 / 疑似同一份",
                "「疑似同一份」在预检中列候选，由人指定接续，不自动并入",
                "扩展名初判、内容特征确认；同一文档的多种格式标为副本，原件全留、只解析一份"]),
        dict(n="03", t="归属与标签", a="人工", r="FR-IMP-05、表 4-2",
             d=["应用系统必填且唯一；业务线、资料类型、来源、版本、密级由人选定，系统不得代选",
                "选不出时置「待确定」，后续补齐"]),
        dict(n="04", t="原件留存", a="系统", r="FR-IMP-03、表 11-1",
             d=["只读快照，与系统同生命周期；转码只改字符编码，不改内容"]),
        dict(n="05", t="解析", a="系统", r="FR-PRS-09、FR-PRS-10、FR-PRS-13、表 5-1",
             d=["程序出段落与调用关系、copybook 出字段、JCL 出作业链路与 DD 映射",
                "文档按不跨节的段落块切，表格按工作表切（合并单元格取左上角）",
                "未识别的项登记并从可切片范围排除，不得以推测值填充"]),
        dict(n="06", t="知识生成", a="系统", r="FR-GEN-05、FR-GEN-06、FR-GEN-13、FR-GEN-14、FR-GEN-32",
             d=["按范围全量成文；节数与节序不随资料增减，缺件只在节后标「资料不足」并注明缺哪一类",
                "齐备度按五档评估（表 6-6）：仅程序 → 加 copybook → 加 JCL → 加技术文档 → 加业务文档",
                "资料变更后只重算直接引用它的单元与四类跨文件汇总（含项目知识总览），不触发全系统重算"]),
        dict(n="07", t="入库建图", a="系统", r="FR-STO-02、FR-STO-05、FR-STO-24、表 8-2",
             d=["切片五项元数据缺一不进检索库；同一对象按唯一键收敛；范围外引用只落占位节点"]),
        dict(n="08", t="问答使用", a="系统", r="FR-QA-05、表 7-5、表 10-2",
             d=["索引通道答事实、解释通道答含义；按来源级别作答，事实类须给出处与行号",
                "资料失效后其产物为「已废止」，不进入回答依据"]),
    ]

    y = 168
    pos = {}
    ys = []
    for nd in nodes:
        lines = []
        for para in nd["d"]:
            lines += wrap(para, 42)
        h = 96 + 21 * len(lines)
        ys.append((y, h, nd, lines))
        pos[nd["n"]] = (y, h)
        y += h + 76

    # 主链竖向箭头（先画，压在框下）
    for i in range(len(ys) - 1):
        y0, h0 = ys[i][0] + ys[i][1], ys[i + 1][0]
        s.append(path("M %d %d L %d %d" % (CX, y0, CX, y - 0) if False else
                      "M %d %d L %d %d" % (CX, y0, CX, ys[i + 1][0] - 2), "#404040", "ag"))

    # 主干节点
    for (yy, h, nd, lines) in ys:
        fill, stroke = (C_HUM_F, C_HUM_L) if nd["a"] == "人工" else (C_SYS_F, C_SYS_L)
        s.append(rect(BX, yy, BW, h, fill, stroke))
        s.append('<circle cx="374" cy="%.1f" r="17" fill="#FFFFFF" stroke="%s" stroke-width="1.8"/>'
                 % (yy + 30, stroke))
        s.append(txt(374, yy + 35, nd["n"], 14.5, stroke, "bold", "middle"))
        s.append(txt(BX + 26, yy + 36, nd["t"], 19, C_TXT, "bold"))
        # 责任主体标签
        tw = 62 if nd["a"] == "人工" else 62
        s.append(rect(BX + BW - tw - 22, yy + 15, tw, 24, "#FFFFFF", stroke, rx=12, lw=1.2))
        s.append(txt(BX + BW - tw / 2 - 22, yy + 31, nd["a"], 12.5, stroke, "bold", "middle"))
        s.append(txt(BX + 26, yy + 60, "依据：" + nd["r"], 12.8, C_MUT))
        ty = yy + 84
        for ln in lines:
            s.append(txt(BX + 26, ty, ln, 14.6, C_TXT))
            ty += 21
        if nd.get("gap"):
            gx, gy = 286, yy + h - 36
            s.append(rect(gx, gy, 108, 26, C_RED_F, C_RED, rx=4, lw=1.3, dash="3 2"))
            s.append(txt(gx + 54, gy + 18, "缺口 " + nd["gap"], 13, C_RED, "bold", "middle"))
            s.append(line("M %d %d L %d %d" % (gx + 108, gy + 13, BX, gy + 13), C_RED, "3 3", 1.2))

    # ---- 左侧阶段括号 ----
    groups = [("接入", 0, 3), ("解析", 4, 4), ("生成", 5, 5), ("入库", 6, 6), ("问答", 7, 7)]
    for gname, a, b in groups:
        y0 = ys[a][0] + 6
        y1 = ys[b][0] + ys[b][1] - 6
        s.append(line("M %d %d L 196 %d L 196 %d L %d %d" % (206, y0, y0, y1, 206, y1),
                      "#BFBFBF", None, 1.2))
        mid = (y0 + y1) / 2
        s.append('<text x="180" y="%.1f" font-family=%s font-size="12.6" fill="#8C8C8C" '
                 'text-anchor="middle" transform="rotate(-90 180 %.1f)">%s</text>'
                 % (mid, FONT, mid, gname))

    # ---- 右侧分支 ----
    y2, h2 = pos["02"]
    cy2 = y2 + h2 / 2

    # R1 重复提交
    r1y, r1h = y2 + 6, 86
    s.append(rect(RX, r1y, RW, r1h, C_BR_F, C_BR_L))
    s.append(txt(RX + 18, r1y + 26, "重复提交", 15.5, C_TXT, "bold"))
    s.append(txt(RX + 18, r1y + 48, "并入已有记录，不进入后续环节", 13, C_MUT))
    s.append(txt(RX + 18, r1y + 68, "沿用上批次结果，不再解析", 13, C_MUT))

    # R2 同路径变更
    r2y, r2h = y2 + 140, 118
    s.append(rect(RX, r2y, RW, r2h, C_BR_F, C_BR_L))
    s.append(txt(RX + 18, r2y + 26, "同路径变更", 15.5, C_TXT, "bold"))
    s.append(txt(RX + 18, r2y + 48, "新版本入库、保留历史版本", 13, C_MUT))
    s.append(txt(RX + 18, r2y + 68, "汇回 05 解析 → 06 重生成", 13, C_MUT))
    s.append(txt(RX + 18, r2y + 88, "旧切片与旧版知识文档下线", 13, C_MUT))

    s.append(path("M %d %d L %d %d" % (BX + BW, y2 + 34, RX - 2, r1y + 44), "#404040", "ag"))
    s.append(label(BX + BW + 34, y2 + 34 - 6, "重复"))
    s.append(path("M %d %d L %d %d" % (BX + BW, y2 + 80, RX - 2, r2y + 50), "#404040", "ag"))
    s.append(label(BX + BW + 42, y2 + 80 - 6, "变更"))
    s.append(label(CX + 46, y2 + h2 + 40, "新增 → 继续主链"))

    # R2 汇回 05
    y5, h5 = pos["05"]
    cy5 = y5 + h5 / 2
    s.append(path("M %d %d L %d %d L %d %d" % (RCX, r2y + r2h, RCX, cy5, BX + BW + 2, cy5),
                  "#404040", "ag"))

    # R3 撤销批次 → 处置
    y7, h7 = pos["07"]
    cy7 = y7 + h7 / 2
    r3y, r3h = cy7 - 46, 112
    s.append(rect(RX, r3y, RW, r3h, C_HUM_F, C_HUM_L))
    s.append(txt(RX + 18, r3y + 26, "撤销批次", 15.5, C_TXT, "bold"))
    s.append(txt(RX + RW - 78, r3y + 26, "人工", 12.5, C_HUM_L, "bold"))
    s.append(txt(RX + 18, r3y + 48, "依据：FR-IMP-16", 12.8, C_MUT))
    s.append(txt(RX + 18, r3y + 70, "按批次撤销已完成的导入", 13, C_TXT))
    s.append(txt(RX + 18, r3y + 90, "原件快照保留，不删除", 13, C_TXT))
    s.append(path("M %d %d L %d %d" % (BX + BW, cy7, RX - 2, cy7), "#404040", "ag"))
    s.append(label((BX + BW + RX) / 2 - 6, cy7 - 9, "撤销"))

    # R3-1 知识层处置（v3.18 补入，绿框）
    gy = r3y + r3h + 34
    gh = 154
    s.append(path("M %d %d L %d %d" % (RCX, r3y + r3h, RCX, gy - 2), C_OK, "ag"))
    s.append(rect(RX, gy, RW, gh, C_OK_F, C_OK, lw=1.6))
    s.append(txt(RX + 18, gy + 26, "知识层处置", 15.5, C_OK, "bold"))
    s.append(txt(RX + RW - 96, gy + 26, "FR-GEN-33", 12.2, C_OK, "bold"))
    lines_g = ["知识文档转「已废止」，引用改占位",
               "切片与向量退出检索（FR-STO-09）",
               "图谱节点与关系下线（FR-STO-27）",
               "四类汇总产物随之重算"]
    for i, ln in enumerate(lines_g):
        s.append(txt(RX + 18, gy + 50 + i * 21, ln, 12.8, C_TXT))

    # R4 单份停用 / 整包重传未出现（同一处置的另两个触发）
    r4y = gy + gh + 30
    r4h = 132
    s.append(path("M %d %d L %d %d" % (RCX, gy + gh, RCX, r4y - 2), C_HUM_L, "ag"))
    s.append(rect(RX, r4y, RW, r4h, C_HUM_F, C_HUM_L))
    s.append(txt(RX + 18, r4y + 26, "单份停用 · 整包重传未出现", 15, C_TXT, "bold"))
    s.append(txt(RX + RW - 78, r4y + 26, "人工", 12.5, C_HUM_L, "bold"))
    s.append(txt(RX + 18, r4y + 48, "依据：FR-GEN-33 表 6-9、FR-IMP-21", 12.4, C_MUT))
    s.append(txt(RX + 18, r4y + 70, "单份停用可恢复；整包须先声明范围，", 12.8, C_TXT))
    s.append(txt(RX + 18, r4y + 90, "未声明时不做比对", 12.8, C_TXT))
    s.append(txt(RX + 18, r4y + 112, "两者处置同上，同批其他资料不受影响", 12.4, C_MUT))

    last_bottom = max(ys[-1][0] + ys[-1][1], r4y + r4h)
    bottom = last_bottom + 44
    H = int(bottom + 46)
    s.append(line("M 64 %d L %d %d" % (bottom + 6, W - 64, bottom + 6), "#D9D9D9", None, 1))
    s.append(txt(64, bottom + 34, "注：右侧四条为资料再次提交、撤销或失效时的路径；绿框为 v3.18 补入的处置规则，撤销落到哪一层见图 3-3。", 12.6, C_MUT))

    out = ['<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
           % (W, H, W, H)]
    out.append(rect(0, 0, W, H, "#FFFFFF", "none", lw=0))
    out += s
    out.append("</svg>")
    return "\n".join(out), W, H


# ---------------------------------------------------------------- 图 3-2

def fig2():
    W = 1500
    BX0, BW0 = 150, 1000
    DIV = 750
    cols = [("导入", 810), ("更新", 930), ("撤销", 1050)]
    rows = [
        ("第一层", "原件快照", "原始文件只读副本", "不可重建 · 与系统同生命周期 · 撤批不删",
         ("执行", "保留", "保留")),
        ("第二层", "解析产物", "解析结果、结构化文本、表格抽取", "可按原件重建",
         ("执行", "替换", "清除")),
        ("第三层", "知识产物", "知识文档、业务术语表、疑点清单", "不可重建 · 保留全部版本",
         ("执行", "另出新版本", "转已废止")),
        ("第四层", "检索索引", "切片、向量、元数据五项、图谱节点", "可按知识产物重建",
         ("执行", "重建", "清除")),
    ]

    s = [DEFS]
    s.append(txt(64, 52, "图 3-3　资料的四层，与撤销动作的落点", 22, C_TXT, "bold"))
    s.append(txt(64, 80, "层序即链路先后：原件 → 解析产物 → 知识产物 → 检索索引；排序依据是链路位置，处置依据看「可否重建」列",
                 13.8, C_MUT))
    s.append(txt(64, 102, "撤销清掉了知识产物的上游（解析结果）与下游（切片、向量、图谱），唯独它本身原先不失效——这是缺口 ① 的由来",
                 12.8, C_MUT))
    hy = 132
    s.append(txt(BX0, hy, "数据层", 13, C_MUT, "bold"))
    for name, cx in cols:
        s.append(txt(cx, hy, name, 13, C_MUT, "bold", "middle"))
    s.append(txt(1176, hy, "可否重建", 13, C_MUT, "bold"))

    y = 158
    for i, (no, name, items, attr, marks) in enumerate(rows):
        h = 88
        last = (marks[2] == "转已废止")
        fill = C_OK_F if last else "#FFFFFF"
        stroke = C_OK if last else "#BFBFBF"
        s.append(rect(BX0, y, BW0, h, fill, stroke, rx=5, lw=1.6 if last else 1.1))
        s.append(line("M %d %d L %d %d" % (DIV, y + 10, DIV, y + h - 10), "#D9D9D9", None, 1.1))
        s.append(txt(BX0 + 18, y + 22, no, 11.6, C_MUT))
        s.append(txt(BX0 + 18, y + 48, name, 16.5, C_TXT, "bold"))
        s.append(txt(BX0 + 18, y + 71, items, 12.4, C_MUT))
        for (mn, mx), mk in zip(cols, marks):
            ok = (mk == "转已废止")
            color = C_OK if ok else ("#1F1F1F" if mk in ("执行", "清除", "替换", "重建") else C_MUT)
            s.append(txt(mx, y + 50, mk, 12.4 if mk == "另出新版本" else 13.2, color,
                         "bold" if ok else "normal", "middle"))
        s.append(txt(1176, y + 50, attr, 12.2, C_MUT))
        y += h + 10

    # 结论说明（表下注，不带框，避免与第四行冲突）
    ty = y + 26
    s.append(rect(BX0, ty - 20, 6, 44, C_OK, C_OK, rx=3, lw=0))
    s.append(txt(BX0 + 20, ty - 2, "已闭合（v3.18）　第三层不再被跳过", 15, C_OK, "bold"))
    s.append(txt(BX0 + 20, ty + 20, "FR-GEN-33 把三种触发（整批撤销、单份停用、整包重传未出现）写成一条规则："
                  "知识产物转「已废止」，其下游切片与向量退出检索、图谱节点下线、四类汇总产物重算，原件快照保留。", 12.6, C_MUT))
    H = ty + 66
    s.append(txt(64, H - 10, "注：知识产物由模型生成、含人工修正，删掉无法按原文复现，故只出新版本、旧版留档；"
                  "表 7-4 的「已废止」进入条件已同步为「源文件变更或资料失效」。", 12.4, C_MUT))

    out = ['<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
           % (W, H, W, H)]
    out.append(rect(0, 0, W, H, "#FFFFFF", "none", lw=0))
    out += s
    out.append("</svg>")
    return "\n".join(out), W, H


# ---------------------------------------------------------------- 图 3-2（逐文件判定）

def fig_tree():
    W = 1500
    EX, EW, EH = 110, 610, 96
    RB, RW2 = 790, 650

    s = [DEFS]
    s.append(txt(64, 52, "图 3-2\u3000一次提交的逐文件判定与重算范围", 22, C_TXT, "bold"))
    s.append(txt(64, 80, "判定单位是文件，不是批次：同一套出口覆盖单份替换、整包重传、多系统分别提交三种形态",
                 13.8, C_MUT))

    # 判定框
    s.append(rect(300, 112, 450, 100, "#F7F7F7", "#595959", rx=6, lw=1.5, dash="5 3"))
    s.append(txt(320, 141, "一次提交：单份 · 若干份 · 整包", 15.5, C_TXT, "bold"))
    s.append(txt(320, 166, "按（路径，内容指纹，文件名主干）逐份比对", 13.4, C_MUT))
    s.append(txt(320, 190, "判定结果逐份落到五个提交出口；人工停用另设一个入口", 13.4, C_MUT))

    # 图例
    lx, ly = 810, 116
    s.append(rect(lx, ly, 24, 16, "#FFFFFF", C_BR_L, rx=3, lw=1.2))
    s.append(txt(lx + 30, ly + 13, "不触发重算", 12.4, C_MUT))
    lx += 128
    s.append(rect(lx, ly, 24, 16, C_SYS_F, C_SYS_L, rx=3, lw=1.2))
    s.append(txt(lx + 30, ly + 13, "触发重算", 12.4, C_MUT))
    lx += 112
    s.append(rect(lx, ly, 24, 16, C_OK_F, C_OK, rx=3, lw=1.3))
    s.append(txt(lx + 30, ly + 13, "v3.18 补入的出口", 12.4, C_MUT))

    exits = [
        ("重复提交", "路径、指纹均与已有记录相同",
         "并入已有记录、累计重复次数；不重解析、不重算", "plain"),
        ("同路径变更", "路径相同、内容指纹不同",
         "新版本入库（留历史）→ 重解析该文件 → 触发重算", "calc"),
        ("新增", "路径与指纹均未出现过",
         "新建记录 → 导入解析 → 触发重算", "calc"),
        ("疑似同一份", "路径与指纹均未命中，文件名主干相同（改名、移动、换格式）",
         "预检列候选，由人指定接续：接续的按「同路径变更」，否则按「新增」", "ok"),
        ("上次有、本次没有", "操作人声明整包重传后，与上一版比对得出",
         "按单份停用处置，清单标「本次未提交」；未声明整包时不做此比对", "ok"),
    ]
    ys = [246, 366, 486, 606, 726]
    for (title, judge, act, kind), yy in zip(exits, ys):
        if kind == "ok":
            fill, stroke, dash = C_OK_F, C_OK, "4 3"
        elif kind == "calc":
            fill, stroke, dash = C_SYS_F, C_SYS_L, None
        else:
            fill, stroke, dash = "#FFFFFF", C_BR_L, None
        s.append(rect(EX, yy, EW, EH, fill, stroke, rx=6, lw=1.5, dash=dash))
        s.append(txt(EX + 18, yy + 27, title, 15.5, C_TXT, "bold"))
        s.append(txt(EX + 18, yy + 50, "判定：" + judge, 12.6, C_MUT))
        s.append(txt(EX + 18, yy + 74, "处置：" + act, 12.8, C_TXT))
        if kind == "ok":
            s.append(txt(EX + EW - 92, yy + 27, "v3.18 补", 11.6, C_OK, "bold", "middle"))

    # 左侧汇流总线
    s.append(line("M 300 164 L 70 164 L 70 774", "#404040", None, 1.6))
    for yy in ys:
        s.append(path("M 70 %d L %d %d" % (yy + EH / 2, EX - 2, yy + EH / 2), "#404040", "ag"))

    # 另一入口：人工单份停用
    uy = 862
    s.append(rect(EX, uy, EW, EH, C_HUM_F, C_HUM_L, rx=6, lw=1.5))
    s.append(txt(EX + 18, uy + 27, "单份停用（另一入口，非提交触发）", 15.5, C_TXT, "bold"))
    s.append(txt(EX + 18, uy + 50, "判定：操作人在导入清单对单份资料执行停用（FR-GEN-33 表 6-9）", 12.6, C_MUT))
    s.append(txt(EX + 18, uy + 74, "处置：同下类失效处置，同批其他资料不受影响；可解除停用后重算", 12.8, C_TXT))

    # 重算范围
    s.append(rect(RB, 300, RW2, 340, "#FFFFFF", C_SYS_L, rx=6, lw=1.5))
    s.append(txt(RB + 18, 332, "重算范围（FR-GEN-32）", 16, C_TXT, "bold"))
    rows = [
        ("① 本文件", "另出新版本（FR-GEN-05），旧切片下线（FR-STO-09）"),
        ("② 直接引用它的单元", "COPY 它的程序 · CALL 它的程序 · 访问它指向的表的程序"),
        ("③ 跨文件汇总（四类）", "数据实体说明 · 业务术语表 · 系统结构说明 · 项目知识总览"),
    ]
    ry = 356
    for name, desc in rows:
        s.append(txt(RB + 18, ry, name, 13.6, C_TXT, "bold"))
        s.append(txt(RB + 18, ry + 22, desc, 12.4, C_MUT))
        ry += 68
    s.append(txt(RB + 18, ry + 4, "项目知识总览按条目刷新，不整份重生成", 12.0, C_MUT))

    s.append(rect(RB + 10, 552, RW2 - 20, 76, C_RED_F, C_RED, rx=5, lw=1.3, dash="4 3"))
    s.append(txt(RB + 26, 576, "④ 跨应用系统：引用方挂在别的应用系统下", 13.4, C_RED, "bold"))
    s.append(txt(RB + 26, 598, "按应用系统边界处理，改动侧登记疑点（表 A-3 待确认）", 12.4, C_RED))

    s.append(path("M %d %d L %d %d" % (EX + EW, 414, RB - 2, 414), "#2F5597", "ab"))
    s.append(path("M %d %d L %d %d" % (EX + EW, 534, RB - 2, 534), "#2F5597", "ab"))

    # 底部说明：口径已定
    ny = 990
    s.append(rect(64, ny, 1400, 96, "#FFFFFF", C_OK, rx=5, lw=1.4, dash="4 3"))
    s.append(txt(82, ny + 26, "口径已定（v3.18）　「单元」有定义，重算不限于本批范围", 14.2, C_OK, "bold"))
    s.append(txt(82, ny + 50, "「生成单元」= 一份程序 / 一份文档 / 一张工作表对应的一份产物（术语表）；"
                  "补 copybook 时上一批的程序按引用关系纳入重算，与是否在本批次内无关。", 12.6, C_MUT))
    s.append(txt(82, ny + 72, "出口 4、5 与失效处置的判据同源：表 4-5 按文件名主干判「疑似同一份」，"
                  "表 6-9 按操作人声明的整包范围判缺失。", 12.6, C_MUT))
    H = ny + 96 + 42
    s.append(txt(64, H - 12, "注：出口 4 的接续与人判挂在预检（FR-IMP-15）；出口 5 与「单份停用」的处置同为"
                  "知识文档转「已废止」、切片与向量下线、图谱节点下线、四类汇总重算（FR-GEN-33）。",
                 12.2, C_MUT))

    out = ['<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
           % (W, H, W, H)]
    out.append(rect(0, 0, W, H, "#FFFFFF", "none", lw=0))
    out += s
    out.append("</svg>")
    return "\n".join(out), W, H


if __name__ == "__main__":
    for name, fn in (("_fig1", fig1), ("_fig2", fig2), ("_fig3", fig_tree)):
        svg, w, h = fn()
        p = os.path.join(OUT, name + ".svg")
        io.open(p, "w", encoding="utf-8", newline="\n").write(svg)
        html = ('<!DOCTYPE html><html><head><meta charset="utf-8">'
                '<style>html,body{margin:0;padding:0;background:#fff;}svg{display:block;}</style>'
                '</head><body>%s</body></html>' % svg)
        io.open(os.path.join(OUT, name + ".html"), "w", encoding="utf-8", newline="\n").write(html)
        print(name, w, h)
