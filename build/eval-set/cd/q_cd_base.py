# -*- coding: utf-8 -*-
"""CardDemo 评测集公共定义：题元组工厂 + 字段常量。"""

CAT_DEV = "01 开发与迁移"
CAT_BIZ = "02 业务分析"
CAT_TEL = "03 行员"
CAT_CFT = "04 冲突与不一致"
CAT_QUAL = "05 质量与边界"

# 问题类型（沿用既有九类，第 7 类本素材无图形资料，整类为 0）
T1 = "1 精准命中"
T2 = "2 同义改写+跨语言"
T3 = "3 信息分散（多跳）"
T4 = "4 知识库不存在（幻觉）"
T5 = "5 冲突信息"
T6 = "6 异常与边界"
T8 = "8 多轮对话"
T9 = "9 元数据过滤"

L_CODE = "代码直出"
L_DOC = "引自文档"
L_MAN = "人工修正"
L_M3 = "模型生成·核验通过"
L_M4 = "模型生成·核验未通过"
L_NONE = "—"

_n = [0]


def pts(*a):
    return "\n".join(f"{i + 1}）{s}" for i, s in enumerate(a))


def Q(cat, scope, q, t, diff, inkb, basis, p, src, lvl, state, chk, rule,
      note="", role="否"):
    """一条题 = 18 列。编号按调用顺序自增，保证唯一且连续。"""
    _n[0] += 1
    return (f"EV-CD-{_n[0]:03d}", cat, scope, q, t, diff, role, inkb, basis,
            p, src, lvl, state, chk, rule, "", "", note)


def count():
    return _n[0]
