# -*- coding: utf-8 -*-
"""
render.py —— 把样式方案套到树上，并抹平直接格式。

「抹平」是关键：直接格式残留会让样式看起来没生效。
但要小心别把 run 删掉 —— run 里可能藏着图片、已有域、批注锚点。
"""
import re
from collections import Counter

from lxml import etree

from . import qn
from .docxio import get_ppr, set_style
from .styles import NAME_ALIASES, StyleTable

# 段落级要抹掉的直接格式（保留 tabs/outlineLvl/sectPr/numPr/pBdr/shd/framePr）
DROP_PARA = ["jc", "spacing", "ind", "snapToGrid", "textAlignment",
             "adjustRightInd", "kinsoku", "wordWrap", "overflowPunct",
             "autoSpaceDE", "autoSpaceDN"]
KEEP_PARA = {"tabs", "outlineLvl", "sectPr", "pStyle", "numPr", "pBdr", "shd",
             "framePr", "rPr"}

# run 级要抹掉的
DROP_RUN = ["rFonts", "sz", "szCs", "kern", "w", "position", "spacing",
            "noProof", "highlight", "lang", "shd", "em", "bdr"]
DROP_RUN_HEADING = DROP_RUN + ["b", "bCs", "i", "iCs", "u", "color"]


def flatten_direct(p, is_heading=False):
    """抹平段落/run 级直接格式，让样式真正生效。"""
    ppr = p.find(qn("pPr"))
    if ppr is not None:
        rpr = ppr.find(qn("rPr"))
        if rpr is not None:
            ppr.remove(rpr)
        for tag in DROP_PARA:
            for e in ppr.findall(qn(tag)):
                ppr.remove(e)
    drop = DROP_RUN_HEADING if is_heading else DROP_RUN
    for r in p.iter(qn("r")):
        rpr = r.find(qn("rPr"))
        if rpr is None:
            continue
        for tag in drop:
            for e in rpr.findall(qn(tag)):
                rpr.remove(e)
        if len(rpr) == 0:
            r.remove(rpr)


def role_of(node):
    """节点 -> (角色, 级别)。"""
    if node.type == "heading":
        return "heading", node.level
    return node.type, 0


def render(root, spec, table, stats=None, dry=False):
    """遍历树给每个节点赋样式。dry=True 时只统计不落笔。"""
    stats = stats if stats is not None else Counter()
    miss = Counter()

    def apply(n):
        if n.type == "document":
            return
        stats[n.type] += 1
        role, lvl = role_of(n)
        want = spec.style_of(role, lvl)
        if not want:
            return                       # null = 不动（封面/空段）
        st = table.resolve(want, NAME_ALIASES.get(want))
        if st is None and role in ("table_text", "paragraph", "caption",
                                   "image"):
            # 目标样式在文档里找不到就退回正文 —— 总得有个样式，
            # 不能因为模板没定义就几千段全不套（踩过：3273 段表格文字没套上）
            body = spec.style_of("paragraph")
            st = table.resolve(body, NAME_ALIASES.get(body))
            if st is not None:
                want = body
                stats["_退回正文"] += 1
        if st is None:
            miss[want] += 1
            return
        sid = st.get(qn("styleId"))
        n.style = want
        if n.ref is None or dry:
            return
        set_style(n.ref, sid)
        flatten_direct(n.ref, is_heading=(n.type == "heading"))

    root.walk(apply)
    stats["_缺样式"] = sum(miss.values())
    return miss


def apply_geometry(spec, table):
    """按 spec.styles.geometry 强制覆盖几何。返回覆盖条数。"""
    n = 0
    for name, kw in (spec["styles"].get("geometry") or {}).items():
        pkw = {k: v for k, v in kw.items()
               if k in ("align", "first_line", "left", "before", "after",
                        "line", "keep_next", "page_break", "keep_lines")}
        rkw = {k: v for k, v in kw.items()
               if k in ("size", "font", "bold", "color")}
        if pkw and table.set_ppr(name, **pkw):
            n += 1
        if rkw and table.set_rpr(name, **rkw):
            n += 1
    return n


def create_new_styles(spec, table):
    """按 spec.styles.new_styles 建新样式（模板里没有的）。"""
    made = []
    for name, kw in (spec["styles"].get("new_styles") or {}).items():
        based = kw.get("based_on")
        if table.get(name) is None:
            table.ensure(name, based_on=based, next_style=kw.get("next"),
                         kind=kw.get("kind", "paragraph"))
            made.append(name)
        pkw = {k: v for k, v in kw.items()
               if k in ("align", "first_line", "left", "before", "after",
                        "line", "keep_next", "page_break", "keep_lines")}
        rkw = {k: v for k, v in kw.items()
               if k in ("size", "font", "bold", "color")}
        if pkw:
            table.set_ppr(name, **pkw)
        if rkw:
            table.set_rpr(name, **rkw)
    return made
