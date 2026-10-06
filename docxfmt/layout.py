# -*- coding: utf-8 -*-
"""
layout.py —— 页面设置 / 分节 / 分页符 / 页眉页脚 / 制表位。

⛔ 多节文档的坑
    `body/sectPr` **只是最后一节**，前面的节在 `w:p/w:pPr/w:sectPr` 里。
    只改 body/sectPr 会导致「只有最后一节生效」。

⛔ 制表位的坑
    换边距/装订线后，版心变了，既有制表位**不会自动跟着变** →
    右对齐内容直接跑出页面（实测公式编号超页 18.5pt 就是这么来的）。
    所以 rescale_tabs() 必须按**所在节**的版心逐节算，取 max(版心) 会错
    （横向节的宽版心会把纵向节的制表位放大 1.6 倍）。
"""
from lxml import etree

from . import qn
from .docxio import el, get_ppr

DEFAULT_MARGIN = {"top": 1440, "bottom": 1440,
                  "left": 1588, "right": 1474, "header": 851, "footer": 992}


# ------------------------------------------------------------------ 节
def all_sectprs(root):
    """按文档顺序返回所有 sectPr。

    最后一节是 body/sectPr；其余在 w:p/w:pPr/w:sectPr。
    """
    out = []
    for p in root.iter(qn("p")):
        ppr = p.find(qn("pPr"))
        if ppr is None:
            continue
        sp = ppr.find(qn("sectPr"))
        if sp is not None:
            out.append(sp)
    body = root.find(qn("body"))
    if body is not None:
        sp = body.find(qn("sectPr"))
        if sp is not None:
            out.append(sp)
    return out


def sect_text_width(sp):
    """节的版心宽度（twips）。已扣掉装订线。"""
    pg = sp.find(qn("pgSz"))
    mar = sp.find(qn("pgMar"))
    if pg is None or mar is None:
        return None
    w = int(pg.get(qn("w")) or 0)
    left = int(mar.get(qn("left")) or 0)
    right = int(mar.get(qn("right")) or 0)
    gut = int(mar.get(qn("gutter")) or 0)
    return max(1000, w - left - right - gut)


def sect_is_landscape(sp):
    pg = sp.find(qn("pgSz"))
    if pg is None:
        return False
    return (pg.get(qn("orient")) or "") == "landscape"


def apply_page(root, margin=None, orientation="keep", first_page_diff=False,
               odd_even=False):
    """设置页面边距 / 方向 / 首页不同 / 奇偶页不同。

    orientation="keep"：横向节保持横向（宽表改纵向会溢出）。
    """
    m = dict(DEFAULT_MARGIN)
    m.update(margin or {})
    n = 0
    for sp in all_sectprs(root):
        land = sect_is_landscape(sp)
        mar = sp.find(qn("pgMar"))
        if mar is None:
            mar = etree.SubElement(sp, qn("pgMar"))
        for k, v in m.items():
            mar.set(qn(k), str(v))
        pg = sp.find(qn("pgSz"))
        if pg is not None and orientation == "all-portrait" and land:
            w = int(pg.get(qn("w")) or 0)
            h = int(pg.get(qn("h")) or 0)
            pg.set(qn("w"), str(min(w, h)))
            pg.set(qn("h"), str(max(w, h)))
            if pg.get(qn("orient")):
                del pg.attrib[qn("orient")]
        if first_page_diff:
            _set_titlepg(sp)
        n += 1
    if odd_even:
        _enable_even_odd(root)
    return n


def _set_titlepg(sp):
    if sp.find(qn("titlePg")) is None:
        sp.append(el("titlePg"))


def _enable_even_odd(root):
    """settings.xml 里开 evenAndOddHeaders。"""
    pass  # 由 Docx.settings 处理


def snapshot_widths(root):
    """记住每节当前版心宽度，rescale_tabs 要用。必须在 apply_page 之前调用。"""
    return [sect_text_width(sp) for sp in all_sectprs(root)]


# ------------------------------------------------------------------ 分页符
def insert_page_break_before(p):
    """给段落加「段前分页」。

    两种做法：
      A. pPr/pageBreakBefore  —— 推荐，不产生空段落，样式照常生效
      B. 插一个只含 <w:br type="page"/> 的空段 —— 会多一个空行
    """
    ppr = get_ppr(p)
    if ppr.find(qn("pageBreakBefore")) is None:
        # pageBreakBefore 必须排在 keepNext/keepLines 之后、framePr 之前，
        # 这里直接 append，Word 对 pPr 子元素顺序容忍度尚可，
        # 但为稳妥放进 pPr 的第一个位置
        ppr.insert(0, el("pageBreakBefore"))


def set_keep_next(p, on=True):
    ppr = get_ppr(p)
    e = ppr.find(qn("keepNext"))
    if on and e is None:
        ppr.insert(0, el("keepNext"))
    elif not on and e is not None:
        ppr.remove(e)


def apply_page_breaks(root, levels=(1,), keep_with_next=True):
    """按标题级别插分页符。返回 (分页数, keepNext 数)。"""
    np = nk = 0
    levels = set(levels or ())
    for n in _all_nodes(root):
        if n.type != "heading" or n.ref is None:
            continue
        if n.level in levels:
            insert_page_break_before(n.ref)
            np += 1
        if keep_with_next:
            set_keep_next(n.ref, True)
            nk += 1
    return np, nk


def _all_nodes(root):
    out = []
    def rec(n):
        out.append(n)
        for c in n.children:
            rec(c)
    rec(root)
    return out


# ------------------------------------------------------------------ 分节
def insert_section_break_before(p, sectpr_proto=None):
    """在段落 p 之前插入分节符。

    OOXML 语义：分节符 = 一个段落的 pPr 里带 sectPr，
    它**结束**的是「它所在段落及其之前」的那一节。
    所以要在 p 前面插一个带 sectPr 的空段，sectpr_proto 是新节的属性模板。
    """
    body = p.getparent()
    newp = etree.Element(qn("p"))
    ppr = etree.SubElement(newp, qn("pPr"))
    sp = copy_of_sectpr(sectpr_proto)
    if sp is not None:
        ppr.append(sp)
    body.insert(list(body).index(p), newp)
    return newp


def copy_of_sectpr(sp):
    if sp is None:
        return None
    import copy
    new = etree.Element(qn("sectPr"))
    for c in sp:
        if c.tag in (qn("headerReference"), qn("footerReference")):
            continue                      # 不复制页眉页脚引用，否则串节
        new.append(copy.deepcopy(c))
    for a in (qn("rsidR"), qn("rsidSect")):
        if sp.get(a):
            new.set(a, sp.get(a))
    return new


def apply_sections(root, insert_levels=(), landscape_patterns=(),
                   margin=None, log=print):
    """按 spec.section 插分节符。

    分节符的语义（很容易搞反）：分节符 = 一个段落的 pPr 里带 sectPr，
    它**结束**的是「它所在段落及其之前」的那一节。
    所以「让 X 单独成一节」要在 X **前面**插一个带 sectPr 的空段。

    insert_levels     : 哪些标题级别前插分节符，如 (1,) 表示一级标题
    landscape_patterns: 标题文字匹配这些正则的节点所在节设为横向
    """
    import re as _re
    n_sec = n_land = 0
    pats = [_re.compile(p) for p in (landscape_patterns or [])]
    before = []
    for node in _all_nodes(root):
        if node.type != "heading" or node.ref is None:
            continue
        want_land = any(p.search(node.text or "") for p in pats)
        if node.level in (insert_levels or ()) or want_land:
            before.append((node, want_land))
    for node, want_land in before:
        sp = new_sectpr(landscape=want_land, margin=margin)
        insert_section_break_before(node.ref, sp)
        n_sec += 1
        if want_land:
            n_land += 1
    if n_sec:
        log("[section] 插入分节符 %d 处（其中横向节 %d）" % (n_sec, n_land))
    return n_sec, n_land


def new_sectpr(landscape=False, margin=None):
    """造一个全新的 sectPr（可作分节符的节属性）。"""
    sp = etree.Element(qn("sectPr"))
    pg = etree.SubElement(sp, qn("pgSz"))
    if landscape:
        pg.set(qn("w"), "16838")
        pg.set(qn("h"), "11906")
        pg.set(qn("orient"), "landscape")
    else:
        pg.set(qn("w"), "11906")
        pg.set(qn("h"), "16838")
    mar = etree.SubElement(sp, qn("pgMar"))
    m = dict(DEFAULT_MARGIN)
    m.update(margin or {})
    for k, v in m.items():
        mar.set(qn(k), str(v))
    etree.SubElement(sp, qn("cols")).set(qn("space"), "425")
    etree.SubElement(sp, qn("docGrid")).set(qn("linePitch"), "312")
    return sp


# ------------------------------------------------------------------ 制表位
def rescale_tabs(root, old_widths):
    """按版心变化等比缩放段落级制表位。

    old_widths: snapshot_widths() 的返回值（apply_page 之前取）
    """
    sect_list = all_sectprs(root)
    # 建立「段落 -> 所属节序号」的映射
    idx_of = {}
    si = 0
    for p in root.iter(qn("p")):
        ppr = p.find(qn("pPr"))
        if ppr is not None and ppr.find(qn("sectPr")) is not None:
            idx_of[id(p)] = si
            si += 1
    # 没显式归属的段落归到「上一节」
    cur = -1
    order = []
    for p in root.iter(qn("p")):
        if id(p) in idx_of:
            cur = idx_of[id(p)]
        order.append(cur if cur >= 0 else 0)

    n_p = n_t = 0
    for i, p in enumerate(root.iter(qn("p"))):
        si = order[i] if i < len(order) else 0
        if si < 0 or si >= len(sect_list) or si >= len(old_widths):
            continue
        neww = sect_text_width(sect_list[si])
        oldw = old_widths[si]
        if not neww or not oldw or neww == oldw:
            continue
        ppr = p.find(qn("pPr"))
        if ppr is None:
            continue
        tabs = ppr.find(qn("tabs"))
        if tabs is None:
            continue
        changed = False
        for t in tabs.findall(qn("tab")):
            pos = t.get(qn("pos"))
            if not pos:
                continue
            v = int(pos)
            nv = int(round(v * neww / oldw))
            t.set(qn("pos"), str(nv))
            changed = True
            n_t += 1
        if changed:
            n_p += 1
    return n_p, n_t
