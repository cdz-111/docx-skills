# -*- coding: utf-8 -*-
r"""
toc.py —— 目录 / 图目录 / 表目录（TOC 域）。

三种域的区别只在 `\c` 开关：
    正文目录  { TOC \\o "1-3" \\h \\z \\u }
    图目录    { TOC \\h \\z \\c "图" }     <- 按 SEQ 标识符收集，不是按样式
    表目录    { TOC \\h \\z \\c "表" }

⚠️ 图/表目录能收集到内容的前提：**题注里的 SEQ 域标识符必须与 \\c 一致**。
   本工具的题注用 `SEQ 图` / `SEQ 表`，所以 \\c "图" / \\c "表" 天然对得上。
   旧式「用样式收集」（`TOC \\t "题注,1"`）也能用，但会把图和表混在一起。

⚠️ 目录标题要**新建独立样式**：居中、无首行缩进、与正文区分开
   （老板明确要求过：目录/图目录/表目录的标题居中且样式与正文区分）。
"""
import re

from lxml import etree

from . import qn
from .docxio import el, field_runs, get_ppr, set_style
from .recognize import TOC_TITLES


def _clear(p):
    for c in list(p):
        if c.tag != qn("pPr"):
            p.remove(c)


def make_title_para(text, style_id=None):
    """目录标题段。"""
    p = etree.Element(qn("p"))
    if style_id:
        set_style(p, style_id)
    p.append(_text_run(text))
    return p


def _text_run(t):
    r = etree.Element(qn("r"))
    e = etree.SubElement(r, qn("t"))
    e.text = t
    return r


def make_toc_para(instr, style_id=None, dirty=True):
    """TOC 域段落。

    dirty="true" 让 Word 打开时提示更新域；不写 dirty 的话，
    已有缓存值的目录不会自动刷新（这是「改了标题但目录不变」的常见原因）。
    """
    p = etree.Element(qn("p"))
    if style_id:
        set_style(p, style_id)
    runs = field_runs(instr, cached="按 F9 或右键更新域", rpr=None)
    # fldChar begin 上打 dirty，Word 打开就会问要不要更新
    fc = runs[0].find(qn("fldChar"))
    fc.set(qn("dirty"), "true")
    for r in runs:
        p.append(r)
    return p


def find_toc_anchor(root, spec):
    """找目录该插在哪里，返回 (锚点元素, 插在锚点之前?)

    toc.after = "first_heading"  -> 插在第一个一级标题之前
    toc.after = "explicit"       -> 插在匹配 toc.anchor 文本的段落之后
    """
    body = root.find(qn("body"))
    if body is None:
        return None, False
    kids = [c for c in body if c.tag == qn("p")]
    # ⚠️ 目录条目段不能当锚点：它们的文本是「一、总体进展情况8」这种
    #    带页码的目录行，会把新目录插到目录区里面去
    kids = [p for p in kids if not _is_toc_entry(p)]

    if spec["toc"].get("after") == "explicit":
        anchor = spec["toc"].get("anchor") or ""
        for p in kids:
            t = "".join(x.text or "" for x in p.iter(qn("t"))).strip()
            if anchor and t.startswith(anchor):
                return p, False       # 插在它之后
        return None, False

    # 默认：第一个一级标题之前
    from .recognize import (RE_L1_CN, RE_L1_ARABIC, RE_APPENDIX,
                            MAX_HEADING_LEN, _outline_lvl, _looks_like_heading)
    for p in kids:
        ppr = p.find(qn("pPr"))
        if ppr is None:
            continue
        ps = ppr.find(qn("pStyle"))
        nm = (ps.get(qn("val")) or "").lower() if ps is not None else ""
        t = "".join(x.text or "" for x in p.iter(qn("t"))).strip()
        if not t:
            continue
        # 已套 heading 1 / outlineLvl=0 —— 自动列表编号的文档标题是「引言」这种光文字
        if nm in ("heading 1", "标题 1", "标题1") or _outline_lvl(p) == 0:
            return p, True
        if RE_L1_CN.match(t) and len(t) <= MAX_HEADING_LEN:
            return p, True
        if RE_APPENDIX.match(t) and len(t) <= MAX_HEADING_LEN:
            return p, True
        if RE_L1_ARABIC.match(t) and len(t) <= MAX_HEADING_LEN \
                and _looks_like_heading(p, t):
            return p, True
    return None, False


def find_existing_toc_titles(root):
    """找出已有的目录/清单标题段（要重建就得先删）。"""
    out = []
    for p in root.iter(qn("p")):
        t = "".join(x.text or "" for x in p.iter(qn("t"))).strip()
        if t in TOC_TITLES and len(t) <= 6:
            out.append(p)
    return out


def remove_existing_toc_blocks(root, spec, log=print):
    """删掉已有的目录块，避免新旧目录并存。

    ⛔ 关键：不能只删「标题 + 紧邻的那个 TOC 域段」——
       一个 TOC 域段落里包含**几十行目录条目**（每行自带
       `HYPERLINK \\l _TocXXXX` + `PAGEREF _TocXXXX` 域），
       删了 TOC 域段但留下条目行，那些条目指向的 `_Toc*` 书签已经不存在了
       → 正文/图表目录里出现「错误！未定义书签。」（踩过）
       所以要把**整块连续条目**一起删。

    两种目录区别对待：
       - 真 TOC **域**目录：域会自刷新，整块删掉安全
       - `WPSOffice手动目录` **静态文字**目录：域刷不了它，删了内容就没了
         → 默认保留并警告，除非 spec.toc.replace_static=true
    """
    removed = 0
    kept_static = 0
    body = root.find(qn("body"))
    if body is None:
        return 0, 0
    kids = list(body)
    i = 0
    while i < len(kids):
        p = kids[i]
        if p.tag != qn("p"):
            i += 1
            continue
        t = "".join(x.text or "" for x in p.iter(qn("t"))).strip()
        if t not in TOC_TITLES or len(t) > 6:
            i += 1
            continue
        # 往后吞掉整块：目录条目（toc 样式 / 目录标题样式 / 含 PAGEREF HYPERLINK 的段）
        j = i + 1
        block = [p]
        while j < len(kids) and kids[j].tag == qn("p"):
            np = kids[j]
            nt = "".join(x.text or "" for x in np.iter(qn("t"))).strip()
            instrs = " ".join(x.text or "" for x in np.iter(qn("instrText")))
            is_entry = ("PAGEREF" in instrs.upper()
                        or "HYPERLINK" in instrs.upper()
                        or "TOC" in instrs.upper())
            is_toc_style = _toc_style(np)
            if is_entry or is_toc_style or (not nt and is_toc_style):
                block.append(np)
                j += 1
                continue
            break
        has_field = any("TOC" in (x.text or "").upper()
                        for bp in block for x in bp.iter(qn("instrText")))
        is_static = not has_field and len(block) > 1
        if is_static and not spec["toc"].get("replace_static"):
            kept_static += 1
            i = j
            continue
        for b in block:
            if b.getparent() is body:
                body.remove(b)
        removed += 1
        i = j
    if kept_static:
        log("[toc] ⚠ 发现 %d 个**静态文字目录**（无 TOC 域，程序刷不了）。\n"
            "       已保留；请在 WPS 里手动更新，或设 spec.toc.replace_static=true "
            "让工具删掉重建。" % kept_static)
    return removed, kept_static


def _toc_style(p):
    """段落的样式名是否属于目录类（toc 1/2/3、目录 1/2/3、table of figures…）。"""
    ppr = p.find(qn("pPr"))
    if ppr is None:
        return False
    ps = ppr.find(qn("pStyle"))
    if ps is None:
        return False
    v = (ps.get(qn("val")) or "").lower()
    return ("toc" in v or "目录" in v or "figures" in v)


def _is_toc_entry(p):
    """这是目录条目行吗（含 PAGEREF/HYPERLINK 域，或用目录样式）。"""
    instrs = " ".join(x.text or "" for x in p.iter(qn("instrText"))).upper()
    if "PAGEREF" in instrs or "HYPERLINK" in instrs:
        return True
    return _toc_style(p)


def insert_toc(root, spec, table, log=print):
    """插正文目录。返回插入的段数。"""
    if not spec.enabled("toc"):
        return 0
    anchor, before = find_toc_anchor(root, spec)
    if anchor is None:
        log("[toc] 找不到锚点，跳过正文目录")
        return 0
    levels = spec["toc"].get("levels", "1-3")
    instr = ' TOC \\o "%s" \\h \\z \\u ' % levels
    title_style = spec.style_of("toc_title") or spec["toc"].get("title_style")
    sid = table.sid(title_style) if title_style else None

    new = []
    if spec["toc"].get("title"):
        new.append(make_title_para(spec["toc"]["title"], sid))
    new.append(make_toc_para(instr))

    parent = anchor.getparent()
    i = list(parent).index(anchor)
    for k, p in enumerate(new):
        parent.insert(i + k if before else i + 1 + k, p)
    log("[toc] 插正文目录 TOC \\o \"%s\"（标题 %r）@ %s"
        % (levels, spec["toc"].get("title"),
           "首个一级标题前" if before else "锚点之后"))
    return len(new)


def insert_list(root, spec, table, kind="图", title=None, log=print):
    """插图目录 / 表目录。kind 决定 \\c 的值，也决定 SEQ 标识符。"""
    key = "figure_list" if kind == "图" else "table_list"
    if not spec.enabled(key):
        return 0
    anchor, before = find_toc_anchor(root, spec)
    if anchor is None:
        log("[toc] 找不到锚点，跳过%s目录" % kind)
        return 0
    instr = ' TOC \\h \\z \\c "%s" ' % kind
    t = title if title is not None else spec[key].get("title")
    title_style = spec[key].get("title_style") or spec.style_of("toc_title")
    sid = table.sid(title_style) if title_style else None

    new = []
    if t:
        new.append(make_title_para(t, sid))
    new.append(make_toc_para(instr))
    parent = anchor.getparent()
    i = list(parent).index(anchor)
    for k, p in enumerate(new):
        parent.insert(i + k if before else i + 1 + k, p)
    log("[toc] 插%s目录 TOC \\c \"%s\"（标题 %r）" % (kind, kind, t))
    return len(new)


def insert_all(root, spec, table, log=print):
    """按 spec 依次插 图目录 / 表目录 / 正文目录。

    先删掉已有的真 TOC 域目录（域会自刷新，删掉安全），再插新的。
    静态文字目录默认保留 —— 域刷不了它，删了内容就没了。
    """
    n = 0
    want = (spec.enabled("toc") or spec.enabled("figure_list")
            or spec.enabled("table_list"))
    if want:
        rm, kept = remove_existing_toc_blocks(root, spec, log=log)
        if rm:
            log("[toc] 删除旧目录块 %d 处" % rm)
    n += insert_list(root, spec, table, "图", log=log)
    n += insert_list(root, spec, table, "表", log=log)
    n += insert_toc(root, spec, table, log=log)
    return n
