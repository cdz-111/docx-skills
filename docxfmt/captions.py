# -*- coding: utf-8 -*-
"""
captions.py —— 题注自动生成 + 交叉引用 + 公式编号。

为什么必须做成域
    真实文档里编号是**手写静态文本**，而且既重复（各章各编一遍）又跳号（缺号）、
    甚至还留着占位符 `图x`。一旦统一编号，原来手写的「表 2」就全错位。
    做成 SEQ + REF 之后，编号怎么变引用都由 Word 自动跟着变。

三条核心经验
1. **定位引用要按「作者原编号精确匹配」，不能按「第 N 个」**。
   优先级：本章节原编号==N → 全文原编号==N（多个取最近）→ 退化「第 N 个」
   → 占位符（`图x`）/方位词（`上表`）走就近。
2. **书签只包「标签+编号」，不包说明文字**，否则 REF 会把整句题注引出来。
3. **引用正则必须带编号**（`[如见](图|表)\\s*([xX]|\\d+)`），
   否则「如**图**像、语音」会被误判成引用。
"""
import copy
import re

from lxml import etree

from . import XSPACE, qn
from . import M
from .docxio import (bookmark, el, field_runs, first_rpr, get_ppr,
                     has_field, has_drawing, in_table, mk_run, new_bookmark_id,
                     p_text, set_style)

RE_CAP_HEAD = re.compile(r"^\s*(图|表)(?:\s*([xX]|\d+)\s*|\s+)")
RE_REF_NUM = re.compile(r"[如见](图|表)\s*([xX]|\d+)")
RE_REF_REL = re.compile(r"[上下](图|表)")
RE_REF_BARE = re.compile(
    r"(?<![\d\-.])(图|表)\s*(\d+)(?![\d\-])(?!\s*[个种项名处条台套人年月日])")
# 公式编号：(1) / （1） / 式(1) /  (12)
RE_FORMULA = re.compile(r"^\s*[(（]\s*(\d{1,3})\s*[)）]\s*$")
RE_FORMULA_REF = re.compile(r"公式\s*[(（]\s*(\d{1,3})\s*[)）]")


# ------------------------------------------------------------------ 工具
def _set_t(t, s):
    t.text = s
    if s != s.strip():
        t.set(XSPACE, "preserve")
    elif XSPACE in t.attrib:
        del t.attrib[XSPACE]


def _nodes(root):
    out = []
    def rec(n):
        out.append(n)
        for c in n.children:
            rec(c)
    rec(root)
    return out


def section_map(root):
    """一次遍历算出每个节点的「最近标题」「最近一级标题」。

    ⚠️ 别写成「每个段落做一次全树递归」—— 那是 O(n^2)，大文档会卡死。
    """
    near, chap1, sec_kind = {}, {}, {}
    def rec(n, h, h1, sk):
        if n.id:
            near[n.id] = h.id if h is not None else "#root"
            chap1[n.id] = h1.id if h1 is not None else "#root"
            sec_kind[n.id] = n.sec_kind if n.sec_kind else sk
        nh = n if n.type == "heading" else h
        nh1 = n if (n.type == "heading" and n.level == 1) else h1
        nsk = n.sec_kind if n.sec_kind else sk
        for c in n.children:
            rec(c, nh, nh1, nsk)
    rec(root, None, None, None)
    return near, chap1, sec_kind


def split_caption_paras(doc_root):
    """图片与题注文字挤在**同一段**时，拆成「图片段」+「题注段」。

    ⚠️ 拆完段落序号会变，所以必须早于 parse()。
    """
    n = 0
    for p in list(doc_root.iter(qn("p"))):
        t = p_text(p).strip()
        if not RE_CAP_HEAD.match(t):
            continue
        runs = p.findall(qn("r"))
        if not runs:
            continue
        draw_at = [i for i, r in enumerate(runs)
                   if r.find(".//" + qn("drawing")) is not None]
        if not draw_at:
            continue
        last = max(draw_at)
        tail = runs[last + 1:]
        if not tail or any(r.find(".//" + qn("drawing")) is not None
                           for r in tail):
            continue                      # 图文交错，交给人工
        ppr = p.find(qn("pPr"))
        newp = etree.Element(qn("p"))
        if ppr is not None:
            nppr = copy.deepcopy(ppr)
            for e in nppr.findall(qn("sectPr")):   # sectPr 留在原段
                nppr.remove(e)
            newp.append(nppr)
        for r in tail:
            newp.append(r)
        parent = p.getparent()
        parent.insert(list(parent).index(p) + 1, newp)
        n += 1
    return n


# ------------------------------------------------------------------ 题注
class Caption:
    __slots__ = ("idx", "kind", "seq", "raw_num", "desc", "section", "chapter",
                 "bookmark", "bk_id", "node", "in_table", "scheme", "sec_kind")

    def __init__(self, idx, kind, desc, raw_num=None, section=None,
                 chapter=None, bookmark=None, bk_id=None, node=None, seq=0,
                 in_table=False, scheme="global", sec_kind=None):
        self.idx = idx
        self.kind = kind
        self.seq = seq
        self.scheme = scheme           # 实际用的编号方案（可能与 spec 不同）
        self.sec_kind = sec_kind       # chapter / front / appendix
        self.raw_num = raw_num
        self.desc = desc
        self.section = section
        self.chapter = chapter
        self.bookmark = bookmark
        self.bk_id = bk_id
        self.node = node
        self.in_table = in_table

    def __repr__(self):
        return "<%s%d %r %s>" % (self.kind, self.seq, self.desc[:16],
                                 self.bookmark)


def _existing_ref_bk(p):
    """题注段上已有的 `_Ref*` 书签要保留 —— 正文里已经有 REF 指过来了，
    删掉就把别人的引用搞断。`_Toc*` 只是目录锚点，可以丢。"""
    for e in p.findall(qn("bookmarkStart")):
        nm = e.get(qn("name")) or ""
        if nm.startswith("_Ref"):
            return nm, e.get(qn("id"))
    return None, None


def collect_captions(root, paras, kinds=("图", "表")):
    """收集题注：树内 caption 节点 + 补扫「插图下方第一段非空文字」。

    图片下方的第一段文字常常就是题注，所以除了树里判为 caption 的，
    还要扫**每个非表格内插图**的下方 1~2 段。
    """
    cap_ids = {}
    for n in _nodes(root):
        if n.type == "caption" and n.id:
            cap_ids[n.id] = n

    extra = 0
    for i, p in enumerate(paras):
        if in_table(p) or not has_drawing(p):
            continue
        for k in (1, 2):
            j = i + k
            if j >= len(paras) or in_table(paras[j]):
                break
            t = p_text(paras[j]).strip()
            if not t:
                continue
            if RE_CAP_HEAD.match(t) and ("p%d" % j) not in cap_ids:
                cap_ids["p%d" % j] = None
                extra += 1
            break

    caps = []
    for pid in sorted(cap_ids, key=lambda x: int(x[1:])):
        i = int(pid[1:])
        p = paras[i]
        t = p_text(p).strip()
        m = RE_CAP_HEAD.match(t)
        if not m:
            continue
        kind = m.group(1)
        if kind not in kinds:
            continue
        raw = m.group(2)
        bk, bk_id = _existing_ref_bk(p)
        caps.append(Caption(i, kind, t[m.end():].strip(),
                            raw_num=int(raw) if (raw and raw.isdigit()) else None,
                            bookmark=bk, bk_id=bk_id, node=cap_ids[pid],
                            in_table=in_table(p)))
    return caps, extra


def number_captions(caps, scheme="global"):
    """编号 + 分配书签名（书签名一律 ASCII，中文书签名在 WPS 上有翻车风险）。

    ⛔ 分章编号有个陷阱：引言/摘要/附录里的题注**不占章号**，
       它们前面没有 `SEQ 章` 锚点，`SEQ 章 \\c` 取不到值，
       图表目录里会显示成未定义书签。
       所以 front/appendix 区的题注强制退化成**全局纯序号**。
    """
    g = {}
    per = {}
    for c in caps:
        g[c.kind] = g.get(c.kind, 0) + 1
        no_chapter = getattr(c, "sec_kind", None) in ("front", "appendix")
        if scheme == "chapter" and not no_chapter and c.chapter != "#root":
            k = per.setdefault(c.chapter, {})
            k[c.kind] = k.get(c.kind, 0) + 1
            c.seq = k[c.kind]
            c.scheme = "chapter"
            if not c.bookmark:
                c.bookmark = "_%s_%s_%d" % (
                    "fig" if c.kind == "图" else "tab", c.chapter, c.seq)
        else:
            c.seq = g[c.kind]
            c.scheme = "global"
            if not c.bookmark:
                c.bookmark = "_%s%d" % (
                    "fig" if c.kind == "图" else "tab", c.seq)
    return caps


def adopt_orphan_bookmarks(caps, paras):
    """空段落上孤零零挂着的 `_Ref*` 书签，作者本意是引用紧随其后的题注。

    有的文档里就有这种：空段上挂着 `_Ref*` 书签，
    Word 重算时解析不到 -> 正文直接显示未定义书签。
    让下一条题注沿用这个名字，原有交叉引用就自动接上了，一个 REF 都不用改。
    """
    by_idx = {c.idx: c for c in caps}
    order = sorted(by_idx)
    adopted = []
    for i, p in enumerate(paras):
        if p_text(p).strip():
            continue
        for b in list(p.findall(qn("bookmarkStart"))):
            nm = b.get(qn("name")) or ""
            if not nm.startswith("_Ref"):
                continue
            nxt = next((j for j in order if j > i), None)
            if nxt is None:
                continue
            c = by_idx[nxt]
            if c.bookmark:
                continue                 # 已有名字，不动（否则整段会被引进来）
            bid = b.get(qn("id"))
            for be in list(p.findall(qn("bookmarkEnd"))):
                if be.get(qn("id")) == bid:
                    p.remove(be)
            p.remove(b)
            c.bookmark = nm
            adopted.append((i, nxt, nm))
    return adopted


def write_caption(p, cap, table, style_name="题注", scheme="global", sep="-",
                  bk_id=None):
    """重写题注段：书签(标签+编号) + 说明文字。

        bookmarkStart _fig1
        r "图 " | SEQ 域
        bookmarkEnd
        r " 说明文字"
    """
    rpr = first_rpr(p)
    for c in list(p):
        if c.tag != qn("pPr"):
            p.remove(c)
    sid = table.sid(style_name) if style_name else None
    if sid:
        set_style(p, sid)

    bs, be = bookmark(cap.bookmark, bk_id)
    p.append(bs)
    p.append(mk_run("%s " % cap.kind, rpr=rpr))

    if scheme == "chapter" and getattr(cap, "scheme", "global") == "chapter":
        # 只有正文章节的题注才带章号。引言/附录的题注前面没有 `SEQ 章` 锚点，
        # `SEQ 章 \c` 取不到值，图表目录里会显示成未定义书签
        for r in field_runs(" SEQ 章 \\c ", cached="1", rpr=rpr):
            p.append(r)
        p.append(mk_run(sep, rpr=rpr))
        p.extend(field_runs(" SEQ %s \\* ARABIC \\s 1 " % cap.kind,
                            cached=str(cap.seq), rpr=rpr))
    else:
        p.extend(field_runs(" SEQ %s \\* ARABIC " % cap.kind,
                            cached=str(cap.seq), rpr=rpr))
    p.append(be)
    if cap.desc:
        p.append(mk_run(" " + cap.desc, rpr=rpr))


def insert_chapter_seq(root):
    """分章编号：每个一级标题开头插一个隐藏的 `SEQ 章`（只递增不显示）。

    题注里的 `SEQ 章 \\c` 取的就是它当前的值。
    """
    n = 0
    for node in _nodes(root):
        if node.type != "heading" or node.level != 1 or node.ref is None:
            continue
        if node.sec_kind == "front":
            continue                    # 引言/摘要不占章号
        p = node.ref
        rpr = first_rpr(p)
        runs = field_runs(" SEQ 章 \\h ", cached="1", rpr=rpr)
        # 结果 run 设为隐藏
        for r in runs:
            ts = r.findall(qn("t"))
            if ts and (ts[0].text or "") == "1":
                rp = r.find(qn("rPr"))
                if rp is None:
                    rp = etree.Element(qn("rPr"))
                    r.insert(0, rp)
                if rp.find(qn("vanish")) is None:
                    rp.insert(0, el("vanish"))
        pos = 1 if p.find(qn("pPr")) is not None else 0
        for k, r in enumerate(runs):
            p.insert(pos + k, r)
        n += 1
    return n


# ------------------------------------------------------------------ 公式编号
def convert_formula_numbers(root, paras, table, style_name="公式编号",
                            scheme="global", sep="-", log=print):
    """把硬编码的公式编号改成 SEQ 域。兼容两种文档形态。

    形态 A：公式与编号**同段**（pPr / r[tab] / oMath / r[tab] / r[(30)]）
        → 只删末尾那个编号 run，其余一律不碰
    形态 B：编号**独立成段**（整段文本就是 `(30)`）
        → 整段重写（段里没有别的东西可丢）

    ⛔ 铁律：形态 A 绝不能整段重写 —— oMath 就在同一段里，
       整段重写会把公式本体删掉。
    """
    if scheme == "none":
        return 0, 0
    nodes = {}
    for n in _nodes(root):
        if n.id:
            nodes[n.id] = n
    cnt = 0
    for i, p in enumerate(paras):
        node = nodes.get("p%d" % i)
        if node is None or node.type not in ("paragraph", "formula_number"):
            continue
        runs = p.findall(qn("r"))
        if not runs:
            continue

        # ---- 形态 B：整段就是编号
        full = p_text(p).strip()
        m_b = RE_FORMULA.match(full)
        has_omml = p.find(".//" + qn("oMath", M)) is not None
        if m_b and not has_omml:
            rpr = first_rpr(p)
            for c in list(p):
                if c.tag != qn("pPr"):
                    p.remove(c)
            sid = table.sid(style_name) if style_name else None
            if sid:
                set_style(p, sid)
            p.append(mk_run("(", rpr=rpr))
            if scheme == "chapter":
                p.extend(field_runs(" SEQ 章 \\c ", cached="1", rpr=rpr))
                p.append(mk_run(sep, rpr=rpr))
                p.extend(field_runs(" SEQ 式 \\* ARABIC \\s 1 ",
                                    cached=m_b.group(1), rpr=rpr))
            else:
                p.extend(field_runs(" SEQ 式 \\* ARABIC ",
                                    cached=m_b.group(1), rpr=rpr))
            p.append(mk_run(")", rpr=rpr))
            cnt += 1
            continue

        # ---- 形态 A：只动末尾那个编号 run
        last = runs[-1]
        ts = last.findall(qn("t"))
        if not ts:
            continue
        m = RE_FORMULA.match("".join(t.text or "" for t in ts))
        if not m:
            continue
        rpr = first_rpr(p)
        p.remove(last)                    # 只删编号 run，oMath 与 tab 都留着
        sid = table.sid(style_name) if style_name else None
        if sid:
            ppr = get_ppr(p)
            ps = ppr.find(qn("pStyle"))
            if ps is None:
                ps = etree.Element(qn("pStyle"))
                ppr.insert(0, ps)
            ps.set(qn("val"), sid)
        p.append(mk_run("(", rpr=rpr))
        if scheme == "chapter":
            p.extend(field_runs(" SEQ 章 \\c ", cached="1", rpr=rpr))
            p.append(mk_run(sep, rpr=rpr))
            p.extend(field_runs(" SEQ 式 \\* ARABIC \\s 1 ", cached=m.group(1),
                                rpr=rpr))
        else:
            p.extend(field_runs(" SEQ 式 \\* ARABIC ", cached=m.group(1),
                                rpr=rpr))
        p.append(mk_run(")", rpr=rpr))
        cnt += 1
    log("[formula] 公式编号转 SEQ 域 %d 处" % cnt)
    return cnt, cnt


# ------------------------------------------------------------------ 交叉引用
def _replace_range_with_field(p, start, end, bookmark_name, display, rpr=None):
    """把段落文本 [start,end) 换成 REF 域，**支持跨 run**。

    真实文档里「如表 1 所示」经常被 Word 拆成好几个 run，
    只处理单 run 会静默漏掉一批引用。

    做法：**不删 run**（run 里可能藏着图片或已有域），
    只把首 run 截成 before、尾 run 截成 after、中间清空，再插域。
    """
    segs = []
    pos = 0
    for t in p.iter(qn("t")):
        s = t.text or ""
        segs.append((t, pos, len(s)))
        pos += len(s)
    if start < 0 or end > pos or start >= end:
        return False
    first = last = None
    for k, (t, p0, ln) in enumerate(segs):
        # ln > 0：空 w:t（len 0）不能当选首尾，否则切片越界
        if ln > 0 and p0 < end and p0 + ln > start:
            if first is None:
                first = k
            last = k
    if first is None:
        return False
    t0, q0, _ = segs[first]
    tl, ql, _ = segs[last]
    before = (t0.text or "")[:start - q0]
    after = (tl.text or "")[end - ql:]        # 切片，不是取单字符

    tail = []
    if first == last:
        _set_t(t0, before)
        if after:
            tail.append(mk_run(after, rpr=rpr))
    else:
        _set_t(t0, before)
        for k in range(first + 1, last):
            _set_t(segs[k][0], "")
        _set_t(tl, after)

    runs = field_runs(" REF %s \\h " % bookmark_name, cached=display, rpr=rpr)
    runs += tail
    anchor = t0.getparent()
    while anchor is not None and anchor.tag != qn("r"):
        anchor = anchor.getparent()
    host = p if anchor is None else anchor.getparent()
    if host is None:
        return False
    i = list(host).index(anchor) if anchor is not None else 0
    for k, r in enumerate(runs):
        host.insert(i + 1 + k, r)
    return True


def link_refs(root, paras, caps, near, log=print):
    """扫描正文里的图/表引用，换成 REF 交叉引用域。"""
    cap_idx = {c.idx for c in caps}
    by_sec, by_kind = {}, {}
    for c in caps:
        by_sec.setdefault(c.section, {}).setdefault(c.kind, []).append(c)
        by_kind.setdefault(c.kind, []).append(c)

    def pick(kind, num, i):
        lst = by_sec.get(near.get("p%d" % i, "#root"), {}).get(kind, [])
        cands = [c for c in lst if c.raw_num == num]
        if len(cands) == 1:
            return cands[0]
        if len(cands) > 1:
            return min(cands, key=lambda c: abs(c.idx - i))
        cands = [c for c in by_kind.get(kind, []) if c.raw_num == num]
        if cands:
            return min(cands, key=lambda c: abs(c.idx - i))
        if 1 <= num <= len(lst):            # 退化：第 N 个
            return lst[num - 1]
        g = by_kind.get(kind, [])
        if 1 <= num <= len(g):
            return g[num - 1]
        return None

    def nearest(kind, i, direction):
        if direction < 0:
            c = [x for x in by_kind.get(kind, []) if x.idx < i]
            return c[-1] if c else None
        c = [x for x in by_kind.get(kind, []) if x.idx > i]
        return c[0] if c else None

    done, miss, skipped = [], [], []
    for i, p in enumerate(paras):
        if i in cap_idx:
            continue
        t = p_text(p)
        if not t or has_field(p, "REF"):
            continue
        hits = []
        # 只替换「图/表 + 编号」，保留前缀的「如/见」→ 输出才是「如图 1 所示」
        for m in RE_REF_NUM.finditer(t):
            num = m.group(2)
            hits.append((m.start(1), m.end(), m.group(1),
                         int(num) if num.isdigit() else None, 0))
        for m in RE_REF_REL.finditer(t):
            hits.append((m.start(), m.end(), m.group(1), None,
                         -1 if m.group(0)[0] == "上" else 1))
        for m in RE_REF_BARE.finditer(t):
            s, e = m.span()
            if any(h[0] <= s < h[1] for h in hits):
                continue
            hits.append((s, e, m.group(1), int(m.group(2)), 0))
        if not hits:
            continue
        hits = sorted(set(hits), key=lambda h: (h[0], h[1]))
        rpr = first_rpr(p)
        # 从后往前替换，避免位置漂移
        for (s, e, kind, num, rel) in sorted(hits, key=lambda h: (h[0], h[1]),
                                             reverse=True):
            if rel:
                cap = nearest(kind, i, rel)
                how = "就近(%s)" % ("上" if rel < 0 else "下")
            elif num is None:
                cap = nearest(kind, i, -1) or nearest(kind, i, 1)
                how = "就近(占位符)"
            else:
                cap = pick(kind, num, i)
                how = ("原编号命中" if (cap and cap.raw_num == num)
                       else "退化第%d个" % num)
            if not cap:
                miss.append((i, t[s:e]))
                continue
            disp = "%s %d" % (kind, cap.seq)
            if _replace_range_with_field(p, s, e, cap.bookmark, disp, rpr):
                done.append((i, t[s:e], cap.bookmark, disp, how))
            else:
                skipped.append((i, t[s:e]))
    log("[xref] 交叉引用 %d 处；定位不到 %d %s；跨 run 跳过 %d"
        % (len(done), len(miss), miss[:3], len(skipped)))
    from collections import Counter
    log("[xref] 定位方式：" + "  ".join(
        "%s=%d" % kv for kv in Counter(d[4] for d in done).most_common()))
    return done, miss, skipped


# ------------------------------------------------------------------ 入口
def process(root, doc_root, paras, table, spec, log=print):
    """题注 + 交叉引用全流程。必须早于 document.xml 落盘。"""
    kinds = tuple(spec["captions"].get("kinds") or ("图", "表"))
    scheme = spec["numbering"].get("caption", "global")
    sep = spec.sep
    cap_style = spec.style_of("caption")

    near, chap1, sec_kind = section_map(root)
    caps, extra = collect_captions(root, paras, kinds)
    for c in caps:
        pid = "p%d" % c.idx
        c.section = near.get(pid, "#root")
        c.chapter = chap1.get(pid, "#root")
        c.sec_kind = sec_kind.get(pid)

    adopted = adopt_orphan_bookmarks(caps, paras)
    number_captions(caps, scheme)

    if scheme == "chapter":
        n_ch = insert_chapter_seq(root)
        log("[caption] 分章编号：插入隐藏 SEQ 章 锚点 %d 个" % n_ch)

    nid = new_bookmark_id(doc_root)
    for c in caps:
        if c.bk_id is None:
            c.bk_id = nid
            nid += 1
    for c in caps:
        write_caption(paras[c.idx], c, table, cap_style, scheme, sep, c.bk_id)
        if c.node is not None:
            c.node.style = cap_style
            c.node.text = "%s %d %s" % (c.kind, c.seq, c.desc)

    log("[caption] 题注 %d 条（图 %d / 表 %d），插图下方补扫 %d，认领孤儿书签 %d"
        % (len(caps), sum(1 for c in caps if c.kind == "图"),
           sum(1 for c in caps if c.kind == "表"), extra, len(adopted)))
    if spec.enabled("xref"):
        link_refs(root, paras, caps, near, log=log)
    return caps
