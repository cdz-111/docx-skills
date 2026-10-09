# -*- coding: utf-8 -*-
"""
insert_captions.py —— S3：给「该有题注但没有」的图 / 表补插题注。

**只出改动事项单（CRS），绝不改文档。** 执行走 `--apply`（结构类通道）。

判定：
    图 —— 非表格内、含 `w:drawing` 的段落。往**下**看 1~2 个非空段，
          以 `图 N` / `表 N` 开头即为已有题注。找不到 → 缺。
    表 —— 每个 `w:tbl`。往**上**看紧邻的段落（表标题在表上方），再看紧邻下方
          的第一段。以 `表 N` 开头即为已有题注。找不到 → 缺。

编号形态：由 spec 的 `numbering.caption`（global `1` / chapter `1-1`）与
`numbering.sep` 决定，**给出的只是建议编号**；最终编号以 `run` 阶段的
SEQ 域重算为准（本模块只补「题注段 + 样式」，不塞域）。

补出来的题注**说明文字留空**（写什么只有作者知道），在批注里点名要补。
"""
import os
import re

from lxml import etree

from . import qn
from .applykit import Applier, _Skip
from .crs import CrsBuilder
from .docxio import (Docx, get_ppr, has_drawing, in_table, p_text,
                     read_style_names, set_style, unpack)
from .docxio import mk_run
from .recognize import RE_CAP_HEAD, mark_sections, parse

CAP_STYLE = "题注"
RE_KIND_NUM = re.compile(r"^\s*([图表])\s*(?:(\d+)\s*[-\u2010\u2011\u2012\u2013]\s*)?(\d+)?")


def _next_nonempty(paras, in_tbl_flags, start, span=3):
    for k in range(1, span + 1):
        j = start + k
        if j >= len(paras):
            return None
        if in_tbl_flags[j]:
            return None
        t = p_text(paras[j]).strip()
        if t:
            return j
    return None


def analyze(docx, spec, log=print):
    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_cap")
    unpack(docx, workdir)
    doc = Docx(workdir)
    root = doc.root
    sid2name = read_style_names(workdir)
    paras = doc.paras()
    intbl = [in_table(p) for p in paras]
    texts = [p_text(p).strip() for p in paras]

    scheme = (spec.get("numbering") or {}).get("caption", "global")
    sep = (spec.get("numbering") or {}).get("sep", "-")

    b = CrsBuilder(docx, sources=("insert-captions",),
                   baseline_note="题注补插单；目标指纹变化即作废")

    # ---- 现有题注编号（求建议编号的起点）
    have = {"图": [], "表": []}
    for t in texts:
        m = RE_CAP_HEAD.match(t)
        if m:
            mm = re.search(r"(\d+)", t)
            if mm:
                have.setdefault(m.group(1), []).append(int(mm.group(1)))
    nextno = {k: (max(v) + 1 if v else 1) for k, v in have.items()}

    # ---- 章节号（分章编号用）
    tree, _i, _c = parse(root, sid2name)
    mark_sections(tree, (spec.get("numbering") or {}).get("skip_keywords"))
    sec_of = {}

    def walk(n, cur):
        if n.type == "heading":
            cur = n
        sec_of[n.meta.get("idx")] = cur
        for c in n.children:
            walk(c, cur)
    walk(tree, None)

    ch_no = {}
    n_ch = 0
    for n in [x for x in _all(tree)]:
        if n.type == "heading" and n.level == 1:
            if n.sec_kind == "front":
                continue
            n_ch += 1
            ch_no[n.meta.get("idx")] = n_ch

    def _chapno(j):
        """段 j 所属的**一级正文章节**号；不在正文章节内返回 None。"""
        h1 = sec_of.get(j)
        while h1 is not None and not (h1.type == "heading" and h1.level == 1):
            h1 = h1.parent
        return ch_no.get(h1.meta.get("idx")) if h1 is not None else None

    # chapter 方案的序号是「**章内**序号」，不是全局序号。
    # 先按章统计**已有**题注数，作为各章的起点（缺图/表按文档序在其上递增）。
    # 只在 chapter 模式下做，global 路径连这段统计都不跑 → 行为零改动。
    ch_have = {}
    if scheme == "chapter":
        for j, t in enumerate(texts):
            m = RE_CAP_HEAD.match(t)
            if not m or m.group(1) not in ("图", "表") or not re.search(r"\d", t):
                continue
            cno = _chapno(j)
            if cno:
                ch_have[(m.group(1), cno)] = ch_have.get((m.group(1), cno), 0) + 1

    def suggest(kind, idx):
        if scheme == "chapter":
            cno = _chapno(idx)
            if cno:
                seq = ch_have.get((kind, cno), 0) + 1
                ch_have[(kind, cno)] = seq
                return "%s %d%s%d" % (kind, cno, sep, seq)
        return "%s %d" % (kind, nextno[kind])

    # ---- 图：非表格内、含 drawing
    n_fig = 0
    for i, p in enumerate(paras):
        if intbl[i] or not has_drawing(p):
            continue
        t = texts[i]
        if RE_CAP_HEAD.match(t):        # 图文同段（题注就在同段里）
            continue
        j = _next_nonempty(paras, intbl, i)
        if j is not None and RE_CAP_HEAD.match(texts[j]):
            continue
        label = suggest("图", i)
        b.add("A", "缺题注",
              {"section": _sec(texts, i), "para": i,
               "anchor": ("图片段：" + (t[:30] or "（无文字）")), "count": 1},
              "（无题注）", label,
              "图片下方 %s 段内没有题注段" % 2,
              risk="mid", executor="insert_caption",
              notes="说明文字需作者补全；编号最终以 run 阶段重算为准")
        nextno["图"] += 1
        n_fig += 1

    # ---- 表：每个 w:tbl 看上方紧邻段
    n_tab = 0
    n_cont = 0
    for ti, tbl in enumerate(root.iter(qn("tbl"))):
        prev = tbl.getprevious()
        # 续表：紧贴在上一个表格之后 —— 同一逻辑表被拆成多个 w:tbl，
        # 题注只挂在该逻辑表的第一个分段之前，续段不该再出题注。
        # （ti 仍是真实 w:tbl 索引，continue 不影响执行器用的 where.tbl_index。）
        if prev is not None and prev.tag == qn("tbl"):
            n_cont += 1
            continue
        ok = False
        if prev is not None and prev.tag == qn("p"):
            if RE_CAP_HEAD.match(p_text(prev).strip()) and \
                    p_text(prev).strip()[:1] == "表":
                ok = True
        if not ok:
            nxt = tbl.getnext()
            while nxt is not None and nxt.tag == qn("p") and \
                    not p_text(nxt).strip():
                nxt = nxt.getnext()
            if nxt is not None and nxt.tag == qn("p") and \
                    p_text(nxt).strip()[:1] == "表" and \
                    RE_CAP_HEAD.match(p_text(nxt).strip()):
                ok = True
        if ok:
            continue
        pi = paras.index(prev) if (prev is not None and prev.tag == qn("p")
                                   and prev in paras) else 0
        label = suggest("表", pi)
        b.add("A", "缺题注",
              {"section": _sec(texts, pi), "para": pi,
               "anchor": "第 %d 个表（表格上方紧邻段不是题注）" % (ti + 1),
               "count": 1, "tbl_index": ti},
              "（无题注）", label,
              "表格上下都没有「表 N」题注段",
              risk="mid", executor="insert_caption",
              notes="说明文字需作者补全；编号最终以 run 阶段重算为准")
        nextno["表"] += 1
        n_tab += 1

    import shutil
    shutil.rmtree(workdir, ignore_errors=True)
    log("[题注] 图缺题注 %d 个；表缺题注 %d 个（跳过续表 %d 个；方案 %s，分隔符 %r）"
        % (n_fig, n_tab, n_cont, scheme, sep))
    return b.build()


def _all(tree):
    out = []

    def rec(n):
        out.append(n)
        for c in n.children:
            rec(c)
    rec(tree)
    return out


def _sec(texts, i):
    for k in range(i, -1, -1):
        t = texts[k]
        m = re.match(r"^\s*(\d{1,2}(?:\.\d{1,2})*)\s+\S", t)
        if m and len(t) <= 60:
            return "%s %s" % (m.group(1), t[m.end() - 1:][:24].strip())
    return "（无章节）"


# ------------------------------------------------------------------ 执行
class CaptionApplier(Applier):
    """补插题注段。只增 `w:p`，其余计数不变。"""

    allowed_delta = {"p": "*", "tbl": 0, "drawing": 0, "bookmarkStart": 0,
                     "bookmarkEnd": 0, "fldChar": 0, "instrText": 0}
    supported = {"insert_caption"}

    def __init__(self, docx, spec=None, author="docx-format 排版"):
        super().__init__(docx, author)
        self.spec = spec or {}

    def apply_one(self, doc, paras, it):
        if it.get("executor") == "manual":
            raise _Skip("该项需人工处理")
        if it.get("executor") != "insert_caption":
            raise _Skip("本执行器不处理 executor=%s" % it.get("executor"))
        w = it.get("where") or {}
        after = it.get("after_override") or it.get("after") or ""
        if not after:
            raise _Skip("缺 after（题注文字）")
        ti = w.get("tbl_index")
        if ti is not None:                       # 表题注：插在**表格之前**
            tbls = list(doc.root.iter(qn("tbl")))
            if ti >= len(tbls):
                raise _Skip("表格号 %s 越界" % ti)
            return self._insert_before(doc, tbls[ti], after, it)
        i = w.get("para")
        if i is None or i >= len(paras):
            raise _Skip("段号 %s 越界" % i)
        return self._insert_after(doc, paras[i], after, it)

    def _make_caption_p(self, doc, text):
        from .docxio import el  # noqa: F401
        newp = etree.Element(qn("p"))
        etree.SubElement(newp, qn("pPr"))
        style = CAP_STYLE
        try:
            from .styles import StyleTable
            st = StyleTable(doc.sty.getroot())
            sid = st.sid(style)
            if sid:
                set_style(newp, sid)
        except Exception:                        # noqa: BLE001
            pass
        newp.append(mk_run(text))
        return newp

    def _insert_before(self, doc, ref, text, it):
        """插在 ref（表格）**紧前方**。

        为什么不能沿用「插在前一个兄弟段之后」：两张表连续时
        （实测表 2 的前一个兄弟就是表 1），前一个兄弟不是 `w:p`，
        旧写法会直接报「表格上方没有可插入位置」。插在表格前则永远成立。
        对「前一个兄弟确实是段落」的情形，两种写法落点完全相同。
        """
        newp = self._make_caption_p(doc, text)
        parent = ref.getparent()
        parent.insert(list(parent).index(ref), newp)
        return 1

    def _insert_after(self, doc, ref_p, text, it):
        newp = self._make_caption_p(doc, text)
        parent = ref_p.getparent()
        parent.insert(list(parent).index(ref_p) + 1, newp)
        return 1


def apply(docx, crs_path, approved_path=None, out=None, spec=None, dry=False,
          log=print):
    ap = CaptionApplier(docx, spec)
    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_capc")
    unpack(docx, workdir)
    ap._before_counts = ap.counts(
        etree.parse(os.path.join(workdir, "word", "document.xml")).getroot())
    import shutil
    shutil.rmtree(workdir, ignore_errors=True)
    return ap.run(crs_path, approved_path, out, dry, log)
