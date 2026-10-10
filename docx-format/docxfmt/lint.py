# -*- coding: utf-8 -*-
"""
lint.py —— 正文校验（S5）。

**只出改动事项单（CRS），绝不改文档。** 发现的每条问题落成 CRS 的一行，
由调用方勾选决定改不改。不要在这里顺手改"明显的错字"。

范围（spec.lint）：
    scope=body      只扫正文：段落(paragraph) + 标题(heading)
    exclude         formula（公式段）/ appendix（附录）

检查项：
    P1 半角标点夹在中文里        → 应改全角
    P2 标点后多余空格            → 去空格
    P3 中文↔数字/英文 之间空格    → 去空格（spec.spacing.normalize_cjk_latin）
    P4 公式与中文之间的空格       → 去空格
    P5 叠字 / 连续标点            → 多余字
    P6 缺字模式（第 个 / 其中为 / 空括号 …）
    P7 相邻重复 run（多录公式/字）
    P8 术语同形异写
    P9 图表引用落空（图 N / 表 N 不存在）
    P10 百分点表述不自洽（由 A 至 B 提高 C，C ≈ B-A）
"""
import os
import re
from collections import Counter, defaultdict

from . import qn
from .crs import CrsBuilder
from .docxio import Docx, read_style_names, unpack
from .recognize import RE_APPENDIX, parse

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

CJK = r"\u4e00-\u9fff\u3000-\u303f\uff00-\uffef"
RE_HALF_PUNCT_CN = re.compile(
    r"[%s][,;:!?](?=[^0-9\s])" % CJK)
RE_PUNCT_SPACE = re.compile(r"[，。；：、）】》”][ \u3000]+(?=[^\s])")
RE_CJK_SPACE = re.compile(r"([%s])([ \u3000]+)([0-9A-Za-z])" % CJK)
RE_LATIN_SPACE = re.compile(r"([0-9A-Za-z])([ \u3000]+)([%s])" % CJK)
# 「图 1」「表 2-1」「式 (3)」「附图 3」里的空格是**体例的一部分**，任何位置都不许去。
# ⛔ 实测教训：正文交叉引用的 `如图 1所示` 里，`图 1` 同时是 REF 域的**缓存文本**；
#    去掉空格会让缓存与题注（仍写 `图 1`）不一致 → 一按 F9 就被 Word 改回去，
#    并让编号治理的 verify_all 第 6 项（REF 缓存 == 语义重算值）直接 FAIL 95 处。
RE_KEEP_SPACE = re.compile(r"(?:附图|附表|图|表|式|公式)\s+\d")
RE_DUP_CHAR = re.compile(r"(的的|了了|是是|在在|和和|与与|地地|得得|着着)")
RE_DUP_PUNCT = re.compile(r"([，。；：])\1+")
RE_MIX_PUNCT = re.compile(r"[，。；：]{2,}")
RE_EMPTY_PAREN = re.compile(r"（\s*）|\(\s*\)|【\s*】")
RE_MISSING = [
    (re.compile(r"第\s*[个页章节步条节]"), "「第 __ 个/章/节」漏了数字"),
    (re.compile(r"其中\s*[为是]?[，,]?\s*为"), "「其中为」漏了变量名"),
    (re.compile(r"[见图表]\s*(?:所示|如下)"), "「图/表 所示」漏了编号"),
    (re.compile(r"约\s*%"), "「约 %」漏了数值"),
    (re.compile(r"[如见][公式]\s*$"), "「如式/见公式」漏了编号"),
]
RE_REF_NUM = re.compile(r"[图表]\s*(\d{1,3})(?!\d)")
RE_PCT = re.compile(
    r"由\s*(\d+(?:\.\d+)?)\s*%?\s*(?:提高|提升|增加|增长|下降|降低|减少)"
    r"[至到]\s*(\d+(?:\.\d+)?)\s*%?[^。；]{0,40}?"
    r"(提高|提升|增加|增长|下降|降低|减少)了?\s*(\d+(?:\.\d+)?)\s*%")

# 已知同形异写组（可扩展）
VARIANT_GROUPS = [
    ["Grounding DINO", "GroundingDINO"],
    ["Deformable DETR", "Deformable-DETR", "DeformableDETR"],
    ["YOLOv10", "YOLO v10", "YOLO v8", "YOLOv8"],
    ["Qwen3.5-397B", "Qwen3.5", "Qwen 3.5"],
    ["Segment Anything Model", "SAM"],
]


def _visible(p):
    """段落可见文本，**只取 w:t**（公式 m:t 不算正文，天然排除）。"""
    return "".join(t.text or "" for t in p.iter(qn("t")))


def _runs_text(p):
    """按文档顺序返回段落里的 run（**含空文本 run**）。

    ⛔ 不能过滤空文本 run：公式 run（`<m:oMath>`，无 `w:t`）夹在两个「、」之间，
       过滤后两个「、」会显得"相邻"，误报「相邻重复 run」。
    """
    out = []
    for r in p.iter(qn("r")):
        ts = r.findall(qn("t"))
        out.append((r, "".join(t.text or "" for t in ts)))
    return out


def _in_appendix(node):
    kind = node.sec_kind
    if kind == "appendix":
        return True
    # 附录内的段落在树里挂在 appendix 标题下 -> 沿 parent 上溯
    n = node
    while n is not None:
        if getattr(n, "sec_kind", None) == "appendix":
            return True
        n = getattr(n, "parent", None)
    return False


def analyze(docx, spec):
    """返回 (crs_dict)。不写盘。"""
    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_lint")
    unpack(docx, workdir)
    doc = Docx(workdir)
    sid2name = read_style_names(workdir)
    root, _index, _cover = parse(doc.root, sid2name,
                                 skip_keywords=spec["numbering"]
                                 .get("skip_keywords"))
    scan = bool((spec.get("spacing") or {}).get("normalize_cjk_latin", True))

    b = CrsBuilder(docx, sources=("lint",),
                   baseline_note="lint 出单；目标指纹变化即作废")

    # 收集正文节点（paragraph / heading），带段号
    nodes = []

    def walk(n):
        if n.type != "document" and n.type != "empty":
            nodes.append(n)
        for c in n.children:
            walk(c)
    walk(root)
    nodes = [n for n in nodes if n.meta.get("idx") is not None]

    body = [n for n in nodes if n.type in ("paragraph", "heading")]
    # 附录排除
    body = [n for n in body if not _in_appendix(n)]

    # 题注编号集合（供 P9）
    cap_nums = {"图": set(), "表": set()}
    for n in nodes:
        if n.type == "caption":
            m = re.match(r"^\s*([图表])\s*(\d{1,3})", n.text or "")
            if m:
                cap_nums.setdefault(m.group(1), set()).add(int(m.group(2)))

    term_forms = defaultdict(Counter)

    for n in body:
        p = n.ref
        sec = _section_of(n)
        txt = _visible(p)
        if not txt.strip():
            continue

        # P1 半角标点（before/after 精确到字符，供 revise 直接替换）
        for m in RE_HALF_PUNCT_CN.finditer(txt):
            s = txt[m.start():m.end()]
            after = s[0] + _full(s[1])
            b.add("A", "标点", _w(sec, n.meta["idx"], s),
                  s, after, "中文之间应使用全角标点",
                  risk="low", executor="text_replace")
        # P2 标点后空格
        for m in RE_PUNCT_SPACE.finditer(txt):
            s = txt[m.start():min(len(txt), m.end() + 1)]
            after = s[0] + s[1:].lstrip(" \u3000")
            b.add("B", "体例", _w(sec, n.meta["idx"], s),
                  s, after, "标点与后续文字之间不应有空格",
                  risk="low", executor="text_replace")
        # P3 中文↔拉丁/数字 之间的空格
        if scan:
            keep = [(m.start(), m.end()) for m in RE_KEEP_SPACE.finditer(txt)]
            for rx in (RE_CJK_SPACE, RE_LATIN_SPACE):
                for m in rx.finditer(txt):
                    # 落在「图 1 / 表 2-1 / 式 (3)」体例区里的一律放过
                    if any(a <= m.start() < z for a, z in keep):
                        continue
                    s = m.group(0)
                    after = m.group(1) + m.group(3)
                    b.add("B", "空格", _w(sec, n.meta["idx"], s),
                          s, after,
                          "中文与数字/英文之间不留空格（本文体例）",
                          risk="low", executor="text_replace")
        # P4 公式与中文之间的空格（run 相邻 oMath）
        _check_formula_gap(p, n, sec, b)
        # P5 叠字 / 连续标点
        for m in RE_DUP_CHAR.finditer(txt):
            b.add("S", "多字",
                  _w(sec, n.meta["idx"], txt[max(0, m.start() - 8):m.end() + 8]),
                  m.group(0), m.group(0)[0],
                  "叠字，疑似多录一个字", risk="low", executor="text_replace")
        for m in RE_DUP_PUNCT.finditer(txt):
            b.add("A", "标点",
                  _w(sec, n.meta["idx"], txt[max(0, m.start() - 8):m.end() + 8]),
                  m.group(0), m.group(0)[0],
                  "同一标点连打", risk="low", executor="text_replace")
        for m in RE_MIX_PUNCT.finditer(txt):
            if RE_DUP_PUNCT.fullmatch(m.group(0)):
                continue
            b.add("C", "标点",
                  _w(sec, n.meta["idx"], txt[max(0, m.start() - 8):m.end() + 8]),
                  m.group(0), "",
                  "连续标点，留哪个需人工判断 → 交人工",
                  risk="low", executor="manual")
        # P6 缺字模式（**只出清单，绝不自动改** —— 补什么字只有作者知道）
        for rx, why in RE_MISSING:
            for m in rx.finditer(txt):
                b.add("S", "缺字",
                      _w(sec, n.meta["idx"],
                         txt[max(0, m.start() - 10):m.end() + 10]),
                      m.group(0), "",
                      why + "（需人工补全）",
                      risk="low", executor="manual")
        for m in RE_EMPTY_PAREN.finditer(txt):
            b.add("S", "缺字",
                  _w(sec, n.meta["idx"], txt[max(0, m.start() - 10):m.end() + 10]),
                  m.group(0), "",
                  "空括号，疑似漏填内容（需人工补全）",
                  risk="low", executor="manual")
        # P7 相邻重复 run（**必须含空文本 run 判邻接**，否则公式占位会假报）
        rs = _runs_text(p)
        for i in range(len(rs) - 1):
            t1, t2 = rs[i][1], rs[i + 1][1]
            if len(t1) >= 2 and t1 == t2 and \
                    rs[i][0].getparent() is rs[i + 1][0].getparent():
                b.add("S", "多录",
                      _w(sec, n.meta["idx"], t1[:20]),
                      t1 + t2, t1,
                      "相邻两个 run 内容完全相同，疑似重复录入",
                      risk="low", executor="run_dedup")
        # P8 术语
        for g in VARIANT_GROUPS:
            for v in g:
                c = txt.count(v)
                if c:
                    term_forms[tuple(g)][v] += c

    # P8 汇总
    for g, cnt in term_forms.items():
        present = [v for v in g if cnt.get(v)]
        if len(present) > 1:
            best = max(present, key=lambda v: cnt[v])
            b.add("A", "术语", {"section": "全文", "para": 0,
                              "anchor": " / ".join(present),
                              "count": sum(cnt.values())},
                  " / ".join("%s(%d)" % (v, cnt[v]) for v in present),
                  best,
                  "同一概念多种写法，需统一（取出现最多的写法）；全文级改动，交人工/批量",
                  risk="low", executor="manual")

    # P9 引用落空
    for n in body:
        txt = _visible(n.ref)
        for m in RE_REF_NUM.finditer(txt):
            kind, num = m.group(0)[0], int(m.group(1))
            if num not in cap_nums.get(kind, set()) and cap_nums.get(kind):
                b.add("S", "引用",
                      _w(_section_of(n), n.meta["idx"],
                         txt[max(0, m.start() - 10):m.end() + 10]),
                      m.group(0), "",
                      "正文引用的%s编号 %d 在题注里不存在（最大 %d）"
                      % (kind, num, max(cap_nums[kind])),
                      risk="low", executor="manual")
    # P10 百分点自洽（注意单位：数值可能是小数(0-1)也可能是百分数）
    for n in body:
        txt = _visible(n.ref)
        for m in RE_PCT.finditer(txt):
            if re.search(r"较|相比|相对|比", m.group(0)):
                continue          # 「较 X 提高」是另一个基准，不是两数之差
            va, vb, c = float(m.group(1)), float(m.group(2)), float(m.group(4))
            d = abs(vb - va)
            if min(abs(d - c), abs(d * 100.0 - c)) > 0.05:
                b.add("A", "数据",
                      _w(_section_of(n), n.meta["idx"], m.group(0)[:40]),
                      m.group(0), "",
                      "由 %.4g 到 %.4g 的差是 %.4g（= %.4g 个百分点），"
                      "与所述 %.4g 不符" % (va, vb, d, d * 100, c),
                      risk="low", executor="manual")

    import shutil
    shutil.rmtree(workdir, ignore_errors=True)
    return b.build()


def _section_of(node):
    p = node
    sec = None
    while p is not None:
        if p.type == "heading":
            sec = p.text
        p = getattr(p, "parent", None)
    return (re.sub(r"\s+", " ", sec or "")[:40] or "（无章节）")


def _w(sec, para, anchor):
    return {"section": sec, "para": para, "anchor": anchor,
            "count": 1}


_FULL = {",": "，", ";": "；", ":": "：", "!": "！", "?": "？"}


def _full(ch):
    return _FULL.get(ch, ch)


def _check_formula_gap(p, n, sec, b):
    """run 紧邻 <m:oMath> 且以空白结尾/开头 -> 公式与中文之间有空格。"""
    kids = [c for c in p if c.tag in (qn("r"), "{%s}oMath" % M_NS)]
    for i, c in enumerate(kids):
        if c.tag != "{%s}oMath" % M_NS:
            continue
        for side, chk in ((-1, "endswith"), (1, "startswith")):
            j = i + side
            if j < 0 or j >= len(kids) or kids[j].tag != qn("r"):
                continue
            ts = kids[j].findall(qn("t"))
            if not ts:
                continue
            s = "".join(t.text or "" for t in ts)
            if getattr(s, chk)(" ") or getattr(s, chk)("\u3000"):
                b.add("B", "空格",
                      _w(sec, n.meta["idx"], s[:20]), s, s.strip(),
                      "公式与中文之间不应有空格",
                      risk="low", executor="text_replace")
                break
