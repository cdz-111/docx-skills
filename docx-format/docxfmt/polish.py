# -*- coding: utf-8 -*-
"""
polish.py —— 排版（S4）：老板的六条硬要求。

**只出改动事项单（CRS），绝不改文档。** 执行走 `--apply`（结构类通道）。

六条要求 → 实现：
    ① 文档 >100 页时一级标题分页       静态（lxml）+ 页数（渲染级）
    ② 数字/英文/公式与中文之间不留空格   **由 lint P3/P4 统一负责**（见下）
    ③ 正文每页不出现 ≥4 行空白          渲染级（COM 量页坐标）
    ④ 不出现单字加标点独占一行          渲染级（末字符 x 坐标）
    ⑤ 题注与对应图/表同页               渲染级（题注段与目标段页码）
    ⑥ 表格美化：数据表三线、文本表全框   静态（lxml 单元格文本量 + 边框来源）

② 的归属：空格规则与 `lint` 的 P3/P4 **完全是同一条**（同一组正则、同一套
例外）。为避免同一处改动出两张互相打架的单子，**S4 直接采用 lint 的输出**，
本模块不重复检测；报告里会明确提示。

━━━ 两个必须记住的实测结论（都踩过）━━━

**结论一：不要按段号映射。**
WPS 的 `Paragraphs` 计数与 lxml 的 `w:p` 计数**并不相等**
（实测 4917 vs 4383）。差在两处：
  - 图形对象段：WPS 的 `Range.Text` 返回 `/`，lxml 侧是空串
  - 表格行结束标记：WPS 会额外计入
所以任何「WPS 段号 = lxml 段号」的假设都会错位。
→ ③④⑤ 一律在 **WPS 段号空间内自洽**解决，对外只报「页码 + 文本锚点」。

**结论二：表格边框的真实来源是 `tcBorders`，不是 `tblBorders`。**
本文档 52 个表里，边框几乎全写在**单元格级** `w:tcPr/w:tcBorders`
（表 24、表 41 连 `tblBorders` 都没有）。而且 `tblStyle` 是 WPS 重写过的
**数字 styleId**（41 / 29 / 30），从 ID 反查语义名不可靠。
→ 判定必须**汇总单元格级边框**；执行时必须**先清掉 `tcBorders`**，
   否则单元格级优先，改 `tblBorders` 等于没改。
"""
import os
import re
import shutil
from collections import Counter, defaultdict

from lxml import etree

from . import qn
from .applykit import Applier, backup, md5, write_receipt
from .crs import CrsBuilder
from .docxio import Docx, get_ppr, pack, read_style_names, unpack
from .recognize import parse

CJK = r"\u4e00-\u9fff\u3400-\u4dbf"
RE_CJK_SPACE = re.compile(r"([%s])([ \u3000]+)([0-9A-Za-z])" % CJK)
RE_LATIN_SPACE = re.compile(r"([0-9A-Za-z])([ \u3000]+)([%s])" % CJK)
RE_CAP_INTEXT = re.compile(r"^\s*([图表])\s*(\d{1,3})")
RE_TOC_LINE = re.compile(r"[\t\u2026]\s*\d+\s*$")
# 「像章节标题」的段落：第X章/节/篇、1.2 编号、一、二、三、附录/参考文献…
# ⚠️ 必须含「第X章」：漏了这一支 → 大量**章末留白**被误报成正文空白（实测 8 页误报 1 页）。
RE_NEW_CHAPTER = re.compile(
    r"^\s*(?:第\s*[一二三四五六七八九十百零〇\d]+\s*[章节篇部讲]|"
    r"\d{1,2}(?:\.\d+)*[\.\s、]|"
    r"[一二三四五六七八九十]+\s*[、.]|"
    r"附\s*录|参考文献|致\s*谢|摘\s*要|目\s*录|引\s*言|绪\s*论|"
    r"结\s*论|讨\s*论|前\s*言)")
TAIL_PUNCT = "。；：，！？、）」》”…"

DEF_PAGE = {"w": 595.3, "h": 841.9, "top": 72.0, "bottom": 72.0,
            "left": 90.0, "right": 90.0}
TW = 20.0


def _norm(s):
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", s or "")


# ================================================================== 上下文
class Ctx:
    """一次解包，供所有检查复用（避免 64MB 文档被反复解包）。"""

    def __init__(self, docx, spec):
        self.src = docx
        self.spec = spec
        self.workdir = os.path.join(os.path.dirname(os.path.abspath(docx)),
                                    "_polish")
        unpack(docx, self.workdir)
        self.doc = Docx(self.workdir)
        self.root = self.doc.root
        self.paras = self.doc.paras()
        self.sid2name = read_style_names(self.workdir)
        self.geom = page_geom(self.root)
        self.fs = body_font_size(spec, self.workdir)

    def close(self):
        shutil.rmtree(self.workdir, ignore_errors=True)


def page_geom(root):
    """从 sectPr 静态读页面几何（pt）。

    ⚠️ 只能用这条路：本机 WPS **不支持** `PageSetup.*`（返回 9999999.0 哨兵）。
    """
    body = root.find(qn("body"))
    sect = None
    if body is not None:
        direct = [s for s in body if s.tag == qn("sectPr")]
        if direct:
            sect = direct[-1]
        else:
            for s in body.iter(qn("sectPr")):
                sect = s
    g = dict(DEF_PAGE)
    if sect is None:
        return g
    sz = sect.find(qn("pgSz"))
    if sz is not None:
        for k in ("w", "h"):
            v = sz.get(qn(k))
            if v and v.isdigit() and int(v) > 1400:
                g[k] = int(v) / TW
    mar = sect.find(qn("pgMar"))
    if mar is not None:
        for k in ("top", "bottom", "left", "right"):
            v = mar.get(qn(k))
            if v and re.match(r"^-?\d+$", v) and abs(int(v)) < 20000:
                g[k] = int(v) / TW
    return g


def body_font_size(spec, workdir):
    g = (spec.get("styles") or {}).get("geometry") or {}
    for nm in ("正文", "Normal"):
        v = (g.get(nm) or {}).get("size")
        if v:
            try:
                return float(v)
            except (TypeError, ValueError):
                pass
    try:
        st = etree.parse(os.path.join(workdir, "word", "styles.xml")).getroot()
        for s in st.iter(qn("style")):
            nm = s.find(qn("name"))
            if nm is not None and (nm.get(qn("val")) or "").lower() in (
                    "normal", "正文"):
                sz = s.find(qn("rPr") + "/" + qn("sz"))
                if sz is not None and (sz.get(qn("val")) or "").isdigit():
                    return int(sz.get(qn("val"))) / 2.0
    except Exception:                                # noqa: BLE001
        pass
    return 12.0


def measure_line_height(pages, default=21.0):
    """实测行高：同页相邻段 y 差的**众数**。

    ⚠️ 中位数不可用：段前/段后间距会把一部分 y 差抬高，中位数被带到
       「行高 + 段距」（实测 29.1pt），空白判据过宽 → 大量漏报。
    ⚠️ 低分位也不可用：同页里夹着表格行、浮动题注，会产生 5~10pt 的
       小差值，p20 取到噪声（实测退化成 12.1pt ≈ 字号），判据过严。
    → 正文行距是**统一**的，所以「单倍行高的整数倍」里**出现次数最多**的
      那个值就是真实行高。取众数最稳。
    """
    ps = [p for p in (pages.get("paras") or [])
          if p.get("y") is not None and p.get("page")]
    diffs = []
    for a, b in zip(ps, ps[1:]):
        if a.get("page") == b.get("page"):
            d = b["y"] - a["y"]
            if 8 < d < 80:
                diffs.append(d)
    if not diffs:
        return default
    c = Counter(round(d * 2) / 2.0 for d in diffs)
    top = c.most_common(3)
    # 众数若明显是「两倍行高」，取它的 1/2 更接近单行
    best = top[0][0]
    if len(top) > 1 and abs(top[1][0] - best / 2.0) < 1.5 and top[1][1] > top[0][1] * 0.5:
        best = top[1][0]
    return best


# ================================================================== ①
def check_pagebreak(ctx, pages, b, log=print):
    conf = ctx.spec.get("pagebreak") or {}
    min_pages = int(conf.get("min_pages", 100) or 100)
    levels = conf.get("before_heading") or [1]
    if 1 not in levels:
        log("[①] spec 未要求一级标题分页（before_heading=%s）→ 跳过" % levels)
        return 0
    total = (pages or {}).get("pages") or 0
    if not total:
        log("[①] 未量页，无法判断是否 >%d 页 → 跳过" % min_pages)
        return 0
    if total <= min_pages:
        log("[①] 文档 %d 页 ≤ %d 页 → 按老板要求不分页" % (total, min_pages))
        return 0

    # ⛔ 必须用 ctx.root 这同一棵树：另起 etree.parse 会造出第二棵树，
    #    paras.index(p) 永远失配（元素 == 是身份比较），段号会全变 0。
    tree, _idx, _cov = parse(
        ctx.root, ctx.sid2name,
        skip_keywords=(ctx.spec.get("numbering") or {}).get("skip_keywords"))
    h1 = []

    def walk(n):
        if n.type == "heading" and n.level == 1 and n.ref is not None:
            h1.append(n)
        for c in n.children:
            walk(c)
    walk(tree)

    pos = {id(p): i for i, p in enumerate(ctx.paras)}
    n = 0
    for nd in h1:
        p = nd.ref
        i = pos.get(id(p))
        ppr = p.find(qn("pPr"))
        if ppr is not None and ppr.find(qn("pageBreakBefore")) is not None:
            continue
        if i and i > 0:
            if any(br.get(qn("type")) == "page"
                   for br in ctx.paras[i - 1].iter(qn("br"))):
                continue
        b.add("A", "分页",
              {"section": nd.text[:40], "para": i if i is not None else 0,
               "anchor": nd.text[:40], "count": 1},
              "（无分页）", "一级标题另起一页",
              "文档共 %d 页（>%d），一级标题应分页" % (total, min_pages),
              risk="mid", executor="pagebreak")
        n += 1
    log("[①] 一级标题 %d 个，需补分页 %d 个（共 %d 页 > %d）"
        % (len(h1), n, total, min_pages))
    return n


# ================================================================== ③
def _has_content(t):
    """段落是否有"看得见"的内容。

    ⛔ 不能只用 `_norm()` 判：图片段落的 `Range.Text` 是 `/`，
       `_norm()` 会把它清成空串，于是「整页最后一段是图片」被误判成
       「最后一段是空段」→ ye 偏小 → 假报一大片空白。
    """
    return bool((t or "").strip().strip("\r\n\x07\x0b\x0c\u2001\u3000 "))


def check_blank(pages, geom, lh, b, log=print):
    """③ 正文每页不出现 ≥4 行空白。章末留白不算。"""
    if not pages:
        log("[③] 无量页数据 → 跳过")
        return 0
    usable_bottom = geom["h"] - geom["bottom"]
    per = defaultdict(list)
    for p in pages.get("paras") or []:
        if p.get("ye") is None or p.get("page") is None:
            continue
        per[p["page"]].append(p)
    if not per:
        return 0
    by_i = {p["i"]: p for p in (pages.get("paras") or [])}
    maxpage = max(per)
    n = skip_chapter = skip_empty = skip_tbl = skip_cap = 0
    blanks = []
    for pg in sorted(per):
        if pg >= maxpage:
            continue                                 # 末页留白正常
        items = [x for x in per[pg] if _has_content(x.get("t"))]
        if not items:
            skip_empty += 1
            continue
        last = max(items, key=lambda x: x["ye"])
        lt = last.get("t") or ""
        blank = usable_bottom - last["ye"]
        if blank < 4 * lh:
            continue
        # 表格页：末段是单元格内容（WPS 用 \x07 标单元格结束）——
        # 断表造成的留白不是"正文空白"，须人工看表格是否该拆
        if "\x07" in lt:
            skip_tbl += 1
            continue
        # 图表后留白：末段是题注，说明该页以图/表收尾，属正常
        if re.match(r"^\s*[图表]\s*\d", lt.strip()):
            skip_cap += 1
            continue
        # 章末留白：往后找 3 段（跨过空段），遇到分页符或新章节标题即算章末
        chapter_end = False
        for k in range(1, 4):
            q = by_i.get(last["i"] + k)
            if q is None:
                chapter_end = True                # 到文档末尾了
                break
            qt = q.get("t") or ""
            if "\x0c" in qt:
                chapter_end = True
                break
            if _has_content(qt):
                if RE_NEW_CHAPTER.match(qt.strip()) and len(qt.strip()) < 60:
                    chapter_end = True
                break                          # 遇到第一个有内容的段就停
        if chapter_end:
            skip_chapter += 1
            continue
        blanks.append(blank)
        b.add("B", "空白",
              {"section": "（第 %d 页）" % pg, "para": 0,
               "anchor": lt[:40], "count": 1,
               "wps_para": last["i"], "page": pg},
              "第 %d 页底部空约 %.0fpt（≈%.1f 行）" % (pg, blank, blank / lh),
              "扩/缩本页正文，或调整本页图片高度",
              "老板要求：正文每页不出现 4 行以上空白",
              risk="mid", executor="manual")
        n += 1
    if blanks:
        blanks.sort()
        log("[③] 扫 %d 页，空白超 4 行 %d 处"
            "（忽略 章末 %d / 表格页 %d / 题注结尾 %d / 空页 %d）"
            "；空白量 中位 %.0fpt / 最大 %.0fpt；行高 %.1fpt"
            % (len(per), n, skip_chapter, skip_tbl, skip_cap, skip_empty,
               blanks[len(blanks) // 2], blanks[-1], lh))
    else:
        log("[③] 扫 %d 页，无超 4 行空白（忽略 章末 %d / 表格页 %d / "
            "题注结尾 %d / 空页 %d；行高 %.1fpt）"
            % (len(per), skip_chapter, skip_tbl, skip_cap, skip_empty, lh))
    return n


# ================================================================== ④
def check_orphan(pages, geom, fs, b, log=print):
    """④ 不出现单字加标点独占一行。

    判据：段末是标点 + 末行 > 首行 + **末字符 x 紧贴左边界**
    （说明末行只有「单字 + 标点」，甚至只有标点）。
    """
    if not pages:
        log("[④] 无量页数据 → 跳过")
        return 0
    left = geom["left"]
    thr = 1.6 * fs
    n = 0
    for p in pages.get("paras") or []:
        le, ls, xe = p.get("ln_end"), p.get("ln"), p.get("xe")
        if le is None or ls is None or xe is None or not p.get("t"):
            continue
        if le <= ls:
            continue
        tail = (p.get("tail") or "").strip()
        if not tail or tail[-1] not in TAIL_PUNCT:
            continue
        if len(_norm(p.get("t"))) < 12:              # 太短的段（题注/标题/编号）
            continue
        if xe - left > thr:
            continue
        b.add("B", "孤字",
              {"section": "（第 %s 页）" % p.get("page"), "para": 0,
               "anchor": (p.get("t") or "")[:40], "count": 1,
               "wps_para": p["i"], "page": p.get("page")},
              "末行只有「%s」（末字符 x=%.0fpt，左边界 %.0fpt）"
              % (tail[-2:], xe, left),
              "微调本段段前/段后、字距，或调整相邻图片宽度",
              "老板要求：不出现单字加标点独占一行",
              risk="mid", executor="manual")
        n += 1
    log("[④] 扫 %d 段，孤字行候选 %d 处（阈值 末字符x-左边界 ≤ %.1fpt）"
        % (len(pages.get("paras") or []), n, thr))
    return n


# ================================================================== ⑤
def check_caption_page(pages, b, log=print):
    """⑤ 题注与对应图/表同页。图题注在下方（往上看图形段），表题注在上方。"""
    if not pages:
        log("[⑤] 无量页数据 → 跳过")
        return 0
    ps = pages.get("paras") or []
    by_i = {p["i"]: p for p in ps}
    n = 0
    for p in ps:
        raw = (p.get("t") or "").strip()
        if "\x0c" in raw:
            raw = raw.replace("\x0c", "").strip()
        if not RE_CAP_INTEXT.match(raw):
            continue
        # ⛔ 目录条目长得像题注（「图 1 xxx \t 12」）—— 必须排除，
        #    否则会把目录里的「图 N」当成正文题注去比页码。
        if RE_TOC_LINE.search(raw) or "\t" in raw:
            continue
        if len(_norm(raw)) < 3:
            continue
        kind = raw[0]
        target = None
        if kind == "图":
            for k in range(1, 4):
                q = by_i.get(p["i"] - k)
                if q is None:
                    break
                qt = q.get("t") or ""
                if "/" in qt:
                    target = q
                    break
                if _norm(qt):
                    break
        else:
            target = by_i.get(p["i"] + 1)
        if target is None or target.get("page") is None:
            continue
        if target.get("page") == p.get("page"):
            continue
        b.add("A", "题注",
              {"section": "（题注在第 %s 页）" % p.get("page"), "para": 0,
               "anchor": raw[:40], "count": 1,
               "wps_para": p["i"], "page": p.get("page")},
              "题注在第 %s 页，%s 在第 %s 页"
              % (p.get("page"), "图" if kind == "图" else "表",
                 target.get("page")),
              "把题注与对应图/表调到同一页"
              "（给题注段加「与下段同页」keep_with_next，或扩缩正文）",
              "老板要求：题注与对应部分同页",
              risk="mid", executor="manual")
        n += 1
    log("[⑤] 题注跨页 %d 处" % n)
    return n


# ================================================================== ⑥
def _read_borders(bs):
    """从一组 w:*Borders 判定形态。返回 three_line / full_grid / None。"""
    if bs is None:
        return None
    got = {}
    for e in bs:
        nm = etree.QName(e).localname
        val = e.get(qn("val"))
        sz = e.get(qn("sz")) or "0"
        got[nm] = (val, int(sz) if sz.isdigit() else 0)

    def on(k):
        v = got.get(k)
        return bool(v) and v[0] not in ("nil", "none") and v[1] > 0
    if not got:
        return None
    h = on("insideH")
    v = on("insideV")
    if on("top") and on("bottom") and not on("left") and not on("right") \
            and not v:
        return "three_line"
    if on("left") and on("right") and (h or v):
        return "full_grid"
    return None


def table_border_kind(tbl):
    """判表格现有边框形态。

    ⛔ 边框主要写在**单元格级** `tcPr/tcBorders`（本机实测），
       `tblBorders` 有时干脆没有。只看 `tblBorders` 会把满框表判成"无边框"。
    """
    k = _read_borders(tbl.find(qn("tblPr") + "/" + qn("tblBorders")))
    if k:
        return k, "表级"
    c = Counter()
    for tc in tbl.iter(qn("tc")):
        kk = _read_borders(tc.find(qn("tcPr") + "/" + qn("tcBorders")))
        c[kk] += 1
    if not c:
        return None, "无"
    k, cnt = c.most_common(1)[0]
    return k, "单元格级(%d/%d)" % (cnt, sum(c.values()))


def check_tables(ctx, b, log=print):
    conf = ctx.spec.get("tables") or {}
    thr = int(conf.get("text_table_min_chars", 60) or 60)
    tbls = list(ctx.root.iter(qn("tbl")))
    stats, n, unknown = [], 0, 0
    for ti, tbl in enumerate(tbls):
        cells = []
        for tc in tbl.iter(qn("tc")):
            s = "".join(t.text or "" for t in tc.iter(qn("t")))
            cells.append(len(s.strip()))
        if not cells:
            continue
        avg = sum(cells) / float(len(cells))
        mx = max(cells)
        kind, where = table_border_kind(tbl)
        want = "full_grid" if (avg >= thr or mx >= 120) else "three_line"
        if kind == want:
            continue
        if kind is None:
            unknown += 1
        stats.append((ti, len(cells), avg, mx, kind, where, want))
    for ti, nc, avg, mx, kind, where, want in stats:
        label = {"three_line": "三线表", "full_grid": "全框表"}[want]
        b.add("B", "表格",
              {"section": "第 %d 个表" % (ti + 1), "para": 0,
               "anchor": "表 %d（%d 格，均 %.1f 字，最长 %d 字）"
                         % (ti + 1, nc, avg, mx), "count": 1,
               "tbl_index": ti},
              "现为 %s（边框写在 %s）"
              % ({"three_line": "三线表", "full_grid": "全框表",
                  None: "无/其他"}[kind], where),
              label,
              "老板要求：数据表用三线表、长文本表用全框"
              "（本表均 %.1f 字 %s 阈值 %d）；"
              "执行时会同步清掉单元格级 tcBorders，否则改表级边框不生效"
              % (avg, "≥" if avg >= thr else "<", thr),
              risk="mid", executor="table_border")
        n += 1
    log("[⑥] 表格 %d 个，需改边框 %d 个（其中原判不出 %d 个；文本表阈值 %d 字）"
        % (len(tbls), n, unknown, thr))
    return n


# ================================================================== 入口
def analyze(docx, spec, pages=None, pdf_result=None, log=print):
    ctx = Ctx(docx, spec)
    lh = measure_line_height(pages) if pages else 14.0
    b = CrsBuilder(docx, sources=("polish",),
                   baseline_note="polish 出单；目标指纹变化即作废")
    if pages:
        b.baseline["render"] = {
            "engine": (pages.get("engine") or {}).get("name"),
            "is_wps": (pages.get("engine") or {}).get("is_wps"),
            "pages": pages.get("pages"),
            "line_height_pt": round(lh, 2),
            "page_geom_pt": {k: round(v, 1) for k, v in ctx.geom.items()},
            "body_font_pt": ctx.fs,
        }
    log("=" * 66)
    log("排版体检（只出清单，不改文档）：%s" % os.path.basename(docx))
    log("  页面几何 %s" % {k: round(v, 1) for k, v in ctx.geom.items()})
    log("  正文字号 %.1fpt　实测行高(p20) %.1fpt" % (ctx.fs, lh))
    if not pages:
        log("  ⚠️ 没有量页数据：③④⑤ 无法检测"
            "（跑 paginate.py，或用 --pages 传进量页 json）")
    log("-" * 66)
    c1 = check_pagebreak(ctx, pages, b, log)
    c3 = check_blank(pages, ctx.geom, lh, b, log)
    c4 = check_orphan(pages, ctx.geom, ctx.fs, b, log)
    c5 = check_caption_page(pages, b, log)
    c6 = check_tables(ctx, b, log)
    if (spec.get("spacing") or {}).get("normalize_cjk_latin", True):
        log("[②] 空格规则归 lint P3/P4 统一负责，本模块不重复出单")
    log("-" * 66)
    log("合计：①%d  ③%d  ④%d  ⑤%d  ⑥%d" % (c1, c3, c4, c5, c6))
    if pdf_result:
        # 渲染级**第二判据**：对同一文档量导出的 PDF 文本层，与 COM 判据对账。
        # 只作证据进 baseline，不参与任何自动决定（门禁不变）。
        from . import pdf_measure as PM
        log("-" * 66)
        log("渲染级第二判据（PDF 文本层）：%d 页；孤字行 %d，大空白 %d，低填充页 %d"
            % (pdf_result.get("pages") or 0,
               (pdf_result.get("summary") or {}).get("orphan_lines", 0),
               (pdf_result.get("summary") or {}).get("big_gaps", 0),
               (pdf_result.get("summary") or {}).get("low_fill_pages", 0)))
        cc = PM.compare(pages, pdf_result, geom=ctx.geom, fs=ctx.fs, log=log)
        # 行高标定对账：③ 的阈值挂在 COM 推出的「行高」上，若它与 PDF 文本层实测的
        # 真实行距差得远（>1.5×），③ 的判据会系统性失真（偏松→漏报 或 偏紧→误报）。
        pdf_pitch = pdf_result.get("pitch") or 0
        calib = None
        if pdf_pitch and lh:
            ratio = lh / pdf_pitch
            calib = {"com_line_height_pt": round(lh, 2),
                     "pdf_pitch_pt": round(pdf_pitch, 2),
                     "ratio": round(ratio, 2)}
            if ratio > 1.5 or ratio < 0.67:
                log("[②③④ 交叉验证] ⚠ 行高标定不一致：COM 行高 %.1fpt vs "
                    "PDF 实测行距 %.1fpt（%.2f×）→ ③ 的「4 行空白」阈值"
                    "可能失真，建议以 PDF 行距复核" % (lh, pdf_pitch, ratio))
        b.baseline["pdf_judgment"] = {
            "pdf": pdf_result.get("pdf"),
            "pdf_md5": pdf_result.get("pdf_md5"),
            "engine": (pdf_result.get("engine") or {}).get("name"),
            "pages": pdf_result.get("pages"),
            "pitch_pt": pdf_result.get("pitch"),
            "summary": pdf_result.get("summary"),
            "cross_check": cc,
            "line_height_calibration": calib,
        }
    ctx.close()
    return b.build()


# ================================================================== 执行
_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap",
    "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
_TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
    "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
    "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
    "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
]


def _insert_ordered(parent, child, order):
    name = etree.QName(child).localname
    want = order.index(name) if name in order else len(order)
    for i, c in enumerate(parent):
        cn = etree.QName(c).localname
        ci = order.index(cn) if cn in order else len(order)
        if ci > want:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def set_pagebreak_before(p):
    ppr = get_ppr(p)
    if ppr.find(qn("pageBreakBefore")) is None:
        _insert_ordered(ppr, etree.Element(qn("pageBreakBefore")), _PPR_ORDER)
    return 1


def set_table_border(tbl, kind):
    """改表级 tblBorders，**并清掉单元格级 tcBorders**（否则不生效）。"""
    pr = tbl.find(qn("tblPr"))
    if pr is None:
        pr = etree.Element(qn("tblPr"))
        tbl.insert(0, pr)
    old = pr.find(qn("tblBorders"))
    if old is not None:
        pr.remove(old)

    def bd(name, val, sz):
        e = etree.Element(qn(name))
        e.set(qn("val"), val)
        if val != "nil":
            e.set(qn("sz"), str(sz))
            e.set(qn("space"), "0")
            e.set(qn("color"), "000000")
        return e
    bs = etree.Element(qn("tblBorders"))
    if kind == "three_line":
        bs.append(bd("top", "single", 12))           # 1.5pt
        bs.append(bd("left", "nil", 0))
        bs.append(bd("bottom", "single", 12))
        bs.append(bd("right", "nil", 0))
        bs.append(bd("insideH", "single", 6))        # 0.75pt
        bs.append(bd("insideV", "nil", 0))
    else:
        for nm in ("top", "left", "bottom", "right", "insideH", "insideV"):
            bs.append(bd(nm, "single", 4))           # 0.5pt
    _insert_ordered(pr, bs, _TBLPR_ORDER)

    # ⛔ 必做：单元格级边框优先级高于表级，不清掉等于没改
    removed = 0
    for tc in tbl.iter(qn("tc")):
        tcpr = tc.find(qn("tcPr"))
        if tcpr is None:
            continue
        for tb in list(tcpr.findall(qn("tcBorders"))):
            tcpr.remove(tb)
            removed += 1
    return removed


class PolishApplier(Applier):
    """执行 ①（一级标题分页）与 ⑥（表格边框）。字符类改动不走这里。"""

    allowed_delta = {"p": 0, "tbl": 0, "drawing": 0, "bookmarkStart": 0,
                     "bookmarkEnd": 0, "fldChar": 0, "instrText": 0}
    supported = {"pagebreak", "table_border"}

    def __init__(self, docx, spec=None, author="docx-format 排版"):
        super().__init__(docx, author)
        self.spec = spec or {}

    def apply_one(self, doc, paras, it):
        from .applykit import _Skip
        ex = it.get("executor")
        if ex == "manual":
            raise _Skip("该项需人工处理（executor=manual）")
        if ex == "pagebreak":
            i = (it.get("where") or {}).get("para")
            if i is None or i >= len(paras):
                raise _Skip("段号 %s 越界" % i)
            return set_pagebreak_before(paras[i])
        if ex == "table_border":
            ti = (it.get("where") or {}).get("tbl_index")
            tbls = list(doc.root.iter(qn("tbl")))
            if ti is None or ti >= len(tbls):
                raise _Skip("表格号 %s 越界" % ti)
            want = "full_grid" if "全框" in (it.get("after") or "") \
                else "three_line"
            set_table_border(tbls[ti], want)
            return 1
        raise _Skip("本执行器不处理 executor=%s" % ex)


def apply(docx, crs_path, approved_path=None, out=None, spec=None, dry=False,
          log=print):
    ap = PolishApplier(docx, spec)
    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_polishc")
    unpack(docx, workdir)
    ap._before_counts = ap.counts(
        etree.parse(os.path.join(workdir, "word", "document.xml")).getroot())
    shutil.rmtree(workdir, ignore_errors=True)
    return ap.run(crs_path, approved_path, out, dry, log)
