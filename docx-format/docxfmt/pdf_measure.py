# -*- coding: utf-8 -*-
"""
pdf_measure.py —— 「渲染级第二判据」：导 PDF → 文本层逐行量。

为什么需要它（和 `paginate.py` 的区别）：

- `paginate.py` 走 **Word/WPS COM 对象模型**（`Range.Information`），逐段拿页号/坐标。
  快、能对回段落，但**依赖对象模型的自洽**——WPS 对 `PageSetup.*` 返回哨兵值，
  域/浮动图片/分栏下 `Information` 也可能与实际排版不一致。
- 本工具走 **导出 PDF 的文本层**：量的是「最终拿到手、会被第三方系统渲染的那份东西」。
  任何 PDF（Word 导的、WPS 导的、LibreOffice 导的、别人给的）都能量，
  与文档从哪来无关。**它才是交付物的真身。**

两者互为**交叉验证**：COM 报了 PDF 不报 → 多半是对象模型噪声；
PDF 报了 COM 不报 → 是真的排版问题（第三方渲染一定看得见）。

产出（JSON）：
    {
      "schema": 1, "tool": "pdf_measure",
      "source": docx 或 pdf 绝对路径, "source_md5": ...,
      "pdf": 实际量的 pdf 路径, "pdf_md5": ...,
      "engine": {Word/WPS 信息} 或 null（直接量现成 PDF 时）,
      "pages": N,
      "geometry": {"w","h","top","bottom","left","right"},   # pt
      "pitch": 行距(pt), "line_height": 单行高(pt),
      "per_page": [ {"page","lines","first_y","last_y","fill",
                     "max_gap","gaps","orphans"} ],
      "findings": {
        "orphan_lines":   [ {"page","text","norm_len","y","page_median_len"} ],
        "big_gaps":       [ {"page","after","before","gap_pt","gap_lines",
                             "y0","y1"} ],
        "low_fill_pages": [ {"page","fill","blank_bottom_pt","blank_lines",
                             "next_head":bool} ]
      },
      "summary": {"pages","orphan_lines","big_gaps","low_fill_pages"}
    }

退出码：0 成功 / 2 用法或量测失败 / 6 用法错误。
"""
import hashlib
import io
import json
import os
import re
import sys
import time

if __name__ == "__main__":
    # ⚠️ 只在「作为脚本直接运行」时重包 stdout。
    # 被 cli 以模块方式 import 时**绝不能**重包：那会夺走主进程 sys.stdout 的
    # buffer，旧 wrapper 被回收时关闭底层流 → 之后所有 print 报
    # 「I/O operation on closed file」。
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                      errors="replace")
    except Exception:                                   # noqa: BLE001
        pass

CJK = r"\u4e00-\u9fff\u3400-\u4dbf"
RE_CJK = re.compile(r"[%s]" % CJK)
# 页眉页脚/页码常见形态：纯数字、罗马数字、第N页
RE_FOOTERISH = re.compile(r"^\s*(?:第?\s*\d+\s*页?|[-—]?\s*\d+\s*[-—]?|"
                          r"[ivxlcIVXLC]{1,6})\s*$")
# 「像章节标题」的首行：第X章/节/篇、1.2 编号、一、二、三、附录/参考文献…
# ⚠️ 必须含「第X章」——本机 WPS 会把「第二章 方法」渲染成「第二章方法」（空格被吞），
#    漏了这一支会把大量**章末留白**误报成版面空白（实测 8 页文档即误报 1 页）。
RE_HEADISH = re.compile(
    r"^\s*(?:第\s*[一二三四五六七八九十百零〇\d]+\s*[章节篇部讲]|"
    r"\d{1,2}(?:\.\d+)*[\.\s、]|"
    r"[一二三四五六七八九十]+\s*[、.]|"
    r"附\s*录|参考文献|致\s*谢|摘\s*要|目\s*录|引\s*言|绪\s*论|"
    r"结\s*论|讨\s*论|前\s*言)")
# 末行若是这些标点收尾，才像「单字带标点独占一行」
TAIL_PUNCT = "。；：，！？、）」》”…"
TW = 20.0                                    # twips -> pt
DEF_PAGE = {"w": 595.3, "h": 841.9, "top": 72.0, "bottom": 72.0,
            "left": 90.0, "right": 90.0}
WD_FORMAT_PDF = 17
WORD_CLSID = "{000209FF-0000-0000-C000-000000000046}"


def say(*a):
    print(*a, flush=True)


def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _norm(s):
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", s or "")


# ---------------------------------------------------------------- 几何
def _docx_geom(docx):
    """从 docx 的 sectPr 静态读页面几何（pt）。

    ⚠️ 与 `polish.page_geom` 同口径：本机 WPS **不支持** `PageSetup.*`
       （返回 9999999.0 哨兵），几何只能从 XML 读。
    """
    import zipfile
    from lxml import etree
    g = dict(DEF_PAGE)
    try:
        with zipfile.ZipFile(docx) as z:
            xml = z.read("word/document.xml")
    except Exception:                                    # noqa: BLE001
        return g
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        root = etree.fromstring(xml)
    except Exception:                                    # noqa: BLE001
        return g
    body = root.find(W + "body")
    sect = None
    if body is not None:
        direct = [s for s in body if s.tag == W + "sectPr"]
        sect = direct[-1] if direct else next(iter(body.iter(W + "sectPr")),
                                              None)
    if sect is None:
        return g
    sz = sect.find(W + "pgSz")
    if sz is not None:
        for k in ("w", "h"):
            v = sz.get(W + k)
            if v and v.isdigit() and int(v) > 1400:
                g[k] = int(v) / TW
    mar = sect.find(W + "pgMar")
    if mar is not None:
        for k in ("top", "bottom", "left", "right"):
            v = mar.get(W + k)
            if v and re.match(r"^-?\d+$", v) and abs(int(v)) < 20000:
                g[k] = int(v) / TW
    return g


# ---------------------------------------------------------------- 导 PDF
def export_pdf(docx, out_pdf):
    """用 Word/WPS COM 把 docx 导成 PDF。返回 (engine_info, 用时秒)。"""
    import win32com.client as win32
    t0 = time.time()
    fallback = None
    engine = None
    for i, how in enumerate((WORD_CLSID, "Word.Application")):
        try:
            app = win32.DispatchEx(how)
        except Exception as e:                           # noqa: BLE001
            say("  [warn] %s 起不来：%r" % (how, repr(e)[:90]))
            continue
        path = str(getattr(app, "Path", "") or "")
        is_wps = ("wps" in path.lower()) or ("kingsoft" in path.lower())
        info = {"name": str(app.Name), "version": str(app.Version),
                "path": path, "is_wps": is_wps}
        if not is_wps:
            engine = (app, info)
            break
        say("  [warn] %s 实际落在 WPS（Path=%s）" % (how, path))
        if i == 0:
            fallback = (app, info)
            continue
        try:
            app.Quit()
        except Exception:                                # noqa: BLE001
            pass
    if engine is None:
        engine = fallback
    if engine is None:
        raise RuntimeError("起不了任何 Word/WPS 实例")
    app, info = engine

    try:
        app.Visible = False
    except Exception:                                    # noqa: BLE001
        pass
    try:
        app.DisplayAlerts = 0
    except Exception:                                    # noqa: BLE001
        pass

    doc = app.Documents.Open(os.path.abspath(docx), ReadOnly=True,
                             AddToRecentFiles=False)
    out_pdf = os.path.abspath(out_pdf)
    err = None
    try:
        # Word：ExportAsFixedFormat(OutputFileName, ExportFormat=17)
        doc.ExportAsFixedFormat(out_pdf, WD_FORMAT_PDF)
    except Exception as e1:                              # noqa: BLE001
        err = e1
        try:
            # WPS/兼容：SaveAs2(FileName, FileFormat=17)
            doc.SaveAs2(out_pdf, WD_FORMAT_PDF)
        except Exception as e2:                          # noqa: BLE001
            try:
                doc.Close(False)
            except Exception:                            # noqa: BLE001
                pass
            try:
                app.Quit()
            except Exception:                            # noqa: BLE001
                pass
            raise RuntimeError("导 PDF 失败：ExportAsFixedFormat=%r / "
                               "SaveAs2=%r" % (repr(e1)[:80], repr(e2)[:80]))
    try:
        doc.Close(False)
    except Exception:                                    # noqa: BLE001
        pass
    try:
        app.Quit()
    except Exception:                                    # noqa: BLE001
        pass
    if err:
        say("  [info] ExportAsFixedFormat 不可用，已用 SaveAs2 导出")
    if not os.path.exists(out_pdf):
        raise RuntimeError("导 PDF 后文件不存在：%s" % out_pdf)
    return info, round(time.time() - t0, 1)


# ---------------------------------------------------------------- 文本层解析
def _page_lines(page):
    """取一页的文本行 + 图片块 bbox（pt，左上原点）。"""
    d = page.get_text("dict")
    lines, images = [], []
    for blk in d.get("blocks", []):
        if blk.get("type") == 1:                         # 图片块
            bb = blk.get("bbox")
            if bb:
                images.append(tuple(bb))
            continue
        for ln in blk.get("lines", []):
            txt = "".join(sp.get("text", "") for sp in ln.get("spans", []))
            bb = ln.get("bbox")
            if bb is None:
                continue
            size = max((sp.get("size") or 0) for sp in ln.get("spans", [])) \
                or 0
            # 过滤零宽/纯空白行
            if not (txt or "").strip():
                continue
            lines.append({"text": txt,
                          "x0": bb[0], "y0": bb[1], "x1": bb[2], "y1": bb[3],
                          "size": round(size, 2)})
    lines.sort(key=lambda l: (round(l["y0"], 1), l["x0"]))
    return lines, images


def _suppressed(gap_y0, gap_y1, images):
    """空白带是否被图片/浮动对象覆盖（覆盖就不算「空白」）。返回覆盖面积比。"""
    band = gap_y1 - gap_y0
    if band <= 0:
        return 0.0
    cov = 0.0
    for (ix0, iy0, ix1, iy1) in images:
        lo, hi = max(gap_y0, iy0), min(gap_y1, iy1)
        if hi - lo > 0:
            cov += (hi - lo)
    return min(1.0, cov / band)


def measure_pdf(pdf, geom=None, min_gap_lines=4.0, orphan_max=2,
                low_fill=0.55, log=print):
    """量一份 PDF 的文本层。geom 不给则用默认页面几何。"""
    try:
        import pymupdf
    except ImportError:                                  # noqa: BLE001
        import fitz as pymupdf
    doc = pymupdf.open(pdf)
    g = dict(geom or DEF_PAGE)
    pages = doc.page_count

    raw = []                                             # 逐页原始行
    for i in range(pages):
        page = doc[i]
        lines, images = _page_lines(page)
        # 丢掉落在页边距之外的页眉页脚
        keep = []
        for l in lines:
            if l["y1"] <= g["top"] - 1.0 or l["y0"] >= g["h"] - g["bottom"] + 1.0:
                continue
            keep.append(l)
        raw.append({"lines": keep, "images": images,
                    "rect": (page.rect.width, page.rect.height)})

    # ---- 全局行距/行高：同页相邻行顶距的**众数**（避开段距抬高）
    diffs, heights = [], []
    for pg in raw:
        ls = pg["lines"]
        for a, b in zip(ls, ls[1:]):
            d = b["y0"] - a["y0"]
            if 5 < d < 60:
                diffs.append(d)
        for l in ls:
            if _norm(l["text"]) and l["size"]:
                heights.append(l["y1"] - l["y0"])
    pitch = _mode(diffs, default=21.0)
    line_h = _mode(heights, default=pitch)

    per_page, F = [], {"orphan_lines": [], "big_gaps": [],
                       "low_fill_pages": []}
    usable_h = max(1.0, g["h"] - g["top"] - g["bottom"])
    for i, pg in enumerate(raw):
        ls = pg["lines"]
        pno = i + 1
        if not ls:
            per_page.append({"page": pno, "lines": 0, "fill": 0.0,
                             "first_y": None, "last_y": None,
                             "max_gap": None, "gaps": [], "orphans": []})
            continue
        # 页内行距众数（比全局更贴合本页）
        p_diffs = [b["y0"] - a["y0"] for a, b in zip(ls, ls[1:])
                   if 5 < (b["y0"] - a["y0"]) < 60]
        p_pitch = _mode(p_diffs, default=pitch)
        med_len = _median([len(_norm(l["text"])) for l in ls]) or 1

        # ---- 大空白：相邻行之间的带
        gaps, orphans = [], []
        for a, b in zip(ls, ls[1:]):
            gap = b["y0"] - a["y1"]
            if gap >= min_gap_lines * p_pitch:
                cov = _suppressed(a["y1"], b["y0"], pg["images"])
                if cov >= 0.5:                           # 被图/表覆盖，不算
                    continue
                gaps.append({"gap_pt": round(gap, 1),
                             "gap_lines": round(gap / p_pitch, 1),
                             "after": (a["text"] or "").strip()[:30],
                             "before": (b["text"] or "").strip()[:30],
                             "y0": round(a["y1"], 1), "y1": round(b["y0"], 1)})
        for gg in gaps:
            F["big_gaps"].append(dict(gg, page=pno))

        # ---- 孤字行：整行归一化长度 ≤ orphan_max 且明显短于本页中位
        for l in ls:
            nl = len(_norm(l["text"]))
            raw_t = (l["text"] or "").strip()
            if nl == 0 or nl > orphan_max:
                continue
            if med_len < 8:
                continue
            if RE_FOOTERISH.match(raw_t):                # 页码等，忽略
                continue
            # 以标点收尾的单字行更像「被挤下来的一个字符」
            punct_tail = raw_t[-1] in TAIL_PUNCT
            orphan = {"page": pno, "text": raw_t, "norm_len": nl,
                      "y": round(l["y0"], 1), "page_median_len": med_len,
                      "punct_tail": punct_tail}
            orphans.append(orphan)
            F["orphan_lines"].append(orphan)

        text_top = min(l["y0"] for l in ls)
        text_bot = max(l["y1"] for l in ls)
        fill = max(0.0, min(1.0, (text_bot - text_top) / usable_h))
        # ---- 页底留白（章末正常 → 看下一页首行是否像标题）
        blank_bottom = max(0.0, (g["h"] - g["bottom"]) - text_bot)
        next_head = _next_page_head(raw, i)
        if fill < low_fill and i < pages - 1 and not next_head:
            F["low_fill_pages"].append({
                "page": pno, "fill": round(fill, 3),
                "blank_bottom_pt": round(blank_bottom, 1),
                "blank_lines": round(blank_bottom / p_pitch, 1),
                "next_head": False})
        per_page.append({
            "page": pno, "lines": len(ls), "fill": round(fill, 3),
            "first_y": round(text_top, 1), "last_y": round(text_bot, 1),
            "max_gap": (max((gg["gap_pt"] for gg in gaps), default=0.0)),
            "gaps": gaps, "orphans": orphans,
            "pitch": round(p_pitch, 2),
            "blank_bottom_pt": round(blank_bottom, 1)})

    doc.close()
    return {
        "schema": 1, "tool": "pdf_measure",
        "pages": pages,
        "geometry": {k: round(v, 1) for k, v in g.items()},
        "pitch": round(pitch, 2), "line_height": round(line_h, 2),
        "per_page": per_page,
        "findings": F,
        "summary": {"pages": pages,
                    "orphan_lines": len(F["orphan_lines"]),
                    "big_gaps": len(F["big_gaps"]),
                    "low_fill_pages": len(F["low_fill_pages"])},
    }


def _mode(xs, default):
    if not xs:
        return default
    from collections import Counter
    c = Counter(round(x * 2) / 2.0 for x in xs)
    top = c.most_common(2)
    best = top[0][0]
    # 众数若是「两倍行距」，取一半更接近单行
    if len(top) > 1 and abs(top[1][0] - best / 2.0) < 1.5 \
            and top[1][1] > top[0][1] * 0.5:
        best = top[1][0]
    return best


def _median(xs):
    if not xs:
        return 0
    s = sorted(xs)
    return s[len(s) // 2]


def _next_page_head(raw, i):
    """下一页首行是否像「章节标题」——是则本页留白属章末，正常。"""
    if i + 1 >= len(raw):
        return True
    nxt = raw[i + 1]["lines"]
    if not nxt:
        return False
    t = (nxt[0]["text"] or "").strip()
    if not t or len(t) > 40:
        return False
    return bool(RE_HEADISH.match(t))


# ---------------------------------------------------------------- 交叉验证
def compare(com_pages, pdf_res, geom=None, fs=12.0, log=print):
    """COM 判据 vs PDF 判据 的**对账**（只作证据，不改任何决定）。

    com_pages: `paginate.py` 的量页 json；pdf_res: 本工具的 json。
    返回 {"orphan": {...}, "blank": {...}}，按页对账。
    """
    g = dict(geom or (pdf_res or {}).get("geometry") or DEF_PAGE)
    left = g["left"]
    thr = 1.6 * fs
    lh = (pdf_res or {}).get("pitch") or 21.0

    com_orphan_pages, com_blank_pages = set(), set()
    if com_pages:
        by_i = {p["i"]: p for p in (com_pages.get("paras") or [])}
        per = {}
        for p in com_pages.get("paras") or []:
            if p.get("page") is not None:
                per.setdefault(p["page"], []).append(p)
        for p in com_pages.get("paras") or []:
            xe, ln, lne, pg = (p.get("xe"), p.get("ln"), p.get("ln_end"),
                               p.get("page"))
            tail = (p.get("tail") or "").strip()
            if (xe is not None and ln is not None and lne is not None
                    and lne > ln and tail and tail[-1] in TAIL_PUNCT
                    and len(_norm(p.get("t"))) >= 12
                    and xe - left <= thr):
                com_orphan_pages.add(pg)
        usable_bottom = g["h"] - g["bottom"]
        maxpage = max(per) if per else 0
        for pg, items in per.items():
            if pg >= maxpage:
                continue
            items = [x for x in items if (x.get("t") or "").strip()]
            if not items:
                continue
            last = max(items, key=lambda x: x.get("ye") or -1)
            if last.get("ye") is None:
                continue
            blank = usable_bottom - last["ye"]
            if blank < 4 * lh:
                continue
            lt = (last.get("t") or "")
            if "\x07" in lt or re.match(r"^\s*[图表]\s*\d", lt.strip()):
                continue
            ce = False
            for k in range(1, 4):
                q = by_i.get(last["i"] + k)
                if q is None:
                    ce = True
                    break
                if "\x0c" in (q.get("t") or ""):
                    ce = True
                    break
                if (q.get("t") or "").strip():
                    if RE_HEADISH.match((q.get("t") or "").strip()) \
                            and len((q.get("t") or "").strip()) < 60:
                        ce = True
                    break
            if not ce:
                com_blank_pages.add(pg)

    pdf_orphan_pages = {o["page"] for o in (pdf_res or {}).get(
        "findings", {}).get("orphan_lines", [])}
    pdf_blank_pages = {b["page"] for b in (pdf_res or {}).get(
        "findings", {}).get("big_gaps", [])} | {
        x["page"] for x in (pdf_res or {}).get("findings", {}).get(
            "low_fill_pages", [])}

    res = {
        "com_orphan_pages": sorted(com_orphan_pages),
        "pdf_orphan_pages": sorted(pdf_orphan_pages),
        "orphan_only_com": sorted(com_orphan_pages - pdf_orphan_pages),
        "orphan_only_pdf": sorted(pdf_orphan_pages - com_orphan_pages),
        "orphan_both": sorted(com_orphan_pages & pdf_orphan_pages),
        "com_blank_pages": sorted(com_blank_pages),
        "pdf_blank_pages": sorted(pdf_blank_pages),
        "blank_only_com": sorted(com_blank_pages - pdf_blank_pages),
        "blank_only_pdf": sorted(pdf_blank_pages - com_blank_pages),
        "blank_both": sorted(com_blank_pages & pdf_blank_pages),
    }
    if log:
        log("[②③④ 交叉验证] 孤字页 COM=%d / PDF=%d（仅COM %d，仅PDF %d，共同 %d）"
            % (len(res["com_orphan_pages"]), len(res["pdf_orphan_pages"]),
               len(res["orphan_only_com"]), len(res["orphan_only_pdf"]),
               len(res["orphan_both"])))
        log("[②③④ 交叉验证] 空白页 COM=%d / PDF=%d（仅COM %d，仅PDF %d，共同 %d）"
            % (len(res["com_blank_pages"]), len(res["pdf_blank_pages"]),
               len(res["blank_only_com"]), len(res["blank_only_pdf"]),
               len(res["blank_both"])))
        if res["blank_only_pdf"] or res["orphan_only_pdf"]:
            log("[②③④ 交叉验证] ⚠ 仅 PDF 报出的页 = 第三方渲染真会出问题的页："
                "空白 %s，孤字 %s"
                % (res["blank_only_pdf"][:12], res["orphan_only_pdf"][:12]))
    return res


# ---------------------------------------------------------------- 编排
def measure(src, pdf=None, out_pdf=None, keep_pdf=True, geom=None,
            min_gap_lines=4.0, orphan_max=2, low_fill=0.55, log=print):
    """总入口：src 可为 .docx（先导 PDF）或 .pdf（直接量）。"""
    t0 = time.time()
    src = os.path.abspath(src)
    engine = None
    if src.lower().endswith(".pdf"):
        pdf = src
        g = geom
    else:
        g = geom or _docx_geom(src)
        if pdf is None:
            pdf = out_pdf or (os.path.splitext(src)[0] + "_渲染量测.pdf")
            engine, secs = export_pdf(src, pdf)
            log("[export] %s %s -> %s（%.1fs）"
                % (engine["name"], engine["version"], pdf, secs))
    res = measure_pdf(pdf, geom=g, min_gap_lines=min_gap_lines,
                      orphan_max=orphan_max, low_fill=low_fill, log=log)
    res["source"] = src
    res["source_md5"] = _md5(src)
    res["pdf"] = pdf
    res["pdf_md5"] = _md5(pdf)
    res["engine"] = engine
    res["elapsed"] = round(time.time() - t0, 1)
    res["captured_at"] = time.strftime("%Y-%m-%dT%H:%M:%S+08:00")
    if engine is None and not keep_pdf:
        try:
            os.remove(pdf)
        except Exception:                                # noqa: BLE001
            pass
    return res


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(
        prog="pdf_measure", description="渲染级第二判据：导 PDF → 文本层逐行量")
    ap.add_argument("infile", help=".docx（自动导 PDF）或 .pdf（直接量）")
    ap.add_argument("-o", "--out", help="结果 json（默认 <名>_PDF量测.json）")
    ap.add_argument("--pdf", help="已有的 PDF（给则不再导出）")
    ap.add_argument("--out-pdf", help="导出的 PDF 路径")
    ap.add_argument("--min-gap-lines", type=float, default=4.0,
                    help="判定大空白的行数阈值（默认 4）")
    ap.add_argument("--orphan-max", type=int, default=2,
                    help="孤字行最大字符数（默认 2）")
    ap.add_argument("--low-fill", type=float, default=0.55,
                    help="页填充率低于此值视为留白页（默认 0.55）")
    ap.add_argument("--vs-com", help="给 paginate.py 的量页 json → 交叉验证")
    a = ap.parse_args(argv)
    if not os.path.exists(a.infile):
        print("找不到 %s" % a.infile)
        return 6
    try:
        res = measure(a.infile, pdf=a.pdf, out_pdf=a.out_pdf,
                      min_gap_lines=a.min_gap_lines, orphan_max=a.orphan_max,
                      low_fill=a.low_fill)
    except Exception as e:                               # noqa: BLE001
        print("[pdf_measure] 失败：%r" % (e,))
        return 2
    if a.vs_com:
        com = json.load(open(a.vs_com, encoding="utf-8"))
        res["cross_check"] = compare(com, res)
    out = a.out or os.path.join(
        os.path.dirname(res["source"]),
        os.path.splitext(os.path.basename(res["source"]))[0] + "_PDF量测.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    s = res["summary"]
    say("-> %s（%d 页；孤字行 %d，大空白 %d，低填充页 %d；行距 %.1fpt，%.1fs）"
        % (out, s["pages"], s["orphan_lines"], s["big_gaps"],
           s["low_fill_pages"], res["pitch"], res["elapsed"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
