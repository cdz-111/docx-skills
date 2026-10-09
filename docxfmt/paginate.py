# -*- coding: utf-8 -*-
"""
paginate.py —— 渲染级量页（独立进程，需 win32com）。

为什么独立进程：主流程用的解释器未必有 pywin32；拆开让两边各用各的。
和 `winfield.py` 同一套路。

产出（JSON）：
    {
      "schema": 1,
      "engine": {"name","version","path","is_wps"},
      "docx": 绝对路径, "docx_md5": ..., "paragraphs": N, "pages": P,
      "stats": {"words","lines","chars","paras"},
      "paras": [ {"i":1,"page":1,"y":72.0,"ln":1,"x":281.5,
                  "page_end":1,"ye":84.0,"xe":300.1,"ln_end":1,
                  "tail":"...末20字","orphan":false,"t":"...首40字"} ]
    }

⚠️ 本机实测：Word CLSID 被 WPS 劫持（Path 在 WPS 目录下、Name 却叫
   "Microsoft Word"）。WPS **不支持** `PageSetup.*`（返回 9999999.0 哨兵），
   但 `Range.Information(...)` 全常量可用（实测 page(3)/vertY(6)/horizX(5)/
   lineNo(10) 全部 36/36 成功）。所以页高与边距**必须**由调用方从
   `sectPr` 静态读取，不要问 COM。

坐标单位：pt，基准是**页面左上角**（不是版心）。
"""
import hashlib
import io
import json
import os
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                              errors="replace")
import win32com.client as win32          # noqa: E402

WORD_CLSID = "{000209FF-0000-0000-C000-000000000046}"

WD_PAGE = 3          # wdActiveEndPageNumber
WD_Y = 6             # wdVerticalPositionRelativeToPage
WD_X = 5             # wdHorizontalPositionRelativeToPage
WD_LINE = 10         # wdFirstCharacterLineNumber
WD_SECT = 2          # wdActiveEndSectionNumber

STAT_WORDS, STAT_LINES, STAT_PAGES, STAT_CHARS, STAT_PARAS = 0, 1, 2, 3, 4

# 段末若是这些标点，才值得查"孤字独占一行"
TAIL_PUNCT = "。；：，！？、）」》”…"


def say(*a):
    print(*a, flush=True)


def open_app():
    """起 Word/WPS。返回 (app, engine_info)。"""
    fallback = None
    for i, how in enumerate((WORD_CLSID, "Word.Application")):
        try:
            app = win32.DispatchEx(how)
        except Exception as e:                      # noqa: BLE001
            say("  [warn] %s 起不来：%r" % (how, repr(e)[:90]))
            continue
        path = str(getattr(app, "Path", "") or "")
        is_wps = ("wps" in path.lower()) or ("kingsoft" in path.lower())
        info = {"name": str(app.Name), "version": str(app.Version),
                "path": path, "is_wps": is_wps}
        if not is_wps:
            return app, info
        say("  [warn] %s 实际落在 WPS（Path=%s）" % (how, path))
        if i == 0:
            fallback = app
            continue
        try:
            app.Quit()
        except Exception:                            # noqa: BLE001
            pass
    if fallback is None:
        raise RuntimeError("起不了任何 Word/WPS 实例")
    path = str(getattr(fallback, "Path", "") or "")
    return fallback, {"name": str(fallback.Name), "version": str(fallback.Version),
                      "path": path, "is_wps": True}


def _md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def measure(src, limit=0, quick=False, verbose=True):
    """打开 src 逐段量。返回结果 dict。"""
    t0 = time.time()
    app, engine = open_app()
    say("[app] %s %s (is_wps=%s) %.1fs"
        % (engine["name"], engine["version"], engine["is_wps"],
           time.time() - t0))
    try:
        app.Visible = False
    except Exception:                                # noqa: BLE001
        pass
    try:
        app.DisplayAlerts = 0
    except Exception:                                # noqa: BLE001
        pass

    t1 = time.time()
    doc = app.Documents.Open(os.path.abspath(src), ReadOnly=True,
                             AddToRecentFiles=False)
    say("[open] %.1fs | Paragraphs=%d Tables=%d InlineShapes=%d"
        % (time.time() - t1, doc.Paragraphs.Count, doc.Tables.Count,
           doc.InlineShapes.Count))

    stats = {}
    for key, code in (("words", STAT_WORDS), ("lines", STAT_LINES),
                      ("paras", STAT_PARAS), ("chars", STAT_CHARS)):
        try:
            stats[key] = int(doc.ComputeStatistics(code))
        except Exception:                            # noqa: BLE001
            stats[key] = None
    pages = stats.get("lines") and None
    try:
        pages = int(doc.ComputeStatistics(STAT_PAGES))
    except Exception:                                # noqa: BLE001
        pages = int(doc.Range(0, 0).Information(WD_PAGE))

    n = doc.Paragraphs.Count
    if limit:
        n = min(n, limit)
    say("[measure] %d 段（quick=%s）…" % (n, quick))

    paras = []
    t2 = time.time()
    bad = 0
    for i in range(1, n + 1):
        try:
            r = doc.Paragraphs(i).Range
            txt = r.Text or ""
            item = {"i": i,
                    "page": int(r.Information(WD_PAGE)),
                    "y": round(float(r.Information(WD_Y)), 2),
                    "ln": int(r.Information(WD_LINE)),
                    "x": round(float(r.Information(WD_X)), 2),
                    "t": txt[:40]}
        except Exception as e:                       # noqa: BLE001
            bad += 1
            item = {"i": i, "err": repr(e)[:70]}
            paras.append(item)
            continue
        if not quick:
            try:
                # 折叠到"最后一个可见字符"（End-1 是段落标记前一位）
                pos = max(r.Start, r.End - 1)
                e_ = doc.Range(pos, pos)
                item["page_end"] = int(e_.Information(WD_PAGE))
                item["ye"] = round(float(e_.Information(WD_Y)), 2)
                item["xe"] = round(float(e_.Information(WD_X)), 2)
                item["ln_end"] = int(e_.Information(WD_LINE))
            except Exception:                        # noqa: BLE001
                pass
            # 尾巴（去掉段落标记 \r / \x07）
            tail = txt.rstrip("\r\n\x07\x0b")
            item["tail"] = tail[-20:]
            # 孤字行候选：段末是标点 且 末行号 > 首行号（说明换过行）
            le, ls = item.get("ln_end"), item.get("ln")
            if (le and ls and le > ls and tail
                    and tail[-1] in TAIL_PUNCT):
                item["orphan_cand"] = True
        paras.append(item)
        if verbose and i % 500 == 0:
            say("   … %d/%d  %.1fs" % (i, n, time.time() - t2))
    say("[measure] 完成 %d 段，失败 %d，耗时 %.1fs"
        % (n, bad, time.time() - t2))

    try:
        doc.Close(False)
    except Exception:                                # noqa: BLE001
        pass
    try:
        app.Quit()
    except Exception:                                # noqa: BLE001
        pass

    return {
        "schema": 1,
        "engine": engine,
        "docx": os.path.abspath(src),
        "docx_md5": _md5(src),
        "paragraphs": n,
        "pages": pages,
        "stats": stats,
        "paras": paras,
        "elapsed": round(time.time() - t0, 1),
        "captured_at": time.strftime("%Y-%m-%dT%H:%M:%S+08:00"),
    }


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(prog="paginate", description="渲染级量页")
    ap.add_argument("infile")
    ap.add_argument("-o", "--out", help="结果 json（默认 <名>_量页.json）")
    ap.add_argument("--limit", type=int, default=0, help="只量前 N 段（调试用）")
    ap.add_argument("--quick", action="store_true", help="只取 page/y/ln")
    a = ap.parse_args(argv)
    out = a.out or os.path.join(
        os.path.dirname(os.path.abspath(a.infile)),
        os.path.splitext(os.path.basename(a.infile))[0] + "_量页.json")
    res = measure(a.infile, limit=a.limit, quick=a.quick)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    say("-> %s（%d 段，%s 页，%.1fs）"
        % (out, res["paragraphs"], res["pages"], res["elapsed"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
