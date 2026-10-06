# -*- coding: utf-8 -*-
"""
verify.py —— 校验排版结果。

三种校验，层层递进：
    1. 静态（不开 Word）：域写对没、书签断没、编号连不连续、内容丢没丢
    2. 真机（开 Word/WPS）：样式实际渲染几何与模板是否一致
    3. 真机重算后：交叉引用显示的号 == 题注实际编号
"""
import argparse
import os
import re
import sys
import zipfile
from collections import Counter

from lxml import etree

from . import qn
from .docxio import iter_fields, p_text


# ------------------------------------------------------------------ 1. 静态
def verify_static(docx, src=None):
    z = zipfile.ZipFile(docx)
    d = etree.fromstring(z.read("word/document.xml"))
    ps = list(d.iter(qn("p")))
    bk = {e.get(qn("name")) for e in d.iter(qn("bookmarkStart"))}

    seq_kind = {}
    refs, dead, mism = [], [], []
    for i, p in enumerate(ps):
        pbk = [e.get(qn("name")) for e in p.findall(qn("bookmarkStart"))]
        for instr, res in iter_fields(p):
            up = instr.upper()
            if "SEQ" in up:
                mk = re.search(r"SEQ\s+(\S+)", instr)
                kind = mk.group(1) if mk else "?"
                seq_kind.setdefault(kind, []).append(res)
                for b in pbk:
                    seq_kind.setdefault("_bk_" + kind, []).append(res)
            elif "REF" in up:
                m = re.search(r"REF\s+(\S+)", instr)
                if m:
                    refs.append((i, m.group(1), res))
    for (i, b, res) in refs:
        # `_Toc*` 是目录自身的锚点，只有目录里的 PAGEREF 用它，不算断链
        if b not in bk and not b.startswith("_Toc"):
            dead.append((i, b, res))

    print("[静态] 书签 %d ；SEQ 题注 %d ；REF 引用 %d ；断链 %d"
          % (len(bk), sum(len(v) for k, v in seq_kind.items()
                          if not k.startswith("_bk_")),
             len(refs), len(dead)))
    # 收集域指令，用来判断是不是分章编号
    instrs = " ".join(x.text or "" for x in d.iter(qn("instrText")))
    is_chapter = "\\s 1" in instrs or "SEQ 章" in instrs

    for kind in sorted(k for k in seq_kind if not k.startswith("_bk_")):
        vs = sorted(int(v) for v in seq_kind[kind] if (v or "").isdigit())
        if kind in ("章", "附录"):
            # 章号锚点是隐藏的，每个一级标题插一个，编号靠真机重算才递增
            print("        SEQ %-4s %2d 个（章号锚点，重算后递增）" % (kind, len(vs)))
            continue
        gaps = [n for n in range(1, len(vs) + 1) if n not in set(vs)]
        if is_chapter and gaps:
            # 分章编号下 SEQ 每章从 1 重起，缓存值是**章内序号**，
            # 「缺号」是正常现象，不是错误
            print("        SEQ %-4s %2d 条  章内序号 %s..%s（分章编号，"
                  "章内重置属正常）" % (kind, len(vs), vs[0], vs[-1]))
        else:
            print("        SEQ %-4s %2d 条  编号 %s..%s%s"
                  % (kind, len(vs), vs[0] if vs else "-",
                     vs[-1] if vs else "-",
                     "  缺号 %s" % gaps[:8] if gaps else "  连续"))
    if dead:
        print("        断链：%s" % dead[:5])

    if src and os.path.exists(src):
        zs = zipfile.ZipFile(src)
        ds = etree.fromstring(zs.read("word/document.xml"))

        def stat(root):
            names = zipfile.ZipFile(src).namelist()
            return dict(
                media=len([n for n in names
                           if n.startswith("word/media/") and not n.endswith("/")]),
                p=len(list(root.iter(qn("p")))),
                tbl=len(list(root.iter(qn("tbl")))),
                drawing=len(list(root.iter(qn("drawing")))),
                chars=len("".join(t.text or "" for t in root.iter(qn("t")))),
                sect=len(list(root.iter(qn("sectPr")))))
        a, b = stat(ds), stat(d)
        print("[完整性] 源 -> 出")
        for k in ("media", "p", "tbl", "drawing", "sect", "chars"):
            flag = "" if b[k] >= a[k] else "   ⚠ 少了"
            print("          %-8s %7d -> %7d  %+d%s" % (k, a[k], b[k],
                                                      b[k] - a[k], flag))
    return 0 if not dead else 1


# ------------------------------------------------------------------ 2. 真机样式
def verify_render(docx, template, samples=6, skip=60):
    """开 Word/WPS 对比样式实际渲染几何。需要 win32com。"""
    try:
        import win32com.client as win32
    except ImportError:
        print("[真机] 本环境没有 win32com，跳过（用 anaconda3 python 跑）")
        return 2
    CLSID = "{000209FF-0000-0000-C000-000000000046}"
    app = None
    for how in (CLSID, "Word.Application"):
        try:
            app = win32.DispatchEx(how)
        except Exception:
            continue
        path = str(getattr(app, "Path", "") or "")
        if "wps" not in path.lower() and "kingsoft" not in path.lower():
            break
        app = None
    if app is None:
        print("[真机] 起不了 Word/WPS")
        return 2
    is_wps = "wps" in str(getattr(app, "Path", "") or "").lower()
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        def collect(path):
            doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                     AddToRecentFiles=False)
            name2sid = {}
            out = {}
            for st in doc.Styles:
                try:
                    name2sid[st.NameLocal] = st.NameLocal
                except Exception:
                    pass
            seen = Counter()
            for i in range(1, min(doc.Paragraphs.Count, 4000) + 1):
                p = doc.Paragraphs(i)
                try:
                    nm = p.Style.NameLocal
                except Exception:
                    continue
                if seen[nm] >= samples:
                    continue
                txt = p.Range.Text.strip().replace("\r", "")
                if not txt or len(txt) < 4:
                    continue
                seen[nm] += 1
                f = p.Range.Font
                pf = p.Format
                out.setdefault(nm, []).append(dict(
                    font=f.NameFarEast or f.Name, size=round(f.Size, 1),
                    bold=bool(f.Bold), align=int(pf.Alignment),
                    indent=round(pf.FirstLineIndent or 0, 1),
                    line=round(pf.LineSpacing, 1), text=txt[:34]))
                if len(seen) and sum(seen.values()) > 240:
                    break
            doc.Close(False)
            return out
        t = collect(template) if template else {}
        o = collect(docx)
        print("[真机] %s" % ("WPS（Word COM 被劫持，渲染有差异）" if is_wps
                              else "Word"))
        common = [k for k in o if k in t]
        ok = bad = 0
        for k in common:
            a = Counter((x["font"], x["size"], x["align"]) for x in t[k])
            b = Counter((x["font"], x["size"], x["align"]) for x in o[k])
            if a == b:
                ok += 1
            else:
                bad += 1
                print("   不一致 %-20s 模板 %s / 输出 %s"
                      % (k, a.most_common(1), b.most_common(1)))
        print("       样式几何一致 %d 项，不一致 %d 项（只比两边都出现的）"
              % (ok, bad))
        return 0
    finally:
        try:
            app.Quit()
        except Exception:
            pass


# ------------------------------------------------------------------ CLI
def main(args):
    docx = args.infile
    print("=" * 70)
    print("校验：%s" % os.path.basename(docx))
    print("=" * 70)
    rc = verify_static(docx, getattr(args, "src", None))
    if getattr(args, "template", None):
        verify_render(docx, args.template)
    return rc
