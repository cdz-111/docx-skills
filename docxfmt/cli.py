# -*- coding: utf-8 -*-
"""
cli.py —— 命令行入口。

    python -m docxfmt run     <in.docx> -o out.docx --template t.docx [选项]
    python -m docxfmt inspect <in.docx>                     # 只侦察，不改
    python -m docxfmt spec    --init spec.json               # 生成 spec 样例
    python -m docxfmt fields  <in.docx> -o out.docx          # 重算全部域
    python -m docxfmt verify  <in.docx> [--template t.docx]  # 校验

run 的常用选项（都能用 --spec 一次说清，这里是快捷方式）：
    --template PATH      照这个模板的样式排
    --spec PATH          完整规范 json（会覆盖前面的快捷选项）
    --toc / --no-toc     正文目录
    --figure-list        插图目录
    --table-list         插表目录
    --caption-scheme global|chapter
    --number-sep SEP
    --pagebreak-levels 1,2
    --dry-run
    --tree-json PATH --tree-txt PATH
"""
import argparse
import json
import os
import sys

from .spec import DEFAULTS, Spec, write_sample


def _ensure_stdout():
    try:
        buf = sys.stdout.detach()
    except Exception:
        buf = sys.stdout.buffer
    import io
    return io.TextIOWrapper(buf, encoding="utf-8", errors="replace")


def _kvlist(s):
    return [x for x in (s or "").replace("，", ",").split(",") if x.strip()]


def build_spec(args):
    """命令行 -> spec dict。--spec 提供基底，快捷选项在其上覆盖。

    用 getattr 兜底：styles / lint 等命令没有 run 的全套参数。
    """
    g = lambda k, d=None: getattr(args, k, d)      # noqa: E731
    base = {}
    if g("spec"):
        with open(g("spec"), "r", encoding="utf-8") as f:
            base = json.load(f)
    if g("template"):
        base.setdefault("source", {})
        base["source"]["style_from"] = "template"
        base["source"]["template"] = g("template")
    if g("requirement"):
        base.setdefault("source", {})
        base["source"]["style_from"] = "requirement"
        base["source"]["requirement"] = g("requirement")
    if g("caption_scheme"):
        base.setdefault("numbering", {})["caption"] = g("caption_scheme")
    if g("formula_scheme"):
        base.setdefault("numbering", {})["formula"] = g("formula_scheme")
    if g("number_sep"):
        base.setdefault("numbering", {})["sep"] = g("number_sep")
    if g("pagebreak_levels") is not None:
        base.setdefault("pagebreak", {})["before_heading"] = \
            [int(x) for x in _kvlist(g("pagebreak_levels"))]
    for key, val in (("toc", g("toc")), ("figure_list", g("figure_list")),
                     ("table_list", g("table_list")),
                     ("captions", g("captions")), ("xref", g("xref"))):
        if val is not None:
            base.setdefault(key, {})["enabled"] = val
    if g("toc_levels"):
        base.setdefault("toc", {})["levels"] = g("toc_levels")
    return base


def cmd_propose(args):
    """导出「待拍板清单」—— 工具判不准的地方，带着上下文交给人/AI 判断。"""
    from .docxio import Docx, read_style_names, unpack
    from .propose import build_proposal, to_report
    workdir = args.workdir or os.path.join(
        os.path.dirname(os.path.abspath(args.infile)), "_propose")
    unpack(args.infile, workdir)
    doc = Docx(workdir)
    sid2name = read_style_names(workdir)
    prop = build_proposal(doc.root, sid2name)
    out_json = args.out or os.path.splitext(args.infile)[0] + "_待拍板.json"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(prop, f, ensure_ascii=False, indent=1)
    print(to_report(prop, args.report))
    print()
    print("-> %s" % out_json)
    if args.report:
        print("-> %s" % args.report)
    print()
    print("下一步：编辑 %s 里的 overrides 字段（或直接手写），" % out_json)
    print("        然后 run 时带 --overrides 喂回去")
    import shutil
    shutil.rmtree(workdir, ignore_errors=True)
    return 0


def cmd_run(args):
    from .pipeline import run
    spec = build_spec(args)
    overrides = None
    if args.overrides:
        with open(args.overrides, "r", encoding="utf-8") as f:
            data = json.load(f)
        overrides = data.get("overrides", data) if isinstance(data, dict) \
            else data
    run(args.infile, args.out, spec,
        workdir=args.workdir, tree_json=args.tree_json, tree_txt=args.tree_txt,
        dry_run=args.dry_run, overrides=overrides)
    print("\n下一步：重算域（否则目录和交叉引用显示的是缓存值）")
    print("  python -m docxfmt fields %s -o %s_域重算.docx"
          % (args.out, os.path.splitext(os.path.basename(args.out))[0]))
    return 0


def cmd_inspect(args):
    from collections import Counter
    from .docxio import Docx, read_style_names, unpack
    from .recognize import (find_cover_end, find_toc_block, parse, stats,
                           tree_report)
    workdir = args.workdir or os.path.join(
        os.path.dirname(os.path.abspath(args.infile)), "_inspect")
    unpack(args.infile, workdir)
    sid2name = read_style_names(workdir)
    doc = Docx(workdir)
    root, index, cover = parse(doc.root, sid2name)
    st = stats(root)
    print("=" * 70)
    print("侦察：%s" % args.infile)
    print("=" * 70)
    print("段落 %d ；表格 %d ；图片段 %d"
          % (len(doc.paras()),
             len(list(doc.root.iter(qn_t("tbl")))),
             len(list(doc.root.iter(qn_t("drawing"))))))
    print("封面区 [0,%d)" % cover)
    print("已有目录区：%s" % (find_toc_block(doc.paras(), sid2name),))
    print()
    print("节点统计：")
    for k, v in st.most_common():
        print("   %-14s %d" % (k, v))
    print()
    print("样式使用（前 20）：")
    c = Counter()
    for p in doc.paras():
        c[sid2name.get(_sid(p), "无样式")] += 1
    for k, v in c.most_common(20):
        print("   %-24s %d" % (k, v))
    if args.tree_txt:
        tree_report(root, args.tree_txt)
        print("\n树形报告 -> %s" % args.tree_txt)
    return 0


def qn_t(t):
    from . import qn
    return qn(t)


def _sid(p):
    from . import qn
    ppr = p.find(qn("pPr"))
    if ppr is None:
        return None
    ps = ppr.find(qn("pStyle"))
    return ps.get(qn("val")) if ps is not None else None


def cmd_spec(args):
    p = write_sample(args.init, args.mode)
    print("spec 样例已写出：%s" % p)
    print("照着改，然后：python -m docxfmt run in.docx -o out.docx --spec %s" % p)
    return 0


def cmd_fields(args):
    """重算全部域（Word/WPS COM）。本机 Word COM 被 WPS 劫持，只能用 WPS。"""
    import shutil
    here = os.path.dirname(os.path.abspath(__file__))
    runner = os.path.join(here, "winfield.py")
    if not os.path.exists(runner):
        print("找不到 winfield.py")
        return 2
    import subprocess
    src = os.path.abspath(args.infile)
    out = os.path.abspath(args.out or args.infile)
    # win32com 只装在部分解释器里，逐个探测比猜路径可靠。
    cands = [sys.executable]
    env_py = os.environ.get("DOCXFMT_PYTHON")
    if env_py:
        cands.insert(0, env_py)
    for base in ("Anaconda3", "anaconda3", "Miniconda3", "miniconda3"):
        p = os.path.expanduser("~") + "\\" + base + "\\python.exe"
        cands.append(p)
    for py in cands:
        if not py or not os.path.exists(py):
            continue
        if os.path.abspath(py) == os.path.abspath(sys.executable) and \
                cands[0] != sys.executable:
            pass
        try:
            probe = subprocess.run(
                [py, "-c", "import win32com.client"], capture_output=True,
                timeout=20)
        except Exception:
            continue
        if probe.returncode != 0:
            continue                       # 这个解释器没有 win32com
        r = subprocess.run([py, runner, src, out], capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0:
            print(r.stdout)
            return 0
        print(r.stdout or "", r.stderr or "")
    print("重算域失败：没找到装了 win32com 的 Python。")
    print("  可设环境变量指定：DOCXFMT_PYTHON=D:\\path\\to\\python.exe")
    return 1


def cmd_verify(args):
    from . import verify
    return verify.main(args)


# ------------------------------------------------------------------ 新命令
def _spec_for(args):
    """给只读类命令（styles / lint）构造生效 spec（含默认模板兜底）。"""
    from .pipeline import _load_requirement
    from .spec import Spec
    sp = Spec(build_spec(args))
    if sp["source"].get("style_from") == "requirement":
        sp = _load_requirement(sp)
    return sp


def _stem(path):
    return os.path.splitext(os.path.basename(path))[0]


def _sibling(path, suffix):
    return os.path.join(os.path.dirname(os.path.abspath(path)),
                        _stem(path) + "_" + suffix)


def cmd_styles(args):
    """S1 样式确认单：逐部分列出目标样式（可手改），只出单不套用。"""
    from . import crs as CRS
    from .confirm import collect, render_html
    spec = _spec_for(args)
    inv, rows, resolved = collect(args.infile, spec)
    hp = args.html or _sibling(args.infile, "样式确认单.html")
    jp = args.out or _sibling(args.infile, "样式确认单.json")
    render_html(inv, rows, resolved, args.infile, hp)
    with open(jp, "w", encoding="utf-8") as f:
        import json as _j
        _j.dump(resolved, f, ensure_ascii=False, indent=2)
    miss = [r["label"] for r in rows if r["missing"]]
    alias = ["%s→%s" % (r["target"], r["resolved_as"]) for r in rows
             if r.get("resolved_as") and r["resolved_as"] != r["target"]]
    print("=" * 70)
    print("样式确认单：%s" % os.path.basename(args.infile))
    print("=" * 70)
    print("段落 %d ；表格 %d ；图片 %d ；公式 %d ；题注 %d ；批注 %d"
          % (inv.get("paragraphs", 0), inv.get("tables", 0),
             inv.get("drawings", 0), inv.get("math", 0),
             inv.get("captions", 0), inv.get("comment_count", 0)))
    print("目标样式走别名命中的：%s" % ("、".join(alias) if alias else "无"))
    print("目标样式在文档里找不到的部分：%s"
          % ("、".join(miss) if miss else "无"))
    print("文档现有样式：%s" % "、".join(sorted(inv.get("style_names", []))[:15]))
    print()
    print("-> %s" % hp)
    print("-> %s（确认/手改后当 --spec 喂回 run）" % jp)
    return 0


def cmd_lint(args):
    """S5 正文校验：只出改动事项单，绝不改文档。"""
    from . import crs as CRS
    from .lint import analyze
    spec = _spec_for(args)
    c = analyze(args.infile, spec)
    jp = args.out or _sibling(args.infile, "改动事项.json")
    hp = args.html or _sibling(args.infile, "改动事项.html")
    CRS.write_json(c, jp)
    CRS.render_html(c, hp)
    s = CRS.summarize(c)
    print("=" * 70)
    print("正文校验（只出清单，不改文档）：%s" % os.path.basename(args.infile))
    print("=" * 70)
    print("共 %d 项　分档 %s" % (s["total"], s["by_grade"]))
    print()
    print("-> %s" % hp)
    print("-> %s" % jp)
    print()
    print("下一步：勾选后执行 —— python -m docxfmt revise <文档> --crs %s" % jp)
    return 0


def cmd_revise(args):
    """S6 修订落盘：只执行已批准的项，以修订模式写入并逐条批注。"""
    from . import revise as RV
    r = RV.apply(args.infile, args.crs, approved_path=args.approved,
                 out=args.out, author=args.author, dry=args.dry)
    rp = args.receipt or _sibling(args.infile, "改动回执.md")
    RV.write_receipt(r, rp)
    print("-> 回执 %s" % rp)
    if r.get("output"):
        print("-> 产物 %s" % r["output"])
    return 0 if r["summary"]["failed"] == 0 else 1


def _find_pywin32():
    """找一个装了 win32com 的解释器（COM 类功能都用它）。"""
    import subprocess
    cands = []
    env_py = os.environ.get("DOCXFMT_PYTHON")
    if env_py:
        cands.append(env_py)
    cands.append(sys.executable)
    for base in ("anaconda3", "Anaconda3", "miniconda3", "Miniconda3"):
        cands.append(os.path.expanduser("~") + "\\" + base + "\\python.exe")
    for py in cands:
        if not py or not os.path.exists(py):
            continue
        try:
            r = subprocess.run([py, "-c", "import win32com.client"],
                               capture_output=True, timeout=25)
        except Exception:                            # noqa: BLE001
            continue
        if r.returncode == 0:
            return py
    return None


def _render_pages(docx, pages_out=None, limit=0, quick=False):
    """调 paginate.py 做渲染级量页，返回结果 dict（失败返回 None）。"""
    import subprocess
    py = _find_pywin32()
    if not py:
        print("[polish] 没找到装了 win32com 的 Python → 渲染级检查降级为跳过。")
        print("         可设 DOCXFMT_PYTHON=D:\\path\\to\\python.exe 指定。")
        return None
    here = os.path.dirname(os.path.abspath(__file__))
    runner = os.path.join(here, "paginate.py")
    want = pages_out or _sibling(docx, "量页.json")
    cmd = [py, runner, os.path.abspath(docx), "-o", want]
    if limit:
        cmd += ["--limit", str(limit)]
    if quick:
        cmd += ["--quick"]
    print("[polish] 用 %s 量页 …" % os.path.basename(py))
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print("[polish] 量页失败：\n%s\n%s" % (r.stdout or "", r.stderr or ""))
        return None
    for line in (r.stdout or "").splitlines()[-4:]:
        print("          " + line)
    with open(want, "r", encoding="utf-8") as f:
        return json.load(f)


def cmd_polish(args):
    """S4 排版体检：六条硬要求，只出清单；--apply 执行结构类改动。"""
    from . import crs as CRS
    from . import polish as PL
    spec = _spec_for(args)

    if getattr(args, "apply", False):
        if not args.crs:
            print("[polish] --apply 需要 --crs 指定改动事项单")
            return 2
        r = PL.apply(args.infile, args.crs, approved_path=args.approved,
                     out=args.out_docx, spec=spec, dry=args.dry)
        rp = args.receipt or _sibling(args.infile, "排版回执.md")
        from .applykit import write_receipt
        write_receipt(r, rp, title="排版改动回执")
        print("-> 回执 %s" % rp)
        if r.get("output"):
            print("-> 产物 %s" % r["output"])
        return 0 if r["summary"]["failed"] == 0 else 1

    pages = None
    if args.pages:
        with open(args.pages, "r", encoding="utf-8") as f:
            pages = json.load(f)
    elif not args.no_render:
        pages = _render_pages(args.infile)

    c = PL.analyze(args.infile, spec, pages)
    jp = args.out or _sibling(args.infile, "排版事项.json")
    hp = args.html or _sibling(args.infile, "排版事项.html")
    CRS.write_json(c, jp)
    CRS.render_html(c, hp)
    s = CRS.summarize(c)
    print()
    print("共 %d 项　分档 %s" % (s["total"], s["by_grade"]))
    print("-> %s" % hp)
    print("-> %s" % jp)
    print()
    print("下一步：勾选后执行（分页/表格边框类）——")
    print("  python -m docxfmt polish <文档> --apply --crs %s" % jp)
    return 0


def cmd_captions(args):
    """S3 题注补插：找「该有题注但没有」的图/表，只出清单。"""
    from . import crs as CRS
    from . import insert_captions as IC
    spec = _spec_for(args)

    if getattr(args, "apply", False):
        if not args.crs:
            print("[captions] --apply 需要 --crs")
            return 2
        r = IC.apply(args.infile, args.crs, approved_path=args.approved,
                     out=args.out_docx, spec=spec, dry=args.dry)
        rp = args.receipt or _sibling(args.infile, "题注回执.md")
        from .applykit import write_receipt
        write_receipt(r, rp, title="题注补插回执")
        print("-> 回执 %s" % rp)
        if r.get("output"):
            print("-> 产物 %s" % r["output"])
        return 0 if r["summary"]["failed"] == 0 else 1

    c = IC.analyze(args.infile, spec)
    jp = args.out or _sibling(args.infile, "题注事项.json")
    hp = args.html or _sibling(args.infile, "题注事项.html")
    CRS.write_json(c, jp)
    CRS.render_html(c, hp)
    s = CRS.summarize(c)
    print()
    print("共 %d 项　分档 %s" % (s["total"], s["by_grade"]))
    print("-> %s" % hp)
    print("-> %s" % jp)
    return 0


def cmd_facts(args):
    """S5b 事实核查：抽原子断言 →（联网核）→ 落 CRS / 出核查表。"""
    from . import factcheck as FC
    res = FC.extract(args.infile, spec=_spec_for(args))
    jp = args.out or _sibling(args.infile, "待核查事实.json")
    mp = args.md or _sibling(args.infile, "事实核查表.md")

    if args.verdicts:
        v = FC.load_json(args.verdicts)
        n = FC.apply_verdicts(res["claims"], v)
        print("[facts] 回填判定 %d 条" % n)

    FC.save_json(res, jp)
    import time as _t
    FC.render_md(res["claims"], mp, doc_name=os.path.basename(args.infile),
                 date_str=args.date or _t.strftime("%Y-%m-%d"))
    print("-> %s" % mp)
    print("-> %s" % jp)

    if args.to_crs:
        c, n = FC.to_crs(args.infile, res["claims"], spec=_spec_for(args))
        from . import crs as CRS
        CRS.write_json(c, args.to_crs)
        hp = args.html or os.path.splitext(args.to_crs)[0] + ".html"
        CRS.render_html(c, hp)
        s = CRS.summarize(c)
        print("-> 事实核查 CRS %s（%d 项，分档 %s）"
              % (args.to_crs, n, s["by_grade"]))
        print("-> %s" % hp)
    else:
        print()
        print("下一步：逐条联网核查 → 写 verdicts.json → 回填：")
        print("   python -m docxfmt facts <文档> --verdicts verdicts.json "
              "--to-crs <文档>_事实核查.json")
    return 0


def _add_common_spec(sp):
    sp.add_argument("--spec", help="完整规范 json（优先）")
    sp.add_argument("--template", help="参考模板 docx")
    sp.add_argument("--requirement", help="文字规范 json")
    return sp


def main(argv=None):
    sys.stdout = _ensure_stdout()
    ap = argparse.ArgumentParser(prog="docxfmt",
                                 description="docx 排版工具箱")
    sub = ap.add_subparsers(dest="cmd", required=True)

    # run
    r = sub.add_parser("run", help="排版")
    r.add_argument("infile")
    r.add_argument("-o", "--out", required=True)
    r.add_argument("--spec", help="规范 json")
    r.add_argument("--template", help="参考模板 docx")
    r.add_argument("--requirement", help="文字规范 json")
    r.add_argument("--caption-scheme", choices=["global", "chapter"])
    r.add_argument("--formula-scheme", choices=["none", "global", "chapter"])
    r.add_argument("--number-sep", help="分章编号分隔符，默认 -")
    r.add_argument("--pagebreak-levels", help="几个级标题前分页，如 1 或 1,2")
    r.add_argument("--toc-levels", help='目录层级，如 "1-3"')
    r.add_argument("--toc", dest="toc", action="store_true", default=None)
    r.add_argument("--no-toc", dest="toc", action="store_false")
    r.add_argument("--figure-list", dest="figure_list", action="store_true",
                   default=None)
    r.add_argument("--no-figure-list", dest="figure_list", action="store_false")
    r.add_argument("--table-list", dest="table_list", action="store_true",
                   default=None)
    r.add_argument("--no-table-list", dest="table_list", action="store_false")
    r.add_argument("--captions", dest="captions", action="store_true",
                   default=None)
    r.add_argument("--no-captions", dest="captions", action="store_false")
    r.add_argument("--xref", dest="xref", action="store_true", default=None)
    r.add_argument("--no-xref", dest="xref", action="store_false")
    r.add_argument("--dry-run", action="store_true")
    r.add_argument("--workdir")
    r.add_argument("--tree-json")
    r.add_argument("--tree-txt")
    r.add_argument("--overrides",
                   help="人工/AI 拍板的层级表 json（propose 产出后编辑它）")
    r.set_defaults(func=cmd_run)

    # propose
    pr = sub.add_parser("propose", help="导出待拍板清单（工具判不准的地方）")
    pr.add_argument("infile")
    pr.add_argument("-o", "--out", help="待拍板清单 json 输出路径")
    pr.add_argument("--report", help="人可读报告输出路径")
    pr.add_argument("--workdir")
    pr.set_defaults(func=cmd_propose)

    # inspect
    i = sub.add_parser("inspect", help="只侦察不改")
    i.add_argument("infile")
    i.add_argument("--workdir")
    i.add_argument("--tree-txt")
    i.set_defaults(func=cmd_inspect)

    # spec
    s = sub.add_parser("spec", help="生成 spec 样例")
    s.add_argument("--init", required=True, help="写出的 json 路径")
    s.add_argument("--mode", choices=["default", "template", "requirement"],
                   default="template")
    s.set_defaults(func=cmd_spec)

    # styles（样式确认单）
    sg = sub.add_parser("styles", help="S1 样式确认单（只出单，不套用）")
    sg.add_argument("infile")
    sg.add_argument("--spec", help="完整规范 json（优先）")
    sg.add_argument("--template", help="参考模板 docx")
    sg.add_argument("--requirement", help="文字规范 json")
    sg.add_argument("--out", help="导出的可编辑 spec json 路径")
    sg.add_argument("--html", help="HTML 确认页路径")
    sg.set_defaults(func=cmd_styles)

    # lint（正文校验）
    lt = sub.add_parser("lint", help="S5 正文校验（只出改动事项单）")
    lt.add_argument("infile")
    lt.add_argument("--spec", help="完整规范 json（优先）")
    lt.add_argument("--template", help="参考模板 docx")
    lt.add_argument("--requirement", help="文字规范 json")
    lt.add_argument("--out", help="CRS json 路径")
    lt.add_argument("--html", help="CRS 可勾选页路径")
    lt.set_defaults(func=cmd_lint)

    # revise（修订模式 + 批注）
    rv = sub.add_parser("revise", help="S6 修订落盘（只改已批准的项）")
    rv.add_argument("infile")
    rv.add_argument("--crs", required=True, help="改动事项单 json")
    rv.add_argument("--approved", help="_approved.json（覆盖 decision）")
    rv.add_argument("-o", "--out", help="输出 docx（默认 <名>_修订.docx）")
    rv.add_argument("--receipt", help="改动回执 md 路径")
    rv.add_argument("--author", default="docx-format 校对", help="修订/批注作者")
    rv.add_argument("--dry", action="store_true", help="干跑，不落盘")
    rv.set_defaults(func=cmd_revise)

    # polish（S4 六条硬要求）
    po = sub.add_parser("polish", help="S4 排版体检（六条硬要求，只出单）")
    po.add_argument("infile")
    _add_common_spec(po)
    po.add_argument("--pages", help="paginate.py 的量页 json（不给则自动量）")
    po.add_argument("--no-render", action="store_true",
                    help="不量页，只做静态检查（①⑥）")
    po.add_argument("--out", help="CRS json 路径")
    po.add_argument("--html", help="CRS 可勾选页路径")
    po.add_argument("--apply", action="store_true", help="执行结构类改动")
    po.add_argument("--crs", help="--apply 时的改动事项单 json")
    po.add_argument("--approved", help="_approved.json")
    po.add_argument("--out-docx", dest="out_docx", help="排版产物路径")
    po.add_argument("--receipt", help="回执 md 路径")
    po.add_argument("--dry", action="store_true")
    po.set_defaults(func=cmd_polish)

    # captions（S3 题注补插）
    cp = sub.add_parser("captions", help="S3 题注补插（只出单）")
    cp.add_argument("infile")
    _add_common_spec(cp)
    cp.add_argument("--out", help="CRS json 路径")
    cp.add_argument("--html", help="CRS 可勾选页路径")
    cp.add_argument("--apply", action="store_true")
    cp.add_argument("--crs", help="--apply 时的改动事项单 json")
    cp.add_argument("--approved", help="_approved.json")
    cp.add_argument("--out-docx", dest="out_docx", help="产物路径")
    cp.add_argument("--receipt", help="回执 md 路径")
    cp.add_argument("--dry", action="store_true")
    cp.set_defaults(func=cmd_captions)

    # facts（S5b 事实核查抽取）
    fa = sub.add_parser("facts", help="S5b 事实核查：抽原子断言 → 落 CRS")
    fa.add_argument("infile")
    _add_common_spec(fa)
    fa.add_argument("--out", help="待核查事实 json 路径")
    fa.add_argument("--md", help="人读核查表 md 路径")
    fa.add_argument("--verdicts", help="判定回填 json")
    fa.add_argument("--to-crs", dest="to_crs", help="产出事实核查 CRS json")
    fa.add_argument("--html", help="CRS 可勾选页路径")
    fa.add_argument("--date", help="核查日期（默认今天）")
    fa.set_defaults(func=cmd_facts)

    # fields
    f = sub.add_parser("fields", help="重算全部域")
    f.add_argument("infile")
    f.add_argument("-o", "--out")
    f.set_defaults(func=cmd_fields)

    # verify
    v = sub.add_parser("verify", help="校验")
    v.add_argument("infile")
    v.add_argument("--template")
    v.set_defaults(func=cmd_verify)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
