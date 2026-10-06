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
    """命令行 -> spec dict。--spec 提供基底，快捷选项在其上覆盖。"""
    base = {}
    if args.spec:
        with open(args.spec, "r", encoding="utf-8") as f:
            base = json.load(f)
    if args.template:
        base.setdefault("source", {})
        base["source"]["style_from"] = "template"
        base["source"]["template"] = args.template
    if args.requirement:
        base.setdefault("source", {})
        base["source"]["style_from"] = "requirement"
        base["source"]["requirement"] = args.requirement
    if args.caption_scheme:
        base.setdefault("numbering", {})["caption"] = args.caption_scheme
    if args.formula_scheme:
        base.setdefault("numbering", {})["formula"] = args.formula_scheme
    if args.number_sep:
        base.setdefault("numbering", {})["sep"] = args.number_sep
    if args.pagebreak_levels is not None:
        base.setdefault("pagebreak", {})["before_heading"] = \
            [int(x) for x in _kvlist(args.pagebreak_levels)]
    for key, val in (("toc", args.toc), ("figure_list", args.figure_list),
                     ("table_list", args.table_list),
                     ("captions", args.captions), ("xref", args.xref)):
        if val is not None:
            base.setdefault(key, {})["enabled"] = val
    if args.toc_levels:
        base.setdefault("toc", {})["levels"] = args.toc_levels
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
    # win32com 只装在部分解释器里（常见于 anaconda）。
    # 找「装了的」而不是「猜路径里的用户名」——硬编码别人机器上的
    # 用户名既不通用也不该进仓库。
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
    s.add_argument("--mode", choices=["template", "requirement"],
                   default="template")
    s.set_defaults(func=cmd_spec)

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
