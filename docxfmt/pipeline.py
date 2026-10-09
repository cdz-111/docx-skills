# -*- coding: utf-8 -*-
"""
pipeline.py —— 主流水线：解包 → 移植样式 → 建树 → 渲染 → 题注/引用
                → 目录/图表目录 → 分页符/分节 → 页面/制表位 → 打包。

顺序有讲究，不能乱：
    1. 拆段（图片+题注同段）  —— 会改段落序号，**必须早于建树**
    2. 样式移植              —— 必须在建 name→id 表之前，否则渲染查不到样式
    3. 建树 + 渲染           —— 给每个段落套上样式
    4. 题注 + 交叉引用       —— 需要「所属章节」信息（来自树）
    5. 目录 / 图表目录       —— 必须在渲染之后，题注套好样式才收集得到
    6. 分页符 / 分节
    7. 页面设置 + 制表位缩放 —— snapshot 必须在 apply_page 之前
    8. 落盘 + 打包
"""
import io
import json
import os
import shutil
import sys
import zipfile
from collections import Counter

from lxml import etree

from . import __version__, qn
from . import captions as CAP
from . import layout, render, toc
from .docxio import Docx, pack, read_style_names, unpack
from .recognize import parse, stats, tree_report
from .spec import Spec
from .styles import NAME_ALIASES, StyleTable, patch_heading_indent, transplant


def _out(s):
    print(s, flush=True)


def _load_requirement(spec):
    """style_from="requirement"：从文字规范 json 读样式名映射。

    requirement 为空 → 用内置默认模板 spec/default.json（用户没给模板时的兜底）。
    """
    from .spec import DEFAULT_REQUIREMENT
    path = spec["source"].get("requirement") or DEFAULT_REQUIREMENT
    if not os.path.exists(path):
        raise SystemExit("[spec] 找不到文字规范文件：%s" % path)
    with open(path, "r", encoding="utf-8") as f:
        req = json.load(f)
    st = req.get("styles") or {}
    m = st.get("map") or req.get("style_map") or {}
    if m:
        spec["styles"]["map"] = m
    for key in ("new_styles", "geometry"):
        v = st.get(key) or req.get(key)
        if v:
            spec["styles"][key] = v
    # 默认模板里其余声明式字段（numbering/toc/tables/spacing/lint…）一并带入
    for key in ("numbering", "toc", "figure_list", "table_list", "captions",
                "xref", "pagebreak", "section", "tables", "spacing", "lint"):
        if key in req and req[key]:
            spec.setdefault(key, {})
            if isinstance(req[key], dict):
                spec[key] = {**spec[key], **req[key]}
    return spec


def _all_style_names(spec):
    """要移植/确保存在的样式 w:name 全集。"""
    names = []
    m = spec["styles"]["map"]
    for v in m.values():
        if isinstance(v, dict):
            names += [x for x in v.values() if x]
        elif v:
            names.append(v)
    names += list((spec["styles"].get("geometry") or {}).keys())
    names += list((spec["styles"].get("new_styles") or {}).keys())
    seen, out = set(), []
    for n in names:
        if n and n not in seen:
            seen.add(n)
            out.append(n)
    return out


def run(src, out, spec, workdir=None, tree_json=None, tree_txt=None,
        dry_run=False, overrides=None, log=_out):
    workdir = workdir or os.path.join(
        os.path.dirname(os.path.abspath(out)), "_unpacked")
    spec = Spec(spec) if not isinstance(spec, Spec) else spec
    if spec["source"].get("style_from") == "requirement":
        spec = _load_requirement(spec)

    log("== docxfmt %s ==" % __version__)
    log("[in ] %s" % src)

    # ---- 1. 解包
    unpack(src, workdir)
    doc = Docx(workdir)
    sid2name = {sid: nm for sid, nm in
                ((s.get(qn("styleId")),
                  (s.find(qn("name")).get(qn("val"))
                   if s.find(qn("name")) is not None else ""))
                 for s in doc.sty.getroot().iter(qn("style")))}

    # ---- 2. 拆段（必须早于建树）
    n_split = CAP.split_caption_paras(doc.root)
    if n_split:
        log("[split] 图片与题注同段，拆出 %d 个题注段" % n_split)
    paras = doc.paras()

    # ---- 3. 样式来源
    style_from = spec["source"].get("style_from", "template")
    if style_from == "template":
        tpl = spec["source"].get("template")
        if not tpl or not os.path.exists(tpl):
            raise SystemExit("[spec] style_from=template 但模板不存在：%s" % tpl)
        with zipfile.ZipFile(tpl) as z:
            troot = etree.fromstring(z.read("word/styles.xml"))
        over, added, remap, missing = transplant(
            doc.sty.getroot(), troot, _all_style_names(spec))
        log("[styles] 移植模板样式：覆盖 %d，新增 %d，basedOn 重映射 %d"
            % (over, added, remap))
        if missing:
            log("[styles] 模板里没有（已按别名找过）：%s" % missing)
        n = patch_heading_indent(doc.sty.getroot())
        if n:
            log("[styles] 标题样式首行缩进兜底 %d 个" % n)
    elif style_from == "requirement":
        src = spec["source"].get("requirement")
        log("[styles] 按文字规范排：%s" % (src or "spec/default.json（内置默认模板）"))
    elif style_from == "keep":
        log("[styles] style_from=keep：只做内容层，不动样式")

    table = StyleTable(doc.sty.getroot())
    made = render.create_new_styles(spec, table)
    if made:
        log("[styles] 新建样式 %s" % made)
    ngeom = render.apply_geometry(spec, table)
    if ngeom:
        log("[styles] 几何覆盖 %d 项" % ngeom)

    # 渲染用的 name -> styleId
    name_to_id = {}
    for nm, st in table.by_name.items():
        sid = st.get(qn("styleId"))
        if sid and nm not in name_to_id:
            name_to_id[nm] = sid

    # ---- 4. 建树 + 渲染
    skip_kw = spec["numbering"].get("skip_keywords") or []
    root, index, cover_end = parse(doc.root, sid2name,
                                   skip_keywords=skip_kw)

    # ---- 4.5 人工/AI 拍板覆盖（propose -> 判断 -> overrides 喂回来）
    if overrides:
        from .propose import apply_overrides
        n_ov = apply_overrides(root, overrides)
        log("[override] 应用人工/AI 判断 %d 条" % n_ov)
        # 改完重新打章节标记
        from .recognize import mark_sections
        mark_sections(root, skip_kw)

    log("[parse] 文档树：封面段 [0,%d) 保留原样" % cover_end)
    st = stats(root)
    log("[parse] " + "  ".join("%s=%d" % kv for kv in st.most_common()
                              if not kv[0].startswith("heading")
                              or kv[0][7:].isdigit()))
    stt = Counter()
    if not dry_run:
        miss = render.render(root, spec, table, stt)
        log("[render] " + "  ".join("%s=%d" % kv for kv in stt.most_common()
                                    if kv[0] != "_缺样式"))
        if miss:
            log("[render] ⚠ 方案里有但文档没有的样式：%s" % dict(miss))

    # ---- 5. 题注 + 交叉引用
    if spec.enabled("captions") and not dry_run:
        CAP.process(root, doc.root, paras, table, spec, log=log)

    # ---- 6. 公式编号
    if spec["numbering"].get("formula", "none") != "none" and not dry_run:
        CAP.convert_formula_numbers(
            root, paras, table,
            style_name=spec.style_of("formula_number"),
            scheme=spec["numbering"]["formula"],
            sep=spec.sep, log=log)

    # ---- 7. 分页符 / 分节
    if not dry_run:
        npb, nkn = layout.apply_page_breaks(
            root, spec["pagebreak"].get("before_heading") or [],
            spec["pagebreak"].get("keep_with_next", True))
        if npb or nkn:
            log("[break] 分页符 %d 处，keepNext %d 处" % (npb, nkn))
        lv = spec["section"].get("insert") or []
        lv = [int(x) for x in lv] if not isinstance(lv, int) else [lv]
        if lv or spec["section"].get("landscape_nodes"):
            layout.apply_sections(
                doc.root, lv, spec["section"].get("landscape_nodes") or [],
                margin=spec.get("page", {}).get("margin"), log=log)

    # ---- 8. 目录 / 图目录 / 表目录
    if not dry_run:
        n = toc.insert_all(doc.root, spec, table, log=log)
        if n:
            log("[toc] 目录类共插入 %d 段" % n)

    # ---- 9. 页面 + 制表位（snapshot 必须在 apply_page 之前）
    if not dry_run:
        old = layout.snapshot_widths(doc.root)
        npage = layout.apply_page(
            doc.root,
            margin=spec.get("page", {}).get("margin"),
            orientation=spec["section"].get("orientation", "keep"),
            first_page_diff=spec["section"].get("first_page_diff", False),
            odd_even=spec["section"].get("odd_even", False))
        log("[page] %d 个节的页面设置已应用（横向节按 keep 保留方向）" % npage)
        np_, nt = layout.rescale_tabs(doc.root, old)
        if np_ or nt:
            log("[tabs] 缩放 %d 个段落、%d 个制表位" % (np_, nt))

    # ---- 10. 导出树
    if tree_json or tree_txt:
        for p in (tree_json, tree_txt):
            if p:
                os.makedirs(os.path.dirname(os.path.abspath(p)), exist_ok=True)
    if tree_json:
        with open(tree_json, "w", encoding="utf-8") as f:
            json.dump(root.to_dict(), f, ensure_ascii=False, indent=1)
    if tree_txt:
        tree_report(root, tree_txt)
        log("[dump] 树 -> %s ；报告 -> %s" % (tree_json, tree_txt))

    # ---- 11. 落盘 + 打包
    if dry_run:
        log("[dry-run] 未落盘")
        return root
    doc.save()
    n = pack(workdir, out)
    log("[pack] %d 个部件 -> %s (%d bytes)" % (n, out, os.path.getsize(out)))
    shutil.rmtree(workdir, ignore_errors=True)
    return root
