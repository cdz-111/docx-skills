#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""格式与编号对齐执行器 —— CR-101/102/201/202/203/204/205

一次原子完成「公式编号对齐」+「样式/字体修复」两批改动：它们都落在同一批
段落的 pPr / rPr 上，分头执行会互相覆盖 pPrChange，故必须同一执行器做完。

  CR-101 (eq_align)     第2章 15 个编号段补制表位 + 公式前 Tab
  CR-102 (eq_align)     第3章 2 个编号段补公式前 Tab（其一并 jc center->left）
  CR-201 (format_scrub) 25 个标题段清直接格式（等线/kern/硬字体/硬字号/硬加粗）
  CR-202 (style_swap)   2.2.3~2.2.5 由 Normal (Web) 归位 heading 3
  CR-203 (style_def)    heading 2 样式定义去 eastAsiaTheme（等线->宋体继承）
  CR-204 (style_swap)   图 3-1/3-3 题注由 Normal (Web) 归位 caption
  CR-205 (format_scrub) 36 段非 Times New Roman 西文回落到 TNR

口径：
  * 只改 word/document.xml 与 word/styles.xml（仅 CR-203 时），其余部件逐字节复制
  * 默认保留修订痕迹（w:pPrChange / w:rPrChange / w:ins / w:del）；--no-trace 关闭
  * 字符串级替换，非改动区逐字节不变
  * 硬断言：各类命中数不符预期即拒绝执行（退码 4）
  * 批注挂在每批首个命中的段落上

用法:
  fmt_eq_fix.py <in.docx> <out.docx> <approved.json> [--dry] [--no-trace]
                [--cr203 a|b] [--author NAME]

退出码: 0 成功 / 2 验收失败 / 3 无批准 / 4 数量断言失败 / 5 指纹不符 / 6 用法错误
"""
import sys
import os
import re
import io
import json
import time
import hashlib
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
def q(t): return W + t
def m(t): return M + t

SUPPORTED = ("eq_align", "format_scrub", "style_swap", "style_def")
DEFAULT_AUTHOR = "docx-format 校对"

TAGP = re.compile(r'<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>')
RUNPAT = re.compile(r'<w:r(?:\s[^>]*)?>.*?</w:r>|<w:r(?:\s[^>]*)?/>', re.S)
RPRPAT = re.compile(r'<w:rPr>.*?</w:rPr>', re.S)
PPRPAT = re.compile(r'<w:pPr>.*?</w:pPr>', re.S)
RFONTSPAT = re.compile(r'<w:rFonts[^>]*/>')
TPAT = re.compile(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', re.S)

# 要清掉的西文/主题字体属性（保留 hint / eastAsia）
RF_DROP = ("ascii", "hAnsi", "cs", "eastAsiaTheme", "asciiTheme",
           "hAnsiTheme", "cstheme")

TABS = ('<w:tabs><w:tab w:val="center" w:pos="4153"/>'
        '<w:tab w:val="right" w:pos="8306"/></w:tabs>')


# ------------------------------------------------------------------ 工具
def now_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def md5_file(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def split_paras(xml):
    spans, depth, st = [], 0, None
    for mo in TAGP.finditer(xml):
        s = mo.group(0)
        if s.endswith("/>"):
            if depth == 0:
                spans.append((mo.start(), mo.end()))
        elif s == "</w:p>":
            depth -= 1
            if depth <= 0:
                if st is not None:
                    spans.append((st, mo.end()))
                st, depth = None, 0
        else:
            if depth == 0:
                st = mo.start()
            depth += 1
    return spans


PPR_SCAN = re.compile(r'</?w:pPr>')


def ppr_span(seg):
    """外层 w:pPr 的 (start, end)。
    必须深度计数——w:pPrChange 内部还嵌着一份记录旧值的 <w:pPr>…</w:pPr>，
    直接用 seg.find("</w:pPr>") 会命中内层那个（曾因此把 Tab 插进档案区）。"""
    s = seg.find("<w:pPr>")
    if s < 0:
        return None
    depth = 0
    for mo in PPR_SCAN.finditer(seg, s):
        if mo.group(0) == "<w:pPr>":
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                return (s, mo.end())
    return None


def get_ppr(seg):
    sp = ppr_span(seg)
    return seg[sp[0]:sp[1]] if sp else None


def run_text(run):
    return "".join(m.group(1) for m in TPAT.finditer(run))


def has_latin(run):
    return bool(re.search(r"[A-Za-z]", run_text(run)))


def clean_rfonts(tag):
    """清掉西文/主题字体属性；若一个都没剩则整个 rFonts 删掉。"""
    pairs = re.findall(r'(w:[A-Za-z]+)="([^"]*)"', tag)
    kept = [(k, v) for k, v in pairs if k.split(":")[1] not in RF_DROP]
    if not any(k == "w:hint" for k, _ in kept) and \
       not any(k == "w:eastAsia" for k, _ in kept):
        kept.append(("w:hint", "eastAsia"))
    if not kept:
        return ""
    return "<w:rFonts " + " ".join('%s="%s"' % kv for kv in kept) + "/>"


def is_non_tnr(tag):
    vals = re.findall(r'w:(?:ascii|hAnsi|cs|asciiTheme|hAnsiTheme|cstheme)="([^"]*)"', tag)
    if not vals:
        return False
    return any(v.strip() != "Times New Roman" for v in vals)


def scrub_rpr(rpr, drop_kern=False, drop_size=False, drop_bold=False):
    """清理一个 <w:rPr>，返回 (新 rPr, 是否有变化)。"""
    body = rpr[len("<w:rPr>"):-len("</w:rPr>")]
    newbody = RFONTSPAT.sub(lambda mo: clean_rfonts(mo.group(0)) or "", body)
    if drop_kern:
        newbody = re.sub(r'<w:kern[^>]*/>', "", newbody)
    if drop_size:
        newbody = re.sub(r'<w:sz[^>]*/>', "", newbody)
        newbody = re.sub(r'<w:szCs[^>]*/>', "", newbody)
    if drop_bold:
        newbody = re.sub(r'<w:b/>|<w:bCs/>', "", newbody)
    new = "<w:rPr>" + newbody + "</w:rPr>"
    return new, new != rpr


def add_ppr_change(new_ppr, old_ppr, ctx):
    inner = old_ppr[len("<w:pPr>"):-len("</w:pPr>")]
    inner = re.sub(r'<w:pPrChange.*?</w:pPrChange>', "", inner, flags=re.S)
    ch = ('<w:pPrChange w:id="%d" w:author="%s" w:date="%s">'
          '<w:pPr>%s</w:pPr></w:pPrChange>'
          % (ctx.nrev(), ctx.author, ctx.date, inner))
    return new_ppr[:-len("</w:pPr>")] + ch + "</w:pPr>"


def add_rpr_change(new_rpr, old_rpr, ctx):
    inner = old_rpr[len("<w:rPr>"):-len("</w:rPr>")]
    inner = re.sub(r'<w:rPrChange.*?</w:rPrChange>', "", inner, flags=re.S)
    ch = ('<w:rPrChange w:id="%d" w:author="%s" w:date="%s">'
          '<w:rPr>%s</w:rPr></w:rPrChange>'
          % (ctx.nrev(), ctx.author, ctx.date, inner))
    return new_rpr[:-len("</w:rPr>")] + ch + "</w:rPr>"


def cref_run(cid):
    return ('<w:r><w:rPr><w:rStyle w:val="38"/><w:rFonts w:hint="eastAsia"/>'
            '<w:sz w:val="24"/><w:szCs w:val="22"/></w:rPr>'
            '<w:commentReference w:id="%d"/></w:r>' % cid)


# ------------------------------------------------------------------ 上下文
class Ctx(object):
    def __init__(self, author, date, trace):
        self.author = author
        self.date = date
        self.trace = trace
        self.bk = 1
        self.rev = 1
        self.cid = 1
        self.comments = []
        self.log = []
        self.hits = {}

    def nrev(self):
        v = self.rev
        self.rev += 1
        return v

    def ncid(self, text):
        v = self.cid
        self.cid += 1
        self.comments.append((v, text))
        return v

    def bump(self, key, n=1):
        self.hits[key] = self.hits.get(key, 0) + n


# ------------------------------------------------------------------ 段落变换
def apply_eq_align(seg, ctx, note_cid):
    """补制表位 / jc->left / 公式前 Tab / 编号前空格 run -> Tab。"""
    old_ppr = get_ppr(seg)
    if old_ppr is None:
        return seg, 0
    ppr = old_ppr
    n = 0
    # 1. 制表位
    if "<w:tabs>" not in ppr:
        ins_at = len("<w:pPr>")
        mo = re.search(r'<w:(autoSpaceDN|bidi|spacing|ind|jc|outlineLvl|rPr|sectPr|pPrChange)\b', ppr)
        if mo:
            ins_at = mo.start()
        ppr = ppr[:ins_at] + TABS + ppr[ins_at:]
        n += 1
    # 2. jc center -> left
    if '<w:jc w:val="center"/>' in ppr:
        ppr = ppr.replace('<w:jc w:val="center"/>', '<w:jc w:val="left"/>', 1)
        n += 1
    # 3. 公式前缺 Tab？（外层 </w:pPr> 与首个 <m:oMath 之间）
    _sp = ppr_span(seg)
    k2 = seg.find("<m:oMath")
    pre_tab = False
    if _sp and 0 <= _sp[1] < k2:
        pre_tab = "<w:tab/>" in seg[_sp[1]:k2]
    if not pre_tab:
        n += 1
    # 4. 编号前的空格 run -> Tab（只在最后一个 </m:oMath> 之后找）
    cut = seg.rfind("</m:oMath>")
    tail = seg[cut:] if cut >= 0 else seg
    sm = re.search(r'<w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t xml:space="preserve">(\s+)</w:t></w:r>', tail)
    if sm:
        n += 1

    if n == 0:
        return seg, 0

    if ctx.trace:
        ppr = add_ppr_change(ppr, old_ppr, ctx)
    seg = seg.replace(old_ppr, ppr, 1)

    # 公式前 Tab
    if not pre_tab:
        if ctx.trace:
            run = ('<w:ins w:id="%d" w:author="%s" w:date="%s">'
                   '<w:r><w:tab/></w:r></w:ins>'
                   % (ctx.nrev(), ctx.author, ctx.date))
        else:
            run = '<w:r><w:tab/></w:r>'
        _sp2 = ppr_span(seg)
        i = _sp2[1] if _sp2 else seg.find(">") + 1
        seg = seg[:i] + run + seg[i:]

    # 编号前空格 run -> Tab
    if sm:
        cut2 = seg.rfind("</m:oMath>")
        head, tail2 = seg[:cut2], seg[cut2:]
        mo = re.search(r'<w:r>((?:<w:rPr>.*?</w:rPr>)?)<w:t xml:space="preserve">(\s+)</w:t></w:r>', tail2)
        if mo:
            if ctx.trace:
                oldtxt = mo.group(2)
                new = ('<w:del w:id="%d" w:author="%s" w:date="%s">'
                       '<w:r>%s<w:delText xml:space="preserve">%s</w:delText></w:r>'
                       '</w:del><w:ins w:id="%d" w:author="%s" w:date="%s">'
                       '<w:r><w:tab/></w:r></w:ins>'
                       % (ctx.nrev(), ctx.author, ctx.date, mo.group(1), oldtxt,
                          ctx.nrev(), ctx.author, ctx.date))
            else:
                new = "<w:r><w:tab/></w:r>"
            i = cut2 + mo.start()
            j = cut2 + mo.end()
            seg = seg[:i] + new + seg[j:]

    if note_cid is not None:
        seg = attach_note(seg, note_cid)
    return seg, n


def apply_substyle(seg, ctx, note_cid, kind):
    """kind: 'heading' (CR-201) / 'caption' (CR-204) / 'font' (CR-205)。
    注意：段 2996/5215 没有 pPr，仍必须处理 run 级。"""
    old_ppr = get_ppr(seg)
    ppr = old_ppr
    n = 0

    # ---- 段落标记 rPr ----
    if old_ppr is not None:
        prm = RPRPAT.search(ppr)
        if kind == "heading":
            if prm:
                new, ch = scrub_rpr(prm.group(0), True, True, True)
                if ch:
                    ppr = ppr.replace(prm.group(0), new, 1)
                    n += 1
        elif kind == "caption":
            if prm:
                new, ch = scrub_rpr(prm.group(0), False, False, True)
                if ch:
                    ppr = ppr.replace(prm.group(0), new, 1)
                    n += 1
        else:  # font
            if prm and is_non_tnr(prm.group(0)):
                new, ch = scrub_rpr(prm.group(0), False, False, False)
                if ch:
                    ppr = ppr.replace(prm.group(0), new, 1)
                    n += 1
                    ctx.bump("cr205_ppr")

    # ---- run 级 ----
    pieces = []
    prev = 0
    for mo in RUNPAT.finditer(seg):
        pieces.append(seg[prev:mo.start()])
        run = mo.group(0)
        rprm = RPRPAT.search(run)
        done = run
        if rprm and "<w:delText" not in run:
            new, ch = rprm.group(0), False
            if kind == "heading":
                new, ch = scrub_rpr(rprm.group(0), True, True, True)
            elif kind == "caption":
                new, ch = scrub_rpr(rprm.group(0), False, False, True)
            elif has_latin(run) and is_non_tnr(rprm.group(0)):
                new, ch = scrub_rpr(rprm.group(0), False, False, False)
            if ch:
                if ctx.trace:
                    new = add_rpr_change(new, rprm.group(0), ctx)
                done = run.replace(rprm.group(0), new, 1)
                n += 1
                if kind == "font":
                    ctx.bump("cr205_run")
        pieces.append(done)
        prev = mo.end()
    pieces.append(seg[prev:])
    seg = "".join(pieces)

    if n == 0:
        return seg, 0
    if old_ppr is not None:
        if ctx.trace:
            ppr = add_ppr_change(ppr, old_ppr, ctx)
        if ppr != old_ppr:
            seg = seg.replace(old_ppr, ppr, 1)
    if note_cid is not None:
        seg = attach_note(seg, note_cid)
    return seg, n


def apply_style_swap(seg, ctx, note_cid, new_style, clean_web, kind):
    """改 pStyle（+可选清网页残留），再清直接格式。"""
    old_ppr = get_ppr(seg)
    if old_ppr is None:
        return seg, 0
    ppr = re.sub(r'<w:pStyle w:val="[^"]*"/>',
                 '<w:pStyle w:val="%s"/>' % new_style, old_ppr, count=1)
    if ppr == old_ppr:
        return seg, 0
    n = 1
    if clean_web:
        p2 = re.sub(r'<w:spacing[^>]*/>', "", ppr)
        p2 = re.sub(r'<w:ind[^>]*/>', "", p2)
        if p2 != ppr:
            ppr = p2
            n += 1
    # 段落标记 rPr 清理（与 clean_web 无关，CR-204 也要清 Arial）
    prm = RPRPAT.search(ppr)
    if prm:
        if kind == "heading":
            new, ch = scrub_rpr(prm.group(0), True, True, True)
        else:
            new, ch = scrub_rpr(prm.group(0), False, False, True)
        if ch:
            ppr = ppr.replace(prm.group(0), new, 1)
            n += 1
    # run 级清理
    pieces = []
    prev = 0
    for mo in RUNPAT.finditer(seg):
        pieces.append(seg[prev:mo.start()])
        run = mo.group(0)
        rprm = RPRPAT.search(run)
        done = run
        if rprm and "<w:delText" not in run:
            if kind == "heading":
                new, ch = scrub_rpr(rprm.group(0), True, True, True)
            else:
                new, ch = scrub_rpr(rprm.group(0), False, False, True)
            if ch:
                if ctx.trace:
                    new = add_rpr_change(new, rprm.group(0), ctx)
                done = run.replace(rprm.group(0), new, 1)
                n += 1
        pieces.append(done)
        prev = mo.end()
    pieces.append(seg[prev:])
    seg = "".join(pieces)

    if ctx.trace:
        ppr = add_ppr_change(ppr, old_ppr, ctx)
    seg = seg.replace(old_ppr, ppr, 1)
    if note_cid is not None:
        seg = attach_note(seg, note_cid)
    return seg, n


def attach_note(seg, cid):
    """把批注区间挂到段首（reference 落在 </w:p> 前，避开任何域结果区）。"""
    if not seg.endswith("</w:p>"):
        return seg
    body = seg[:-len("</w:p>")]
    i = body.find(">") + 1          # <w:p ...> 结束处
    return (body[:i] + '<w:commentRangeStart w:id="%d"/>' % cid + body[i:]
            + '<w:commentRangeEnd w:id="%d"/>' % cid + cref_run(cid) + "</w:p>")


# ------------------------------------------------------------------ styles
def patch_heading2(styles, old_theme):
    """heading 2（styleId=3）去 eastAsiaTheme。返回 (新styles, 命中数)"""
    mo = re.search(r'<w:style [^>]*w:styleId="3"[^>]*>.*?</w:style>', styles, re.S)
    if not mo:
        return styles, 0
    blk = mo.group(0)
    new = blk.replace('<w:rFonts w:eastAsiaTheme="minorEastAsia"/>',
                      '<w:rFonts w:hint="eastAsia"/>')
    new = new.replace('<w:rFonts w:hint="eastAsia" w:eastAsiaTheme="minorEastAsia"/>',
                      '<w:rFonts w:hint="eastAsia"/>')
    # 通用：删掉 rFonts 里的 eastAsiaTheme
    new = re.sub(r'(<w:rFonts[^>]*?)\s+w:eastAsiaTheme="[^"]*"', r'\1', new)
    if new == blk:
        return styles, 0
    return styles.replace(blk, new, 1), 1


# ------------------------------------------------------------------ 主流程
def main(argv):
    if len(argv) < 4:
        sys.stderr.write(__doc__)
        return 6
    inp, out, appr = argv[1], argv[2], argv[3]
    dry = "--dry" in argv
    trace = "--no-trace" not in argv
    cr203 = None
    if "--cr203" in argv:
        cr203 = argv[argv.index("--cr203") + 1].strip().lower()
    author = argv[argv.index("--author") + 1] if "--author" in argv else DEFAULT_AUTHOR
    if not os.path.isfile(inp):
        sys.stderr.write("找不到输入: %s\n" % inp)
        return 6
    if not os.path.isfile(appr):
        sys.stderr.write("缺少批准文件（CRS 门禁）: %s\n" % appr)
        return 3

    ap = json.load(io.open(appr, encoding="utf-8"))
    items = {i["id"]: i for i in ap.get("items", [])}
    todo = {k for k, v in items.items()
            if v.get("decision") in ("approved", "modified")}
    if not todo:
        sys.stderr.write("批准文件无 approved/modified 项\n")
        return 3

    cur = md5_file(inp)
    if ap.get("target_md5") and cur != ap["target_md5"]:
        sys.stderr.write("指纹不符: %s != %s\n" % (cur, ap["target_md5"]))
        return 5

    z = zipfile.ZipFile(inp)
    doc = z.read("word/document.xml").decode("utf-8")
    styles = z.read("word/styles.xml").decode("utf-8")
    comments = z.read("word/comments.xml").decode("utf-8")
    cx = z.read("word/commentsExtended.xml").decode("utf-8")
    root = etree.fromstring(doc.encode("utf-8"))
    spans = split_paras(doc)
    n_lxml = sum(1 for _ in root.iter(q("p")))
    assert len(spans) == n_lxml, "段落口径不一致 %d vs %d" % (len(spans), n_lxml)

    date = now_iso()
    ctx = Ctx(author, date, trace)
    # id 起点
    bks = [int(x) for x in re.findall(r'<w:bookmarkStart w:id="(\d+)"', doc)]
    rvs = [int(x) for x in re.findall(r'<w:(?:ins|del) w:id="(\d+)"', doc)]
    cds = [int(x) for x in re.findall(r'<w:comment w:id="(\d+)"', comments)]
    ctx.bk = (max(bks) + 1) if bks else 1
    ctx.rev = (max(rvs) + 1) if rvs else 1
    ctx.cid = (max(cds) + 1) if cds else 1
    rev0 = ctx.rev
    cid0 = ctx.cid

    # ---- 收集各项段号 ----
    def paras(cid_):
        it = items.get(cid_)
        if not it or it.get("decision") not in ("approved", "modified"):
            return []
        return list(it["where"]["para"] or [])

    eq101, eq102 = paras("CR-101"), paras("CR-102")
    h201, s202 = paras("CR-201"), paras("CR-202")
    c204, f205 = paras("CR-204"), paras("CR-205")

    # ---- 批注 ----
    notes = {}
    if eq101 or eq102:
        notes["eq"] = ctx.ncid(
            "【CR-101/102】公式编号右对齐（共 17 段）：文档用「制表位法」——制表位 "
            "center@4153（文本区正中，放公式）+ right@8306（右边距，放编号），文本区宽 "
            "8306 twips 与页边距 left=right=1800 完全吻合。第2章 15 段漏设制表位且整行居中，"
            "第3章 2 段（3-26/3-62）缺公式前的 Tab；本处补齐，效果与其余 73 段同构。")
    if h201:
        notes["h201"] = ctx.ncid(
            "【CR-201】标题直接格式清理（共 25 段）：这些标题被塞入了直接格式——"
            "w:eastAsiaTheme=\"minorEastAsia\"（中文渲染为 Word 主题字体「等线」，"
            "即「一眼 AI」观感的主因）、w:kern 字距、硬写字体与字号、硬加粗。"
            "此处清除，恢复与第 1 章一致的「纯样式继承」（中文宋体、西文 Times New Roman）。")
    if s202:
        notes["s202"] = ctx.ncid(
            "【CR-202】2.2.3~2.2.5 标题样式归位（共 3 段）：原为 Normal (Web)（网页粘贴残留），"
            "改为与同章 2.3.1/2.4.1 一致的 heading 3；网页残留的段间距、缩进、硬字号、"
            "硬加粗一并清除，保留 keepNext/keepLines/outlineLvl。")
    if c204:
        notes["c204"] = ctx.ncid(
            "【CR-204】图 3-1 / 图 3-3 题注样式归位（共 2 段）：原为 Normal (Web)，"
            "硬写 Arial/宋体/加粗；改为与全文其余 302 个图/表题注一致的 caption 样式。")
    if f205:
        notes["f205"] = ctx.ncid(
            "【CR-205】非 Times New Roman 西文回落（共 72 处）：①「Times New Roman Regular」"
            "是无效字体名（正确写法没有 Regular），Word 找不到会静默回落，集中在参考文献；"
            "② w:ascii/ hAnsi 被指定为宋体/仿宋，西文渲染成中文字形；③ 少量 Arial；"
            "④ 非法的主题值 asciiTheme=\"minorEastAsia\"（合法值只有 minorHAnsi/majorHAnsi）。"
            "上述直接格式已清除，西文统一回落为 Normal 样式的 Times New Roman。")
    s203 = (cr203 == "a" and
            items.get("CR-203", {}).get("decision") in ("approved", "modified"))
    if s203:
        notes["s203"] = ctx.ncid(
            "【CR-203】二级标题中文字体统一（样式级）：heading 2 定义原带 "
            "w:eastAsiaTheme=\"minorEastAsia\"，全文 29 段二级标题的中文因此渲染为「等线」，"
            "而标题 1/3/4 均为宋体。已删除该主题字体属性，改为继承宋体，与其余标题一致。"
            "（样式表改动无法用修订标记呈现，故以本批注说明。）")

    # ---- 逐段改写 ----
    h2_idx = None
    for i, (s, e) in enumerate(spans):
        if '<w:pStyle w:val="3"/>' in doc[s:e]:
            h2_idx = i
            break
    plan = []           # (idx, kind, detail)
    out_parts = []
    prev = 0
    for i, (s, e) in enumerate(spans):
        out_parts.append(doc[prev:s])
        seg = doc[s:e]
        new = seg
        nn = 0

        if i in eq101 or i in eq102:
            new, k = apply_eq_align(new, ctx, notes.get("eq") if i == min(eq101 + eq102) else None)
            if k:
                nn += k
                ctx.bump("eq_align")
                plan.append((i, "eq_align", "制表位/jc/Tab x%d" % k))
        if i in h201:
            cid = notes.get("h201") if i == min(h201) else None
            new, k = apply_substyle(new, ctx, cid, "heading")
            if k:
                nn += k
                ctx.bump("cr201")
                plan.append((i, "cr201", "%d 处直接格式" % k))
        if i in s202:
            cid = notes.get("s202") if i == min(s202) else None
            new, k = apply_style_swap(new, ctx, cid, "4", True, "heading")
            if k:
                nn += k
                ctx.bump("cr202")
                plan.append((i, "cr202", "Normal(Web)->heading 3"))
        if i in c204:
            cid = notes.get("c204") if i == min(c204) else None
            new, k = apply_style_swap(new, ctx, cid, "12", False, "caption")
            if k:
                nn += k
                ctx.bump("cr204")
                plan.append((i, "cr204", "Normal(Web)->caption"))
        if i in f205:
            cid = notes.get("f205") if i == min(f205) else None
            new, k = apply_substyle(new, ctx, cid, "font")
            if k:
                nn += k
                ctx.bump("cr205", k)
                plan.append((i, "cr205", "%d 处非 TNR" % k))

        if s203 and i == h2_idx:
            new = attach_note(new, notes["s203"])
            ctx.bump("s203_note")

        out_parts.append(new)
        prev = e
    out_parts.append(doc[prev:])
    new_doc = "".join(out_parts)

    new_styles = styles
    if cr203 in ("a", "b"):
        if cr203 == "a":
            new_styles, k = patch_heading2(styles, True)
            if k:
                ctx.bump("cr203")
                plan.append((-1, "cr203", "heading 2 去 eastAsiaTheme"))
        else:
            plan.append((-1, "cr203", "保持现状（选 b）"))

    # ---- 批注写入 ----
    new_comments = patch_comments(comments, ctx)
    new_cx = patch_cx(cx, ctx)

    # ---- 报告 ----
    print("=== 格式与编号对齐 · %s ===" % ("DRY-RUN" if dry else "落盘"))
    print("输入 : %s" % inp)
    print("作者 : %s   时间: %s" % (author, ctx.date))
    print("轨迹 : %s" % ("保留修订痕迹 (pPrChange/rPrChange/ins/del)" if trace else "不留痕（直接改）"))
    print("批准 : %s" % ", ".join("%s=%s" % (k, items[k].get("decision"))
                                 for k in sorted(todo)))
    print("\n--- 改动明细 ---")
    for (i, kind, det) in plan:
        print("  段%-6s %-10s %s" % (i if i >= 0 else "(样式)", kind, det))
    print("\n--- 统计 ---")
    print("  CR-101/102 编号段对齐 : %d 段" % ctx.hits.get("eq_align", 0))
    print("  CR-201 标题清直接格式 : %d 段" % ctx.hits.get("cr201", 0))
    print("  CR-202 标题样式归位   : %d 段" % ctx.hits.get("cr202", 0))
    print("  CR-204 题注样式归位   : %d 段" % ctx.hits.get("cr204", 0))
    print("  CR-205 非 TNR 回落    : 段落标记 %d 段 + run 级 %d 处（共 %d 段受影响）"
          % (ctx.hits.get("cr205_ppr", 0), ctx.hits.get("cr205_run", 0),
             sum(1 for p in plan if p[1] == "cr205")))
    print("  CR-203 样式定义       : %d 处" % ctx.hits.get("cr203", 0))
    print("  新增批注              : %d" % len(ctx.comments))
    print("  修订 id 区间          : %d..%d" % (rev0, ctx.rev - 1))

    # ---- 结构校验 ----
    ndoc = etree.fromstring(new_doc.encode("utf-8"))
    def cnt(r, tag):
        return sum(1 for _ in r.iter(q(tag)))
    print("\n--- 结构校验（相对原文档，应为 0）---")
    bad = False
    for tag in ("p", "tbl", "drawing", "bookmarkStart", "bookmarkEnd",
                "fldChar", "commentRangeStart", "commentRangeEnd",
                "commentReference"):
        d = cnt(ndoc, tag) - cnt(root, tag)
        if tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            want = len(ctx.comments)
            ok = (d == want)
        else:
            ok = (d == 0)
        if not ok:
            bad = True
        print("  %-20s %+5d  %s" % ("w:" + tag, d, "OK" if ok else "FAIL"))

    if dry:
        print("\n[DRY] 未写盘。")
        return 2 if bad else 0
    if bad:
        sys.stderr.write("结构校验失败，拒绝写盘。\n")
        return 2

    # ---- 写盘 ----
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            data = z.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc.encode("utf-8")
            elif item.filename == "word/styles.xml":
                data = new_styles.encode("utf-8")
            elif item.filename == "word/comments.xml":
                data = new_comments.encode("utf-8")
            elif item.filename == "word/commentsExtended.xml":
                data = new_cx.encode("utf-8")
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.create_system = item.create_system
            zo.writestr(zi, data)
    z.close()
    print("\n写出 : %s (%d B)" % (out, os.path.getsize(out)))
    print("md5  : %s" % md5_file(out))

    receipt = {
        "generated_at": ctx.date,
        "author": author,
        "trace": trace,
        "target_md5": ap.get("target_md5"),
        "output_md5": md5_file(out),
        "supported": list(SUPPORTED),
        "items": {k: items[k].get("decision") for k in sorted(items)},
        "hits": ctx.hits,
        "plan": [{"para": i, "cr": k, "detail": d} for i, k, d in plan],
        "comments": [{"id": c, "text": t} for c, t in ctx.comments],
        "untouched": ["正文文字/编号/书签/域/批注原文/公式本体/图片（均不在本轮范围）"],
    }
    rp = os.path.splitext(out)[0] + "_回执.json"
    io.open(rp, "w", encoding="utf-8").write(json.dumps(receipt, ensure_ascii=False, indent=1))
    print("回执 : %s" % rp)
    return 0


def patch_comments(comments_xml, ctx):
    frag = []
    for cid, text in ctx.comments:
        pid = paraid(cid, text)
        frag.append(
            '<w:comment w:id="%d" w:author="%s" w:date="%s" w:initials="">'
            '<w:p w14:paraId="%s"><w:pPr><w:pStyle w:val="13"/>'
            '<w:spacing w:before="194" w:after="194"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
            '</w:comment>' % (cid, ctx.author, ctx.date, pid, text))
    j = comments_xml.rfind("</w:comments>")
    return comments_xml[:j] + "".join(frag) + comments_xml[j:]


def patch_cx(cx_xml, ctx):
    frag = []
    for cid, text in ctx.comments:
        frag.append('<w15:commentEx w15:paraId="%s" w15:done="0"/>' % paraid(cid, text))
    j = cx_xml.rfind("</w15:commentsEx>")
    return cx_xml[:j] + "".join(frag) + cx_xml[j:]


def paraid(cid, text):
    return hashlib.md5(("fmt%d-%s" % (cid, text[:10])).encode("utf-8")).hexdigest()[:8].upper()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
