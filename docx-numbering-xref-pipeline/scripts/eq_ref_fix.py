#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""公式引用治理专用执行器 —— docx-numbering-xref-pipeline

一次原子完成「公式引用」专项的 5 条 CRS 条目。它们落在同一批 run 上
（编号段末 run / 引用 run / 域结果 run），拆开分头执行必然互相破坏，
故必须同一个执行器一次做完：

  CR-001 (eq_bookmark)   90 个公式编号段补 `_KT<章>_E<序号>` 书签锚点
  CR-002 (eq_bookmark)   3 处失效 REF 域（「错误！未定义书签。」）恢复
  CR-003 (eq_ref_field)  13 处手打引用 → 真 REF 域
  CR-004 (text_replace)  编号括号统一全角（老板选「全角」）
  CR-005 (renumber)      第3章编号改章内（老板选「章内编号」）

实现口径（与项目硬约定一致）：
  * 只改 word/document.xml / comments.xml / commentsExtended.xml，其余部件
    逐 ZipInfo 原样复制。
  * 全部改动以修订标记（w:ins / w:del）呈现，并按 CR 分类挂批注说明原因。
  * 字符串级定位替换（不做 lxml 回写），保证非改动区逐字节不变。
  * 输入只读；输出写独立新文件。
  * 硬断言：编号段 / 引用 / 坏域 的数量与预期不符即拒绝执行（退码 4）。

用法:
  eq_ref_fix.py <in.docx> <out.docx> <approved.json> [--dry] [--author NAME]

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

SUPPORTED = ("eq_bookmark", "eq_ref_field", "text_replace", "renumber")
DEFAULT_AUTHOR = "docx-format 校对"

TAGP = re.compile(r'<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>')
RUNPAT = re.compile(r'<w:r(?:\s[^>]*)?>.*?</w:r>|<w:r(?:\s[^>]*)?/>', re.S)
RPRPAT = re.compile(r'<w:rPr>.*?</w:rPr>', re.S)
TPAT = re.compile(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', re.S)
FULLNUM = re.compile(r'^[（(]\s*(\d+(?:[-–—]\d+)?)\s*[)）]$')
ENDNUM = re.compile(r'^(\s*)[（(]\s*(\d+(?:[-–—]\d+)?)\s*[)）]\s*$')
REFPAT = re.compile(r'((?:公式|式)\s*)[（(]\s*(\d+(?:[-–—]\d+)?)\s*[)）]')
BADTEXT = "错误！未定义书签。"

EXPECT_NUM = 90
EXPECT_NUM_BY_CHAP = {2: 15, 3: 75}
EXPECT_REF = 13
EXPECT_BAD = 3


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


def runs_in(seg):
    return list(RUNPAT.finditer(seg))


def rpr_of(run):
    mo = RPRPAT.search(run)
    return mo.group(0) if mo else ""


def t_of(run):
    mo = TPAT.search(run)
    return mo.group(1) if mo else None


def ptext_of_seg(seg):
    return "".join(mo.group(1) for mo in TPAT.finditer(seg))


# ------------------------------------------------------------------ 片段
def w_del(text, rpr, rid, author, date):
    return ('<w:del w:id="%d" w:author="%s" w:date="%s">'
            '<w:r>%s<w:delText xml:space="preserve">%s</w:delText></w:r>'
            '</w:del>' % (rid, author, date, rpr, text))


def w_ins(text, rpr, rid, author, date):
    return ('<w:ins w:id="%d" w:author="%s" w:date="%s">'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
            '</w:ins>' % (rid, author, date, rpr, text))


def w_ins_inner(inner, rid, author, date):
    return ('<w:ins w:id="%d" w:author="%s" w:date="%s">%s</w:ins>'
            % (rid, author, date, inner))


def w_field(instr, cache, rpr):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r>%s<w:t>%s</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            % (instr, rpr, cache))


def plain_run(text, rpr):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, text)


def cref_run(cid):
    return ('<w:r><w:rPr><w:rStyle w:val="38"/><w:rFonts w:hint="eastAsia"/>'
            '<w:sz w:val="24"/><w:szCs w:val="22"/></w:rPr>'
            '<w:commentReference w:id="%d"/></w:r>' % cid)


# ------------------------------------------------------------------ 章节
def style_map(z):
    st = etree.fromstring(z.read("word/styles.xml"))
    d = {}
    for s in st.iter(q("style")):
        nm = s.find(q("name"))
        d[s.get(q("styleId"))] = nm.get(q("val")) if nm is not None else None
    return d


def heading_marks(doc, id2name):
    marks = []
    for i, p in enumerate(doc.iter(q("p"))):
        ppr = p.find(q("pPr"))
        e = ppr.find(q("pStyle")) if ppr is not None else None
        nm = id2name.get(e.get(q("val"))) if e is not None else None
        if (nm or "").lower().startswith("heading 1") and \
           "".join(t.text or "" for t in p.iter(q("t"))).strip():
            marks.append(i)
    return marks


def chap_of(marks, idx):
    c = 0
    for k, mk in enumerate(marks):
        if mk <= idx:
            c = k
        else:
            break
    return c


# ------------------------------------------------------------------ 主流程
class Ctx(object):
    def __init__(self, author, date):
        self.author = author
        self.date = date
        self.bk = 1320
        self.rev = 1
        self.cid = 4
        self.comments = []
        self.log = []

    def nbk(self):
        v = self.bk; self.bk += 1; return v

    def nrev(self):
        v = self.rev; self.rev += 1; return v

    def ncid(self, text):
        v = self.cid; self.cid += 1
        self.comments.append((v, text))
        return v


def main(argv):
    if len(argv) < 4:
        sys.stderr.write(__doc__)
        return 6
    inp, out, appr = argv[1], argv[2], argv[3]
    dry = "--dry" in argv
    author = argv[argv.index("--author") + 1] if "--author" in argv else DEFAULT_AUTHOR
    if not os.path.isfile(inp):
        sys.stderr.write("找不到输入: %s\n" % inp); return 6
    if not os.path.isfile(appr):
        sys.stderr.write("缺少批准文件（CRS 门禁）: %s\n" % appr); return 3

    ap = json.load(io.open(appr, encoding="utf-8"))
    items = {i["id"]: i for i in ap.get("items", [])}
    todo = [k for k, v in items.items() if v.get("decision") in ("approved", "modified")]
    if not todo:
        sys.stderr.write("批准文件无 approved/modified 项\n"); return 3

    cur = md5_file(inp)
    if ap.get("target_md5") and cur != ap["target_md5"]:
        sys.stderr.write("指纹不符: %s != %s\n" % (cur, ap["target_md5"])); return 5

    z = zipfile.ZipFile(inp)
    xml = z.read("word/document.xml").decode("utf-8")
    comments = z.read("word/comments.xml").decode("utf-8")
    cx = z.read("word/commentsExtended.xml").decode("utf-8")
    doc = etree.fromstring(z.read("word/document.xml"))
    id2name = style_map(z)
    marks = heading_marks(doc, id2name)

    spans = split_paras(xml)
    n_lxml = sum(1 for _ in doc.iter(q("p")))
    assert len(spans) == n_lxml, "段落口径不一致 %d vs %d" % (len(spans), n_lxml)

    date = now_iso()
    ctx = Ctx(author, date)
    plist = list(doc.iter(q("p")))

    # ---------- 第一遍：分类所有段落 ----------
    num_paras = {}       # idx -> (value, chapter)
    ref_paras = {}       # idx -> (matches, chapter)
    bad_paras = {}       # idx -> instr
    for i, (s, e) in enumerate(spans):
        seg = xml[s:e]
        p = plist[i]
        has_omath = p.find(".//" + m("oMath")) is not None
        # 编号段：含公式 + 末个非空 run 严格是「(n)」/「(章-n)」
        # （整段文本可能夹属文，故不能要求整段就是编号——段1906 即如此）
        last_t = last_text_run(seg)
        if has_omath and last_t is not None:
            lm = ENDNUM.match(t_of(last_t.group(0)))
            if lm and int(lm.group(2).split("-")[-1]) <= 150:
                num_paras[i] = (lm.group(2), chap_of(marks, i))
                continue
        # 引用段
        full = "".join(mo.group(1) for mo in TPAT.finditer(seg))
        ms = list(REFPAT.finditer(full))
        if ms:
            ref_paras[i] = (ms, chap_of(marks, i))
        # 坏域段
        bk = find_bad_run(seg)
        if bk is not None:
            bad_paras[i] = instr_before(seg, bk.start())

    # ---------- 硬断言 ----------
    errs = []
    if len(num_paras) != EXPECT_NUM:
        errs.append("编号段 %d != %d" % (len(num_paras), EXPECT_NUM))
    per_chap = {}
    for _i, (_v, _c) in num_paras.items():
        per_chap[_c] = per_chap.get(_c, 0) + 1
    for ch, want in EXPECT_NUM_BY_CHAP.items():
        got = per_chap.get(ch, 0)
        if got != want:
            errs.append("第%d章编号段 %d != %d" % (ch, got, want))
    n_ref = sum(len(v[0]) for v in ref_paras.values())
    if n_ref != EXPECT_REF:
        errs.append("引用 %d != %d" % (n_ref, EXPECT_REF))
    if len(bad_paras) != EXPECT_BAD:
        errs.append("坏域 %d != %d" % (len(bad_paras), EXPECT_BAD))
    if errs:
        sys.stderr.write("数量断言失败（文档已变，重出单）：\n  " + "\n  ".join(errs) + "\n")
        return 4

    # ---------- 分配批注 ----------
    by_chap = {}
    for i, (v, ch) in num_paras.items():
        by_chap.setdefault(ch, []).append(i)
    num_note = {}
    for ch in sorted(by_chap):
        lst = sorted(by_chap[ch])
        n = len(lst)
        if ch == 3:
            txt = ("【CR-001/004/005】公式编号规范化（第%d章，共%d处）：① 补书签锚点 "
                   "`_KT%d_E01`~`_KT%d_E%02d`（原文档书签已丢失，仅存 3 处坏域残骸 "
                   "`_KT2_E01/E04/E11` 可反推命名体例）；② 编号括号统一为全角；"
                   "③ 编号由全局 `1`~`%d` 改为章内 `%d-1`~`%d-%d`。"
                   "全部改动以修订标记呈现（删旧值 + 插新值）。"
                   % (ch, n, ch, ch, n, n, ch, ch, n))
        else:
            txt = ("【CR-001/004/005】公式编号规范化（第%d章，共%d处）：① 补书签锚点 "
                   "`_KT%d_E01`~`_KT%d_E%02d`（原文档书签已丢失，仅存 3 处坏域残骸可反推命名体例）；"
                   "② 编号括号统一为全角；③ 该章编号本为章内式 `%d-n`，保持不变。"
                   "全部改动以修订标记呈现。" % (ch, n, ch, ch, n, ch))
        num_note[ch] = ctx.ncid(txt)
    bad_note = ctx.ncid(
        "【CR-002】修复失效交叉引用（共 3 处）：原 REF 域目标书签 `_KT2_E01`/`_KT2_E04`/"
        "`_KT2_E11` 已随书签丢失，Word 显示「错误！未定义书签。」。补书签后域缓存同步更新为"
        "正确编号 `（2-1）`/`（2-4）`/`（2-11）`。")
    ref_note = ctx.ncid(
        "【CR-003】公式引用转为可刷新交叉引用（共 13 处）：正文中手打的「式(2-2)」等 13 处"
        "引用改为 REF 域（「式」前缀保留为普通文本）。效果：与图/表引用同体例，"
        "支持 F9 刷新与 Ctrl+点击跳转。")

    # 批注挂首个命中处
    first_of = {}
    for ch in sorted(by_chap):
        first_of[ch] = min(by_chap[ch])
    first_bad = min(bad_paras)
    first_ref = min(ref_paras)

    # ---------- 第二遍：逐段改写 ----------
    out_parts = []
    prev = 0
    stats = dict(num=0, ref=0, bad=0)
    plan_nums, plan_refs, plan_bads = [], [], []

    for i, (s, e) in enumerate(spans):
        out_parts.append(xml[prev:s])
        seg = xml[s:e]
        new_seg = seg

        # ---- 编号段 ----
        if i in num_paras:
            val, ch = num_paras[i]
            seq = int(val.split("-")[-1])
            newnum = "（%d-%d）" % (ch, seq)
            bkname = "_KT%d_E%02d" % (ch, seq)
            runmo = last_text_run(new_seg)
            old_run = runmo.group(0)
            rpr = rpr_of(old_run)
            lead = ENDNUM.match(t_of(old_run)).group(1)
            old_num_txt = t_of(old_run).strip()
            cid = num_note[ch] if i == first_of[ch] else None
            bkid, did, iid = ctx.nbk(), ctx.nrev(), ctx.nrev()
            parts = []
            if lead:
                parts.append(plain_run(lead, rpr))
            if cid is not None:
                parts.append('<w:commentRangeStart w:id="%d"/>' % cid)
            parts.append(w_del(old_num_txt, rpr, did, ctx.author, ctx.date))
            # 书签只包新的编号（ins），不包被删的旧值——否则显示修订标记时
            # REF 会把两段文本都算进去（显示成 "(2-1)（2-1）"）。
            parts.append('<w:bookmarkStart w:id="%d" w:name="%s"/>' % (bkid, bkname))
            parts.append(w_ins(newnum, rpr, iid, ctx.author, ctx.date))
            parts.append('<w:bookmarkEnd w:id="%d"/>' % bkid)
            if cid is not None:
                parts.append('<w:commentRangeEnd w:id="%d"/>' % cid)
                parts.append(cref_run(cid))
            new_seg = replace_at(new_seg, runmo.start(), runmo.end(), "".join(parts))
            stats["num"] += 1
            plan_nums.append((i, val, newnum, bkname))
            ctx.log.append("  [编号] 段%-6d %-8s -> %-8s  书签 %-12s %s"
                           % (i, "(%s)" % val, newnum, bkname,
                              ("批注#%d" % cid) if cid else ""))

        # ---- 引用段 ----
        if i in ref_paras:
            new_seg, k = rewrite_refs(new_seg, ref_paras[i][0], ref_paras[i][1],
                                     ctx, plan_refs, i,
                                     ref_note if i == first_ref else None)
            stats["ref"] += k

        # ---- 坏域段 ----
        if i in bad_paras:
            bk = find_bad_run(new_seg)
            if bk is not None:
                run = bk.group(0)
                rpr = rpr_of(run)
                instr = bad_paras[i] or ""
                mm = re.search(r'REF\s+_KT(\d+)_E(\d+)', instr)
                if mm:
                    ch, seq = int(mm.group(1)), int(mm.group(2))
                    newnum = "（%d-%d）" % (ch, seq)
                    did, iid = ctx.nrev(), ctx.nrev()
                    cid = bad_note if i == first_bad else None
                    mid = (w_del(BADTEXT, rpr, did, ctx.author, ctx.date)
                           + w_ins(newnum, rpr, iid, ctx.author, ctx.date))
                    ops = [(bk.start(), bk.end(), mid)]
                    if cid is not None:
                        # 批注区间必须包住整个域、且 reference 落在域**外**
                        # （放进 separate~end 结果区会被 F9 重算抹掉 → 批注成孤儿）
                        b0 = field_edge(new_seg, bk.start(), "begin", True)
                        e1 = field_edge(new_seg, bk.end(), "end", False)
                        if b0 is not None and e1 is not None and b0 <= bk.start() and e1 >= bk.end():
                            ops.append((e1, e1, '<w:commentRangeEnd w:id="%d"/>' % cid + cref_run(cid)))
                            ops.append((b0, b0, '<w:commentRangeStart w:id="%d"/>' % cid))
                        else:
                            cid = None
                    for (a, b, s) in sorted(ops, key=lambda x: -x[0]):
                        new_seg = new_seg[:a] + s + new_seg[b:]
                    stats["bad"] += 1
                    plan_bads.append((i, instr.strip(), newnum))
                    ctx.log.append("  [坏域] 段%-6d %-22s -> %s%s"
                                   % (i, instr.strip(), newnum,
                                      ("  批注#%d" % cid) if cid else ""))

        out_parts.append(new_seg)
        prev = e
    out_parts.append(xml[prev:])
    new_xml = "".join(out_parts)

    new_comments = patch_comments(comments, ctx)
    new_cx = patch_cx(cx, ctx)

    # ---------- 报告 ----------
    print("=== 公式引用治理 · %s ===" % ("DRY-RUN" if dry else "落盘"))
    print("输入 : %s" % inp)
    print("作者 : %s   时间: %s" % (author, date_of(ctx)))
    print("批准 : %s" % ", ".join("%s=%s" % (k, items[k].get("decision")) for k in sorted(items)))
    print("\n--- 改动明细 ---")
    for line in ctx.log:
        print(line)
    print("\n--- 统计 ---")
    print("  编号段改写 + 书签 : %d" % stats["num"])
    print("  引用转 REF 域     : %d" % stats["ref"])
    print("  坏域修复          : %d" % stats["bad"])
    print("  新增批注          : %d" % len(ctx.comments))
    print("  书签 id 区间      : 1320..%d" % (ctx.bk - 1))
    print("  修订 id 区间      : 1..%d" % (ctx.rev - 1))

    # ---------- 结构校验 ----------
    ndoc = etree.fromstring(new_xml.encode("utf-8"))
    def cnt(root, tag):
        return sum(1 for _ in root.iter(q(tag.split(":")[-1])))
    checks = [("w:p", 0), ("w:tbl", 0), ("w:drawing", 0),
              ("w:bookmarkStart", stats["num"]), ("w:bookmarkEnd", stats["num"]),
              ("w:fldChar", stats["ref"] * 3),
              ("w:commentRangeStart", len(ctx.comments)),
              ("w:commentRangeEnd", len(ctx.comments)),
              ("w:commentReference", len(ctx.comments)),
              ("w:ins", stats["num"] + stats["ref"] + stats["bad"]),
              ("w:del", stats["num"] + stats["ref"] + stats["bad"])]
    print("\n--- 结构校验（相对原文档的增减）---")
    bad = False
    for tag, want in checks:
        d = cnt(ndoc, tag) - cnt(doc, tag)
        ok = (d == want)
        if not ok:
            bad = True
        print("  %-22s %+5d  %s" % (tag, d, "OK" if ok else "FAIL(期望 %+d)" % want))

    if dry:
        print("\n[DRY] 未写盘。")
        return 2 if bad else 0
    if bad:
        sys.stderr.write("结构校验失败，拒绝写盘。\n"); return 2

    # ---------- 写盘 ----------
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            data = z.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
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
        "generated_at": date_of(ctx),
        "author": author,
        "target_md5": ap.get("target_md5"),
        "output_md5": md5_file(out),
        "supported": list(SUPPORTED),
        "items": {k: items[k].get("decision") for k in sorted(items)},
        "stats": stats,
        "bookmarks": [{"para": a, "old": b, "new": c, "name": d} for a, b, c, d in plan_nums],
        "refs": [{"para": a, "old": b, "name": c, "new": d} for a, b, c, d in plan_refs],
        "bad_refs": [{"para": a, "instr": b, "new": c} for a, b, c in plan_bads],
        "comments": [{"id": c, "text": t} for c, t in ctx.comments],
        "untouched": ["正文文字/样式/排版/图表编号/公式本体（均不在本轮范围）"],
    }
    rp = os.path.splitext(out)[0] + "_回执.json"
    io.open(rp, "w", encoding="utf-8").write(json.dumps(receipt, ensure_ascii=False, indent=1))
    print("回执 : %s" % rp)
    return 0


def date_of(ctx):
    return ctx.date


def last_text_run(seg):
    hit = None
    for mo in runs_in(seg):
        t = t_of(mo.group(0))
        if t and t.strip():
            hit = mo
    return hit


def replace_at(seg, a, b, new):
    return seg[:a] + new + seg[b:]


def find_bad_run(seg):
    """定位「错误！未定义书签。」所在的 run（BADTEXT 全文唯一，直接按文本找，
    避免 find_first 抓到段内新插入域的 separate）。"""
    i = seg.find(BADTEXT)
    if i < 0:
        return None
    # 必须在域结果区里：前面有 separate、后面有 end
    if seg.rfind('w:fldCharType="separate"', 0, i) < 0:
        return None
    if seg.find('w:fldCharType="end"', i) < 0:
        return None
    for mo in RUNPAT.finditer(seg):
        if mo.start() <= i < mo.end():
            return mo
    return None


def instr_before(seg, pos):
    mo = None
    for mm in re.finditer(r'<w:instrText[^>]*>(.*?)</w:instrText>', seg[:pos], re.S):
        mo = mm
    return mo.group(1) if mo else None


def field_edge(seg, pos, kind, back):
    """找 pos 之前(back=True)/之后 最近的 fldCharType=kind，返回其所在 run 的
    起点(start) 或 终点(end)。用于把批注区间套在域外面。"""
    if back:
        k = seg.rfind('w:fldCharType="%s"' % kind, 0, pos)
    else:
        k = seg.find('w:fldCharType="%s"' % kind, pos)
    if k < 0:
        return None
    for mo in RUNPAT.finditer(seg):
        if mo.start() <= k < mo.end():
            return mo.start() if back else mo.end()
    return None


def rewrite_refs(seg, matches, chap, ctx, plan, pid, note_cid=None):
    """把段内 REFPAT 命中改为 REF 域。返回 (新段, 命中数)。
    note_cid 非空时，首个命中处挂批注区间。"""
    # 纯文本 run（不含域标记）
    pieces = []
    for mo in runs_in(seg):
        run = mo.group(0)
        if "<w:fldChar" in run or "<w:instrText" in run:
            continue
        t = t_of(run)
        if t is None:
            continue
        pieces.append([mo.start(), mo.end(), run, t])
    full = "".join(p[3] for p in pieces)
    ms = list(REFPAT.finditer(full))
    if not ms:
        return seg, 0
    tag = [-1] * len(full)
    for k, mm in enumerate(ms):
        for p in range(mm.start(), mm.end()):
            tag[p] = k
    repls = []
    emitted = {}
    noted = [False]
    pos = 0
    for (st, en, run, t) in pieces:
        parts = []
        cur = 0
        changed = False
        rpr = rpr_of(run)
        while cur < len(t):
            k = tag[pos + cur]
            nxt = cur + 1
            while nxt < len(t) and tag[pos + nxt] == k:
                nxt += 1
            txt = t[cur:nxt]
            if k == -1:
                parts.append(plain_run(txt, rpr))
            else:
                changed = True
                if k not in emitted:
                    mm = ms[k]
                    pre, val = mm.group(1), mm.group(2)
                    seq = int(val.split("-")[-1])
                    bkname = "_KT%d_E%02d" % (chap, seq)
                    newnum = "（%d-%d）" % (chap, seq)
                    did, iid = ctx.nrev(), ctx.nrev()
                    inner = plain_run(pre, rpr) + w_field(
                        " REF %s \\h " % bkname, newnum, rpr)
                    blk = [w_del(mm.group(0), rpr, did, ctx.author, ctx.date),
                           w_ins_inner(inner, iid, ctx.author, ctx.date)]
                    if note_cid is not None and not noted[0]:
                        noted[0] = True
                        blk = (['<w:commentRangeStart w:id="%d"/>' % note_cid]
                               + blk
                               + ['<w:commentRangeEnd w:id="%d"/>' % note_cid,
                                  cref_run(note_cid)])
                    parts.extend(blk)
                    emitted[k] = True
                    plan.append((pid, mm.group(0), bkname, newnum))
                    ctx.log.append("  [引用] 段%-6d %-12s -> REF %-12s" % (pid, mm.group(0), bkname))
            cur = nxt
        if changed:
            repls.append((st, en, "".join(parts)))
        pos += len(t)
    for (st, en, new) in sorted(repls, reverse=True):
        seg = replace_at(seg, st, en, new)
    return seg, len(ms)


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
    return hashlib.md5(("eqref-%d-%s" % (cid, text[:10])).encode("utf-8")).hexdigest()[:8].upper()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
