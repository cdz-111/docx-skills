#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""题注编号修正执行器 —— CR-301/302/303/304/305/306

一次原子完成「5 段借号题注改标准题注 + 5 处正文引用改指 + 题注样式归位
+ caption 误用清理 + 标题直接格式清理 + 题注硬字体清理」。

  CR-301 (cap_rebuild)  5 段 REF 域题注 -> 标准 SEQ 题注（+ _Ref 书签）
  CR-302 (xref_fix)     5 处正文 REF 域改指正确目标 + 域结果同步
  CR-303 (style_swap)   图3-32/3-33/4-16 补 caption 样式
  CR-304 (style_swap)   「（图像输入）」等去 caption 样式；1 空段删除
  CR-305 (format_scrub) 22 段标题清直接格式
  CR-306 (format_scrub) 段2330 题注正文 run 硬字体清理（含在 CR-301 内）

口径（老板 2026-10-10 20:34 确认）：
  * Q1 = a 段2328 引用保持不动
  * Q2 = b 直接改干净（不加 w:ins/w:del/pPrChange）+ 逐条批注
  * 只改 word/document.xml；新建 word/comments.xml + commentsExtended.xml，
    并在 [Content_Types].xml 与 word/_rels/document.xml.rels 注册；其余部件逐字节复制
  * 字符串级替换，非改动区逐字节不变
  * 硬断言：域块数 / 段号集合不符预期即拒绝执行（退码 4）

用法:
  capfix.py <in.docx> <out.docx> <approved.json> [--dry] [--author NAME]

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

DEFAULT_AUTHOR = "docx-format 校对"

TAGP = re.compile(r'<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>')
PPR_SCAN = re.compile(r'</?w:pPr>')
TPAT = re.compile(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', re.S)
RPRPAT = re.compile(r'<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>', re.S)

# REF 域 + 嵌套 STYLEREF + 硬编码编号 的整块（异常题注 / 正文引用共用）
REFBLK = re.compile(
    r'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    r'<w:r><w:instrText xml:space="preserve">'
    r'\s*REF\s+(?P<ref>\S+)\s+\\h\s*</w:instrText></w:r>'
    r'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    r'<w:r><w:t xml:space="preserve">(?P<label>[图表])\s*</w:t></w:r>'
    r'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    r'<w:r><w:instrText xml:space="preserve">'
    r'\s*STYLEREF\s+1\s+\\s\s*</w:instrText></w:r>'
    r'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    r'<w:r><w:t>(?P<ch>\d+)</w:t></w:r>'
    r'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    r'<w:r><w:t>-(?P<num>\d+)</w:t></w:r>'
    r'<w:r><w:fldChar w:fldCharType="end"/></w:r>')

# ---------------- 计划表 ----------------
# CR-301：段 -> (标签, 新编号, 新书签名)
PLAN301 = {
    2330: ('图', 2,  '_Ref900301'),
    4793: ('表', 26, '_Ref900302'),
    4894: ('表', 27, '_Ref900303'),
    4933: ('表', 28, '_Ref900304'),
    5051: ('表', 29, '_Ref900305'),
}
# CR-302：段 -> (旧目标, 新目标, 旧编号, 新编号)
PLAN302 = {
    4792: ('_Ref12658', '_Ref900302', 4, 26),
    4893: ('_Ref12913', '_Ref900303', 5, 27),
    4932: ('_Ref13442', '_Ref900304', 6, 28),
    5050: ('_Ref14278', '_Ref900305', 7, 29),
    5108: ('_Ref14278', '_Ref900305', 7, 29),
}
CR303 = [5169, 5176, 6650]          # 补 caption 样式
CR304_DROP = [5921, 5990]           # 去 caption 样式
CR304_DEL = [7765]                  # 空段删除
# CR-305：段 -> 清理动作集
P305 = {
    2323: ('keepNext', 'keepLines', 'jc'),
    453:  ('keepNext', 'keepLines', 'ind'),
    2246: ('keepNext', 'keepLines'),
    2254: ('keepNext', 'keepLines'),
    2262: ('keepNext', 'keepLines'),
    2289: ('keepNext', 'keepLines'),
    591:  ('ind',), 999: ('ind',), 1115: ('ind',), 7653: ('ind',),
    345:  ('ind',), 1668: ('ind',), 2132: ('ind',), 2317: ('ind',),
    1669: ('rpffont',), 1676: ('rpffont',), 1821: ('rpffont',),
    1976: ('rpffont',), 2126: ('rpffont',), 2141: ('rpffont',),
    2190: ('rpffont',), 2218: ('rpffont',),
}
# 批注挂载点（每 CR 一条，段号互不重叠）
NOTE_AT = {'CR-301': 4793, 'CR-302': 4792, 'CR-303': 5169,
           'CR-304': 5921, 'CR-305': 2323, 'CR-306': 2330}
NOTE_TEXT = {
    'CR-301': "【CR-301】原题注用「插入→交叉引用」做成的 REF 域，编号从别的题注借用——"
              "造成 图3-1 出现 2 次、表3-4~3-7 各出现 2 次。已改为标准题注："
              "独立 _Ref 书签 + STYLEREF 1 \\s 域 + '-' + SEQ 图/表 \\* ARABIC \\s 1 域，"
              "与全篇 300 个题注逐字同构；编号顺延为 图3-2 / 表3-26~3-29（老板选定「顺延」口径）。",
    'CR-302': "【CR-302】正文引用原指向「借用编号」的那张表（点进去跳到无关表格："
              "讲 VisDrone 却跳到「创造力数据集」）。已将 REF 目标改指 CR-301 新建的 _Ref 书签，"
              "域结果缓存同步为正确编号，F9 后一致。",
    'CR-303': "【CR-303】本段题注的段落样式丢失（无 pStyle，等同正文），字体/字号/居中与"
              "其余 300 个题注不一致。已补 caption 样式。",
    'CR-304': "【CR-304】本段（表内说明行）误套 caption 题注样式，会被图（表）目录收录成"
              "怪条目。已改回正文样式（跟随 Normal）。",
    'CR-305': "【CR-305】标题段遗留的粘贴直接格式（keepNext=0 / keepLines=0 / jc / ind / "
              "color 等）覆盖了样式定义——其中段 1821/1976/2126 正是「从 2.3.2 起标题异常」"
              "那片区域，上一轮只改了样式归属、未清直接格式，本轮一并清除，回归 heading 样式。",
    'CR-306': "【CR-306】题注正文 run 硬写 w:ascii=Arial/宋体 + 加粗 + 字距，与其余题注"
              "（继承 Times New Roman + 宋体）不一致。已删除，随 caption 样式继承。",
}


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


def ppr_span(seg):
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


def paraid(cid, text):
    return hashlib.md5(("cap%d-%s" % (cid, text[:10])).encode("utf-8")).hexdigest()[:8].upper()


def cref_run(cid):
    return ('<w:r><w:rPr><w:rStyle w:val="38"/><w:rFonts w:hint="eastAsia"/>'
            '<w:sz w:val="24"/><w:szCs w:val="22"/></w:rPr>'
            '<w:commentReference w:id="%d"/></w:r>' % cid)


def attach_note(seg, cid):
    """批注区间：start 在段首，end+reference 在段尾（均落在任何域区间之外）。"""
    if not seg.endswith("</w:p>"):
        return seg
    body = seg[:-len("</w:p>")]
    i = body.find(">") + 1
    return (body[:i] + '<w:commentRangeStart w:id="%d"/>' % cid + body[i:]
            + '<w:commentRangeEnd w:id="%d"/>' % cid + cref_run(cid) + "</w:p>")


def new_cap_block(bkid, bkname, label, ch, num):
    return ('<w:bookmarkStart w:id="%d" w:name="%s"/>'
            '<w:r><w:t xml:space="preserve">%s </w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> STYLEREF 1 \\s </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>%s</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            '<w:r><w:t>-</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> SEQ %s \\* ARABIC \\s 1 </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>%d</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            '<w:bookmarkEnd w:id="%d"/>') % (bkid, bkname, label, ch, label, num, bkid)


def renumber_ch3(doc):
    """把第3章所有图/表题注的域缓存编号刷成顺延后的正确值（打开即正确，无需 F9）。"""
    sp = split_paras(doc)
    h1 = []
    for i, (a, b) in enumerate(sp):
        seg = doc[a:b]
        if '<w:pStyle w:val="2"/>' in seg:
            h1.append((i, ''.join(m.group(1) for m in TPAT.finditer(seg))))
    s = e = None
    for k, (i, t) in enumerate(h1):
        if '认知驱动场景化感知' in t:
            s = i
            e = h1[k + 1][0] if k + 1 < len(h1) else len(sp)
    if s is None:
        return doc, 0, []
    edits = []
    nseg = 0
    unmatched = []
    RPR_OPT = r'(?:<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>)?'
    for kind in ('图', '表'):
        n = 0
        pat = re.compile(
            r'(<w:instrText[^>]*> SEQ ' + kind + r' \\\* ARABIC \\s 1 </w:instrText></w:r>'
            r'<w:r>' + RPR_OPT + r'<w:fldChar w:fldCharType="separate"/></w:r>'
            r'<w:r>' + RPR_OPT + r'<w:t[^>]*>)(\d+)(</w:t>)')
        for i in range(s, e):
            a, b = sp[i]
            seg = doc[a:b]
            if (' SEQ %s \\* ARABIC \\s 1 ' % kind) not in seg:
                continue
            body = re.sub(r'<w:instrText[^>]*>.*?</w:instrText>', '', seg, flags=re.S)
            txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', body, re.S))
            if not txt.startswith(kind):
                continue
            n += 1
            hits = list(pat.finditer(seg))
            if len(hits) != 1:
                unmatched.append((i, kind, len(hits)))
                continue
            m = hits[0]
            if int(m.group(2)) != n:
                edits.append((a + m.start(2), a + m.end(2), str(n)))
                nseg += 1
    for st, en, v in sorted(edits, reverse=True):
        doc = doc[:st] + v + doc[en:]
    return doc, len(edits), unmatched


CANON_PPR = '<w:pPr><w:pStyle w:val="12"/><w:spacing w:before="194" w:after="194"/></w:pPr>'
CANON_RPR = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
LONGSPACING = ('<w:spacing w:before="194" w:beforeLines="50" w:beforeAutospacing="0" '
               'w:after="194" w:afterLines="50" w:afterAutospacing="0"/>')


# ------------------------------------------------------------------ 各 CR
def do_cr301(seg, ctx, idx):
    label, num, bkname = PLAN301[idx]
    bkid = ctx.nbk()
    m = REFBLK.search(seg)
    if not m:
        ctx.fail.append("CR-301 段%d 未找到 REF 域块" % idx)
        return seg
    if m.group('label') != label:
        ctx.fail.append("CR-301 段%d 标签不符 %s!=%s" % (idx, m.group('label'), label))
        return seg
    ch = m.group('ch')
    newblk = new_cap_block(bkid, bkname, label, ch, num)
    out = seg[:m.start()] + newblk + seg[m.end():]

    # pPr 规范化为与标准题注一致
    sp = ppr_span(out)
    if sp:
        ppr = out[sp[0]:sp[1]]
        out = out[:sp[0]] + CANON_PPR + out[sp[1]:]
        if ppr != CANON_PPR:
            ctx.bump("cr301_ppr")
    # 段2330 另有 run 级硬字体（CR-306）
    if idx == 2330:
        n_rpr = len(RPRPAT.findall(out))
        out = RPRPAT.sub(CANON_RPR, out)
        ctx.bump("cr306_rpr", n_rpr)
    ctx.bump("cr301")
    ctx.log.append((idx, "CR-301", "%s -> %s %s-%d" % (m.group('label'), label, ch, num)))
    return out


def do_cr302(seg, ctx, idx):
    oldref, newref, oldnum, newnum = PLAN302[idx]
    n = 0
    def rep(m):
        nonlocal n
        if m.group('ref') != oldref:
            ctx.fail.append("CR-302 段%d 旧目标不符 %s!=%s" % (idx, m.group('ref'), oldref))
            return m.group(0)
        if int(m.group('num')) != oldnum:
            ctx.fail.append("CR-302 段%d 旧编号不符 %s!=%d" % (idx, m.group('num'), oldnum))
            return m.group(0)
        blk = m.group(0)
        blk = blk.replace(' REF %s \\h ' % oldref, ' REF %s \\h ' % newref, 1)
        blk = blk.replace('<w:t>-%d</w:t>' % oldnum, '<w:t>-%d</w:t>' % newnum, 1)
        n += 1
        return blk
    out = REFBLK.sub(rep, seg)
    if n != 1:
        ctx.fail.append("CR-302 段%d 命中 %d 次（应 1）" % (idx, n))
    ctx.bump("cr302", n)
    ctx.log.append((idx, "CR-302", "%s -> %s (3-%d -> 3-%d)" % (oldref, newref, oldnum, newnum)))
    return out


def do_cr303(seg, ctx, idx):
    sp = ppr_span(seg)
    if not sp:
        ctx.fail.append("CR-303 段%d 无 pPr" % idx)
        return seg
    ppr = seg[sp[0]:sp[1]]
    if '<w:pStyle' in ppr:
        ctx.fail.append("CR-303 段%d 已有 pStyle" % idx)
        return seg
    out = seg[:sp[0]] + ppr.replace('<w:pPr>', '<w:pPr><w:pStyle w:val="12"/>', 1) + seg[sp[1]:]
    ctx.bump("cr303")
    ctx.log.append((idx, "CR-303", "补 pStyle=caption"))
    return out


def do_cr304_drop(seg, ctx, idx):
    sp = ppr_span(seg)
    if not sp or '<w:pStyle w:val="12"/>' not in seg[sp[0]:sp[1]]:
        ctx.fail.append("CR-304 段%d 无 caption 样式" % idx)
        return seg
    out = seg.replace('<w:pStyle w:val="12"/>', '', 1)
    ctx.bump("cr304_drop")
    ctx.log.append((idx, "CR-304", "去 caption 样式"))
    return out


def do_cr305(seg, ctx, idx):
    ops = P305[idx]
    sp = ppr_span(seg)
    if not sp:
        return seg
    ppr = seg[sp[0]:sp[1]]
    old = ppr
    if 'keepNext' in ops:
        ppr = ppr.replace('<w:keepNext w:val="0"/>', '', 1)
    if 'keepLines' in ops:
        ppr = ppr.replace('<w:keepLines w:val="0"/>', '', 1)
    if 'jc' in ops:
        ppr = ppr.replace('<w:jc w:val="both"/>', '', 1)
        ppr = ppr.replace('<w:jc w:val="both" />', '', 1)
    if 'ind' in ops:
        ppr = re.sub(r'<w:ind\b[^>]*/>', '', ppr, count=1)
    if 'rpffont' in ops:
        def scrub(m):
            r = m.group(0)
            r = re.sub(r'<w:color\b[^>]*/>', '', r)
            r = re.sub(r'<w:kern\b[^>]*/>', '', r)
            r = re.sub(r'<w:szCs\b[^>]*/>', '', r)
            r = re.sub(r'<w:sz\b[^>]*/>', '', r)
            return r
        ppr = RPRPAT.sub(scrub, ppr, count=1)
        # rPr 清空则删掉空壳
        ppr = re.sub(r'<w:rPr>\s*</w:rPr>', '', ppr)
    if ppr != old:
        seg = seg[:sp[0]] + ppr + seg[sp[1]:]
        ctx.bump("cr305")
        ctx.log.append((idx, "CR-305", "清直接格式"))
    return seg


# ------------------------------------------------------------------ 主流程
def main(argv):
    if len(argv) < 4:
        sys.stderr.write(__doc__)
        return 6
    inp, out, appr = argv[1], argv[2], argv[3]
    dry = "--dry" in argv
    author = argv[argv.index("--author") + 1] if "--author" in argv else DEFAULT_AUTHOR
    if not os.path.isfile(inp):
        sys.stderr.write("找不到输入: %s\n" % inp)
        return 6
    if not os.path.isfile(appr):
        sys.stderr.write("缺少批准文件（CRS 门禁）: %s\n" % appr)
        return 3

    ap = json.load(io.open(appr, encoding="utf-8"))
    items = {i["id"]: i for i in ap.get("items", [])}
    todo = {k for k, v in items.items() if v.get("decision") in ("approved", "modified")}
    if not todo:
        sys.stderr.write("批准文件无 approved/modified 项\n")
        return 3

    cur = md5_file(inp)
    if ap.get("target_md5") and cur != ap["target_md5"]:
        sys.stderr.write("指纹不符: %s != %s\n" % (cur, ap["target_md5"]))
        return 5

    z = zipfile.ZipFile(inp)
    doc = z.read("word/document.xml").decode("utf-8")
    ct = z.read("[Content_Types].xml").decode("utf-8")
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    spans = split_paras(doc)

    class Ctx(object):
        def __init__(self):
            self.bk = max([int(x) for x in re.findall(r'<w:bookmarkStart w:id="(\d+)"', doc)] or [0]) + 1
            self.cid = 0
            self.hits = {}
            self.log = []
            self.fail = []
            self.comments = []

        def nbk(self):
            v = self.bk
            self.bk += 1
            return v

        def ncid(self, text):
            v = self.cid
            self.cid += 1
            self.comments.append((v, text))
            return v

        def bump(self, k, n=1):
            self.hits[k] = self.hits.get(k, 0) + n

    ctx = Ctx()
    date = now_iso()

    # ---- 硬断言：目标段号必须存在且是预期形态 ----
    expect = set(PLAN301) | set(PLAN302) | set(CR303) | set(CR304_DROP) | set(CR304_DEL) | set(P305)
    if max(expect) >= len(spans):
        sys.stderr.write("段号越界\n")
        return 4

    # ---- 批注 ----
    notes = {}
    for cr in ("CR-301", "CR-302", "CR-303", "CR-304", "CR-305", "CR-306"):
        if cr in todo:
            notes[cr] = ctx.ncid(NOTE_TEXT[cr])

    # ---- 逐段改写 ----
    out_parts = []
    prev = 0
    for i, (s, e) in enumerate(spans):
        out_parts.append(doc[prev:s])
        seg = doc[s:e]
        new = seg
        if i in PLAN301:
            new = do_cr301(new, ctx, i)
        if i in PLAN302:
            new = do_cr302(new, ctx, i)
        if i in CR303:
            new = do_cr303(new, ctx, i)
        if i in CR304_DROP:
            new = do_cr304_drop(new, ctx, i)
        if i in CR304_DEL:
            new = ''
            ctx.bump("cr304_del")
            ctx.log.append((i, "CR-304", "删除空段"))
        if i in P305:
            new = do_cr305(new, ctx, i)
        # 批注
        for cr, at in NOTE_AT.items():
            if i == at and cr in notes:
                new = attach_note(new, notes[cr])
        out_parts.append(new)
        prev = e
    out_parts.append(doc[prev:])
    new_doc = "".join(out_parts)

    # ---- 顺延同步：刷第3章题注缓存编号 ----
    new_doc, n_ren, unmatched = renumber_ch3(new_doc)
    ctx.hits["renumber"] = n_ren
    if unmatched:
        ctx.fail.append("顺延同步未匹配题注(段,类型,命中数): %s" % str(unmatched[:6]))

    # ---- 结构校验 ----
    def cnt(x, tag):
        return len(re.findall(r'<%s(?:\s[^>]*)?/?>' % tag, x))

    checks = []
    for tag, want in (("w:p", -1), ("w:tbl", 0), ("w:drawing", 0),
                      ("w:bookmarkStart", 5), ("w:bookmarkEnd", 5),
                      ("w:fldChar", 0)):
        d = cnt(new_doc, tag) - cnt(doc, tag)
        checks.append((tag, d, want, d == want))
    # 批注三件套应为批注数的 2 倍 / 1 倍
    for tag, mult in (("w:commentRangeStart", 1), ("w:commentRangeEnd", 1), ("w:commentReference", 1)):
        d = cnt(new_doc, tag) - cnt(doc, tag)
        checks.append((tag, d, len(notes), d == len(notes)))

    print("=== 命中 ===")
    print("  CR-301 题注重建     : %d 段" % ctx.hits.get("cr301", 0))
    print("  CR-302 引用改指     : %d 处" % ctx.hits.get("cr302", 0))
    print("  CR-303 补题注样式   : %d 段" % ctx.hits.get("cr303", 0))
    print("  CR-304 去题注样式   : %d 段 / 删段 %d" % (ctx.hits.get("cr304_drop", 0), ctx.hits.get("cr304_del", 0)))
    print("  CR-305 标题清格式   : %d 段（计划 %d）" % (ctx.hits.get("cr305", 0), len(P305)))
    print("  CR-306 run 硬字体   : %d 处" % ctx.hits.get("cr306_rpr", 0))
    print("  编号顺延同步        : %d 处域缓存（打开即正确，无需 F9）" % ctx.hits.get("renumber", 0))
    print("  新增批注            : %d" % len(notes))
    print("\n=== 结构校验 ===")
    bad = False
    for tag, d, want, ok in checks:
        if not ok:
            bad = True
        print("  %-22s %+4d (expect %+d)  %s" % (tag, d, want, "OK" if ok else "FAIL"))

    if ctx.fail:
        print("\n=== 断言失败 ===")
        for f in ctx.fail:
            print("  !", f)
        bad = True

    if dry:
        print("\n[DRY] 未写盘。")
        return 2 if bad else 0
    if bad:
        sys.stderr.write("结构校验/断言失败，拒绝写盘。\n")
        return 2

    # ---- 构造批注部件 ----
    W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    cxml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<w:comments %s %s %s>' % (W_NS, W14, W15)]
    ex = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<w15:commentsEx %s>' % W15]
    for cid, text in ctx.comments:
        pid = paraid(cid, text)
        cxml.append('<w:comment w:id="%d" w:author="%s" w:date="%s" w:initials="FMT">'
                    '<w:p w14:paraId="%s"><w:pPr><w:pStyle w:val="13"/></w:pPr>'
                    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                    '</w:p></w:comment>' % (cid, author, date, pid, text))
        ex.append('<w15:commentEx w15:paraId="%s" w15:done="0"/>' % pid)
    cxml.append('</w:comments>')
    ex.append('</w15:commentsEx>')
    comments_xml = "".join(cxml)
    cx_xml = "".join(ex)

    # 注册
    rid_c = "rId247"
    rid_e = "rId248"
    new_ct = ct.replace('</Types>',
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>'
        '</Types>')
    new_rels = rels.replace('</Relationships>',
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
        '<Relationship Id="%s" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>'
        '</Relationships>' % (rid_c, rid_e))

    # ---- 写盘 ----
    def zi_of(item):
        zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
        zi.compress_type = item.compress_type
        zi.external_attr = item.external_attr
        zi.internal_attr = item.internal_attr
        zi.create_system = item.create_system
        return zi

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            if item.filename == "word/document.xml":
                data = new_doc.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                data = new_ct.encode("utf-8")
            elif item.filename == "word/_rels/document.xml.rels":
                data = new_rels.encode("utf-8")
            else:
                data = z.read(item.filename)
            zo.writestr(zi_of(item), data)
        zo.writestr(zi_of(zipfile.ZipInfo("word/comments.xml")), comments_xml.encode("utf-8"))
        zo.writestr(zi_of(zipfile.ZipInfo("word/commentsExtended.xml")), cx_xml.encode("utf-8"))
    z.close()
    print("\n写出 : %s (%d B)" % (out, os.path.getsize(out)))
    print("md5  : %s" % md5_file(out))

    receipt = {
        "generated_at": date, "author": author,
        "target_md5": ap.get("target_md5"), "output_md5": md5_file(out),
        "items": {k: items[k].get("decision") for k in sorted(items)},
        "hits": ctx.hits, "plan": [{"para": i, "cr": c, "detail": d} for i, c, d in ctx.log],
        "comments": [{"id": c, "text": t} for c, t in ctx.comments],
        "untouched": ["正文文字（除引用编号）/其它书签/图片/公式本体/表格/页眉页脚"],
    }
    rp = os.path.splitext(out)[0] + "_回执.json"
    io.open(rp, "w", encoding="utf-8").write(json.dumps(receipt, ensure_ascii=False, indent=1))
    print("回执 : %s" % rp)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
