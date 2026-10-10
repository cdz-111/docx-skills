# -*- coding: utf-8 -*-
"""独立验收（不复用 refcache.py 任何函数）—— CRS-2026-10-10-REFCACHE
   用法: verify_round5.py <src.docx> <out.docx>
"""
import re, sys, zipfile, hashlib
from collections import Counter

SRC, OUT = sys.argv[1], sys.argv[2]
zs = zipfile.ZipFile(SRC)
zo = zipfile.ZipFile(OUT)
ns = zs.namelist()
no = zo.namelist()

PASS, FAIL = [], []


def ck(name, ok, extra=""):
    (PASS if ok else FAIL).append("%s %s" % (name, extra))


# ---------- 0. 部件表 & 逐字节一致 ----------
ck("部件数一致", len(ns) == len(no), "src=%d out=%d" % (len(ns), len(no)))
ck("部件名集合一致", set(ns) == set(no),
   "only_in_src=%s only_in_out=%s" % (sorted(set(ns) - set(no)), sorted(set(no) - set(ns))))
CHANGED = {"word/document.xml", "word/comments.xml", "word/commentsExtended.xml"}
diff = []
for n in ns:
    if n in CHANGED:
        continue
    if zs.read(n) != zo.read(n):
        diff.append(n)
ck("非改动部件逐字节一致", not diff, str(diff[:6]))

# ---------- 1. well-formed ----------
import xml.etree.ElementTree as ET
bad_xml = []
for n in no:
    if n.endswith(".xml") or n.endswith(".rels"):
        try:
            ET.fromstring(zo.read(n))
        except Exception as e:
            bad_xml.append((n, str(e)[:60]))
ck("XML 部件全 well-formed", not bad_xml, str(bad_xml[:4]))

d_src = zs.read("word/document.xml").decode("utf-8")
d = zo.read("word/document.xml").decode("utf-8")

# ---------- 2. 结构计数与基线一致 ----------
for tag in ("w:p", "w:tbl", "w:drawing", "w:bookmarkStart", "w:bookmarkEnd",
            "w:fldChar", "w:tabs", "m:oMath", "w:fldSimple"):
    a = len(re.findall(r'<%s(?:\s[^>]*)?/?>' % tag, d_src))
    b = len(re.findall(r'<%s(?:\s[^>]*)?/?>' % tag, d))
    ck("计数不变 %s" % tag, a == b, "%d -> %d" % (a, b))

# ---------- 3. 批注 ----------
cm = zo.read("word/comments.xml").decode("utf-8")
cme = zo.read("word/commentsExtended.xml").decode("utf-8")
n_cm = len(re.findall(r'<w:comment ', cm))
ck("批注数 = 12", n_cm == 12, str(n_cm))
ids = sorted(int(x) for x in re.findall(r'<w:comment w:id="(\d+)"', cm))
ck("批注 id 连续 0..11", ids == list(range(12)), str(ids))
ck("commentsEx 数 = 12", len(re.findall(r'<w15:commentEx ', cme)) == 12)
for t in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
    ck("正文 %s = 12" % t, len(re.findall(r'<%s[^>]*/>' % t, d)) == 12)
ck("无修订痕迹", d.count('<w:ins ') == 0 and d.count('<w:del ') == 0)

# ---------- 4. CR-401 REF 缓存一致性（独立实现） ----------
tags = re.compile(r'<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>')
sp = []
dep = 0; st = None
for mo in tags.finditer(d):
    s = mo.group(0)
    if s.endswith("/>"):
        if dep == 0:
            sp.append((mo.start(), mo.end()))
    elif s == "</w:p>":
        dep -= 1
        if dep <= 0:
            if st is not None:
                sp.append((st, mo.end()))
            st = None; dep = 0
    else:
        if dep == 0:
            st = mo.start()
        dep += 1

import bisect
starts = [a for a, b in sp]
para_of = lambda p: bisect.bisect_right(starts, p) - 1

# 题注：书签 -> 编号（SEQ 域结果 + 字面编号双取）
bk2num = {}
for a, b in sp:
    seg = d[a:b]
    m = re.search(r'SEQ\s+([图表])\s+\\\*\s+ARABIC\s+\\s\s+1\s*</w:instrText>'
                  r'(?:(?!</w:r>).)*?</w:r>'
                  r'(?:<w:r>(?:(?!</w:r>).)*?</w:r>)*?'
                  r'<w:r>(?:(?!</w:r>).)*?<w:t[^>]*>(\d+)</w:t>', seg, re.S)
    if not m:
        continue
    for bm in re.finditer(r'<w:bookmarkStart\b[^>]*w:name="(_Ref[^"]+)"', seg):
        bk2num.setdefault(bm.group(1), int(m.group(2)))

def get_ppr(seg):
    """深度计数取本段 pPr（不受 pPrChange 内层 </w:pPr> 干扰）。"""
    s = seg.find('<w:pPr>')
    if s < 0:
        return ''
    dep = 0
    for mo in re.finditer(r'</?w:pPr>', seg[s:]):
        if mo.group(0) == '<w:pPr>':
            dep += 1
        else:
            dep -= 1
            if dep == 0:
                return seg[s:s + mo.end()]
    return ''


TOK = re.compile(r'<w:fldChar(?P<a>[^>]*?)w:fldCharType="(?P<ft>begin|separate|end)"[^>]*/>'
                 r'|<w:instrText[^>]*>(?P<ins>[^<]*)</w:instrText>'
                 r'|<w:t(?:\s[^>]*)?>(?P<t>.*?)</w:t>', re.S)


def refs(seg):
    """REF 域扫描；嵌套域结果并入父域，得到完整显示文本。"""
    stack, out = [], []
    for mo in TOK.finditer(seg):
        if mo.group('ft'):
            ft = mo.group('ft')
            if ft == 'begin':
                stack.append({'i': None, 'r': False, 'txt': ''})
            elif ft == 'separate':
                if stack:
                    stack[-1]['r'] = True
            else:
                if stack:
                    fr = stack.pop()
                    if fr['i'] and fr['i'].startswith('REF '):
                        out.append(fr)
                    if stack and fr['r']:
                        stack[-1]['txt'] += fr['txt']
        elif mo.group('ins') is not None:
            if stack and stack[-1]['i'] is None:
                stack[-1]['i'] = mo.group('ins').strip()
        else:
            if stack and stack[-1]['r']:
                stack[-1]['txt'] += mo.group('t')
    return out


mism, dangling, tot, odd = [], [], 0, []
for a, b in sp:
    seg = d[a:b]
    for fr in refs(seg):
        m = re.match(r'REF\s+(_Ref\w+)', fr['i'])
        if not m:
            continue
        tot += 1
        tg = m.group(1)
        real = bk2num.get(tg)
        if real is None:
            dangling.append((para_of(a), tg))
            continue
        mm = re.search(r'([图表])\s*(\d+)\s*-\s*(\d+)', re.sub(r'[\s\u00a0]+', '', fr['txt']))
        if not mm:
            odd.append((para_of(a), fr['txt'][:20], real))
            continue
        shown = int(mm.group(3))
        if shown != real:
            mism.append((para_of(a), tg, shown, real))
ck("REF 目标无悬空", not dangling, str(dangling[:5]))
ck("REF 文本形态可解析", not odd, str(odd[:5]))
ck("REF 显示 = 题注编号（0 不符）", not mism, "不符 %d 处 %s" % (len(mism), mism[:6]))
print("   [info] 全文 REF 域总数 = %d" % tot)

# ---------- 5. Q1：0 处 asciiTheme=minorEastAsia 数字 run ----------
DIG = re.compile(r'[0-9]')
RUN = re.compile(r'<w:r(?:>|\s[^>]*>)((?:(?!</w:r>).)*?)</w:r>', re.S)
q1left = []
for i, (a, b) in enumerate(sp):
    seg = d[a:b]
    ppr = get_ppr(seg)
    ppr_theme = 'asciiTheme="minorEastAsia"' in ppr
    for mo in RUN.finditer(seg):
        inner = mo.group(1)
        txt = ''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', inner, re.S))
        if not DIG.search(txt):
            continue
        rp = re.search(r'<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>', inner, re.S)
        rf = re.search(r'<w:rFonts[^>]*/>', rp.group(0)) if rp else None
        rft = rf.group(0) if rf else ''
        if 'w:ascii="' in rft:
            continue
        if 'asciiTheme="minorEastAsia"' in rft or (ppr_theme and 'w:ascii="' not in rft):
            q1left.append((i, txt[:24]))
ck("Q1 残留 = 0", not q1left, str(q1left[:6]))

# ---------- 6. CR-402/403/404 ----------
def pseg(i):
    a, b = sp[i]
    return d[a:b]

for i in (4442, 4462, 4468):
    ck("CR-402 段%d 无 仿宋 西文" % i,
       not re.search(r'w:(?:ascii|hAnsi)="仿宋"', pseg(i)))
ck("CR-402 段4470 无 仿宋", 'w:ascii="仿宋"' not in pseg(4470))
for i in (5169, 5176):
    ck("CR-403 段%d 无 宋体 rPr" % i, 'w:ascii="宋体"' not in pseg(i))
ck("CR-404 段595 段落标记无 宋体西文",
   'w:ascii="宋体"' not in get_ppr(pseg(595)) and 'w:hAnsi="宋体"' not in get_ppr(pseg(595)))
for i in (6514, 6515, 6516):
    ck("CR-404 段%d 无 仿宋 西文" % i, not re.search(r'w:(?:ascii|hAnsi)="仿宋"', pseg(i)))
ck("CR-404 全文无无效字体名", 'Times New Roman Regular' not in d,
   str(len(re.findall(r'Times New Roman Regular', d))))

# ---------- 7. Q2：标题段 run 级冗余格式 = 0 ----------
Q2 = [591, 622, 904, 999, 1115, 1669, 1676, 1821, 1976, 2126,
      2141, 2190, 2218, 2254, 2262, 2289, 2323, 7575, 7653]
left = []
for i in Q2:
    seg = pseg(i)
    for mo in RUN.finditer(seg):
        inner = mo.group(1)
        if 'commentReference' in inner:
            continue
        rp = re.search(r'<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>', inner, re.S)
        if not rp:
            continue
        for pat in (r'<w:color\s+w:val="000000"\s*/>', r'<w:kern\s+w:val="2"\s*/>',
                    r'<w:sz\s+w:val="28"\s*/>', r'<w:szCs\s+w:val="28"\s*/>',
                    r'<w:bCs\s+w:val="0"\s*/>', r'<w:bCs\s*/>'):
            if re.search(pat, rp.group(0)):
                left.append((i, pat))
ck("Q2 残留 = 0", not left, str(left[:6]))

# ---------- 8. 全文字体最终态：含 ASCII 的 run 西文全 TNR ----------
S = {}
for m in re.finditer(r'<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>',
                     zo.read("word/styles.xml").decode("utf-8"), re.S):
    sid, body = m.group(1), m.group(2)
    bo = re.search(r'<w:basedOn\s+w:val="([^"]+)"', body)
    rf = re.search(r'<w:rFonts([^/>]*)/>', body)
    S[sid] = (bo.group(1) if bo else None, rf.group(1) if rf else None)


def style_font(sid):
    seen = set()
    while sid and sid in S and sid not in seen:
        seen.add(sid)
        bo, rf = S[sid]
        if rf:
            a = re.search(r'w:ascii="([^"]*)"', rf)
            if a and a.group(1):
                return a.group(1)
        sid = bo
    return None


LAT = re.compile(r'[A-Za-z0-9]')
non = []
for i, (a, b) in enumerate(sp):
    seg = d[a:b]
    if '<m:oMath' in seg:
        pass
    sm = re.search(r'<w:pStyle\s+w:val="([^"]+)"', seg)
    sid = sm.group(1) if sm else None
    pm = get_ppr(seg)
    pf = None
    if pm:
        prpr = re.search(r'<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>', pm, re.S)
        if prpr:
            rf = re.search(r'<w:rFonts[^>]*/>', prpr.group(0))
            if rf:
                aa = re.search(r'w:ascii="([^"]*)"', rf.group(0))
                pf = aa.group(1) if aa else None
    for mo in RUN.finditer(seg):
        inner = mo.group(1)
        if '<m:oMath' in inner or 'm:t>' in inner:
            continue
        txt = ''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', inner, re.S))
        if not LAT.search(txt):
            continue
        rp = re.search(r'<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>', inner, re.S)
        f = None
        if rp:
            rf = re.search(r'<w:rFonts[^>]*/>', rp.group(0))
            if rf:
                aa = re.search(r'w:ascii="([^"]*)"', rf.group(0))
                f = aa.group(1) if aa else None
        if f is None:
            f = pf
        if f is None:
            f = style_font(sid)
        if f is None:
            f = 'Times New Roman'
        if f != 'Times New Roman':
            non.append((i, f, txt[:26]))
ck("含 ASCII 的 run 西文全为 TNR", not non,
   "非 TNR %d 处 %s" % (len(non), non[:8]))

# ---------- 9. 体积变化归因 ----------
ss = sum(x.compress_size for x in zs.infolist())
os_ = sum(x.compress_size for x in zo.infolist())
print("\n   [info] 压缩后总字节 src=%d out=%d (Δ%+d)" % (ss, os_, os_ - ss))
print("   [info] 原始总字节 src=%d out=%d" %
      (sum(x.file_size for x in zs.infolist()), sum(x.file_size for x in zo.infolist())))
print("   [info] 解压后未变（除 3 个改动部件）：", not diff)

print("\n================ 验收结果 ================")
for x in PASS:
    print("  PASS  " + x)
for x in FAIL:
    print("  FAIL  " + x)
print("\nPASS %d / FAIL %d" % (len(PASS), len(FAIL)))
sys.exit(2 if FAIL else 0)
