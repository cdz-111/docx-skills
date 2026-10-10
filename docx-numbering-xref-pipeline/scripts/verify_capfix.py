# -*- coding: utf-8 -*-
"""独立验收：题注编号修正（CR-301~306）"""
import zipfile, re, sys, hashlib

A = sys.argv[1]   # 基线
B = sys.argv[2]   # 交付件
PASS = FAIL = 0

def chk(name, ok, detail=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        print("  [PASS] %s" % name)
    else:
        FAIL += 1
        print("  [FAIL] %s   %s" % (name, detail))

TAGP = re.compile(r'<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>')

def load(p):
    z = zipfile.ZipFile(p)
    d = {n: z.read(n) for n in z.namelist()}
    z.close()
    return d

za, zb = load(A), load(B)

print("=== 1. 部件清单 ===")
newp = set(zb) - set(za)
chk("新增 2 个批注部件", newp == {"word/comments.xml", "word/commentsExtended.xml"}, str(newp))
chk("部件总数 +2", len(zb) == len(za) + 2, "%d vs %d" % (len(za), len(zb)))
chk("无部件缺失", set(za) <= set(zb), str(set(za) - set(zb)))

print("=== 2. 非改动部件逐字节一致 ===")
CHANGED = {"word/document.xml", "[Content_Types].xml", "word/_rels/document.xml.rels"}
diff = [n for n in za if n not in CHANGED and za[n] != zb[n]]
chk("非改动部件逐字节一致", not diff, str(diff[:5]))

da = za["word/document.xml"].decode("utf-8")
db = zb["word/document.xml"].decode("utf-8")

def spa(xml):
    sp, dep, st = [], 0, None
    for mo in TAGP.finditer(xml):
        s = mo.group(0)
        if s.endswith("/>"):
            if dep == 0: sp.append((mo.start(), mo.end()))
        elif s == "</w:p>":
            dep -= 1
            if dep <= 0:
                if st is not None: sp.append((st, mo.end()))
                st, dep = None, 0
        else:
            if dep == 0: st = mo.start()
            dep += 1
    return sp

Sa, Sb = spa(da), spa(db)

print("=== 3. 结构计数 ===")
for tag, wantA, wantB in (("w:p", 8274, 8273), ("w:tbl", 84, 84), ("w:drawing", 227, 227)):
    n = len(re.findall(r'<%s(?:\s[^>]*)?/?>' % tag, db))
    chk("%s = %d" % (tag, wantB), n == wantB, "实际 %d" % n)

print("=== 4. 书签 ===")
bs = re.findall(r'<w:bookmarkStart w:id="(\d+)"', db)
be = re.findall(r'<w:bookmarkEnd w:id="(\d+)"', db)
chk("bookmarkStart 1415", len(bs) == 1415, str(len(bs)))
chk("bookmarkEnd 1415", len(be) == 1415, str(len(be)))
chk("id 唯一", len(set(bs)) == len(bs))
chk("start/end 配对", sorted(bs) == sorted(be))
names = re.findall(r'<w:bookmarkStart[^>]*w:name="([^"]*)"', db)
for nm in ("_Ref900301", "_Ref900302", "_Ref900303", "_Ref900304", "_Ref900305"):
    chk("书签 %s 存在" % nm, nm in names)
# 302~305 被正文引用（CR-302）；301 为新建可引用目标，当前无正文引用（Q1=a 保持段2328 不动）
for nm in ("_Ref900302", "_Ref900303", "_Ref900304", "_Ref900305"):
    chk("书签 %s 被正文引用" % nm, ("REF %s " % nm) in db)

print("=== 5. 第3章题注编号连续 ===")
def ch3_range(sp):
    h1 = []
    for i, (a, b) in enumerate(sp):
        seg = db[a:b]
        if '<w:pStyle w:val="2"/>' in seg:
            t = re.sub(r'\s+', ' ', ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', seg, re.S)))
            h1.append((i, t))
    s = e = None
    for k, (i, t) in enumerate(h1):
        if '认知驱动场景化感知' in t:
            s, e = i, h1[k+1][0]
    return s, e

s3, e3 = ch3_range(Sb)
def scan(kind):
    out = []
    for i in range(s3, e3):
        seg = db[Sb[i][0]:Sb[i][1]]
        if ('SEQ %s ' % kind) in seg and re.search(r'<w:t[^>]*>%s[^>]*</w:t>' % kind, seg):
            m = re.search(r'<w:t>([^<]*)</w:t>', seg)
            txt = re.sub(r'\s+', ' ', ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', seg, re.S)))
            if txt.startswith(kind):
                vals = re.findall(r'<w:t>(\d+)</w:t>', seg)
                out.append((i, vals))
    return out

for kind, total in (("图", 45), ("表", 51)):
    lst = scan(kind)
    nums = []
    RPR_OPT = r'(?:<w:rPr>(?:(?!</w:rPr>).)*?</w:rPr>)?'
    pat = re.compile(
        r'<w:instrText[^>]*> SEQ ' + kind + r' \\\* ARABIC \\s 1 </w:instrText></w:r>'
        r'<w:r>' + RPR_OPT + r'<w:fldChar w:fldCharType="separate"/></w:r>'
        r'<w:r>' + RPR_OPT + r'<w:t[^>]*>(\d+)</w:t>')
    miss = []
    for i, vals in lst:
        m = pat.search(db[Sb[i][0]:Sb[i][1]])
        if m:
            nums.append(int(m.group(1)))
        else:
            nums.append(-1)
            miss.append(i)
    chk("第3章 %s 题注 %d 个" % (kind, total), len(lst) == total, "实际 %d" % len(lst))
    chk("第3章 %s 缓存全部可解析" % kind, not miss, str(miss[:6]))
    chk("第3章 %s 编号 1..%d 连续" % (kind, total), nums == list(range(1, total + 1)), str(nums))

print("=== 6. 5 段题注结构与标准题注同构 ===")
# 标准题注段2327 的结构签名
def sig(idx):
    seg = db[Sb[idx][0]:Sb[idx][1]]
    return ('<w:bookmarkStart ' in seg and ' STYLEREF 1 \\s ' in seg
            and ' SEQ 图 \\* ARABIC \\s 1 ' in seg or ' SEQ 表 \\* ARABIC \\s 1 ' in seg)
for i in (2330, 4793, 4894, 4933, 5051):
    seg = db[Sb[i][0]:Sb[i][1]]
    ok = (' STYLEREF 1 \\s ' in seg and re.search(r' SEQ [图表] \\\* ARABIC \\s 1 ', seg)
          and '<w:bookmarkStart' in seg and '<w:bookmarkEnd' in seg
          and ' REF ' not in seg)
    chk("段%d 为标准题注（SEQ+书签，无 REF）" % i, bool(ok))

print("=== 7. 正文引用改指 ===")
MAP = {4792: "_Ref900302", 4893: "_Ref900303", 4932: "_Ref900304", 5050: "_Ref900305", 5108: "_Ref900305"}
for i, tgt in MAP.items():
    seg = db[Sb[i][0]:Sb[i][1]]
    chk("段%d REF -> %s" % (i, tgt), ("REF %s \\h" % tgt) in seg)

print("=== 8. 题注样式 ===")
for i in (5169, 5176, 6650):
    seg = db[Sb[i][0]:Sb[i][1]]
    chk("段%d 补 caption 样式" % i, '<w:pStyle w:val="12"/>' in seg[:200])
for i in (5921, 5990):
    seg = db[Sb[i][0]:Sb[i][1]]
    chk("段%d 去 caption 样式" % i, '<w:pStyle w:val="12"/>' not in seg)
chk("段7765 已删除（原空段）", len(Sb) == 8273)

print("=== 9. 标题直接格式 ===")
bad = []
for i in (2323, 453, 2246, 2254, 2262, 2289, 591, 999, 1115, 7653, 345, 1668, 2132, 2317):
    seg = db[Sb[i][0]:Sb[i][1]]
    if re.search(r'<w:keepNext w:val="0"/>|<w:keepLines w:val="0"/>|<w:jc w:val="both"/>', seg):
        bad.append(i)
for i in (1669, 1676, 1821, 1976, 2126, 2141, 2190, 2218):
    seg = db[Sb[i][0]:Sb[i][1]]
    pp = re.search(r'<w:pPr>.*?</w:pPr>', seg, re.S)
    if pp and '<w:color w:val="000000"/>' in pp.group(0):
        bad.append(i)
chk("CR-305 22 段标题直接格式已清", not bad, str(bad))

print("=== 10. 段2330 run 硬字体 ===")
seg = db[Sb[2330][0]:Sb[2330][1]]
chk("段2330 无 Arial/宋体硬字体", 'w:ascii="Arial"' not in seg and 'w:ascii="宋体"' not in seg)

print("=== 11. 域三态平衡 ===")
for tag in ("<w:fldChar w:fldCharType=\"begin\"/>", "<w:fldChar w:fldCharType=\"separate\"/>", "<w:fldChar w:fldCharType=\"end\"/>"):
    ca = da.count(tag); cb = db.count(tag)
    chk("fldChar %s 计数不变" % tag[38:45], ca == cb, "%d vs %d" % (ca, cb))

print("=== 12. 批注部件 ===")
cm = zb["word/comments.xml"].decode("utf-8")
cx = zb["word/commentsExtended.xml"].decode("utf-8")
chk("comments.xml 6 条", cm.count("<w:comment ") == 6, str(cm.count("<w:comment ")))
chk("commentsExtended 6 条", cx.count("<w15:commentEx ") == 6, str(cx.count("<w15:commentEx ")))
ct = zb["[Content_Types].xml"].decode("utf-8")
chk("Content_Types 有 comments 声明", "comments+xml" in ct and "commentsExtended+xml" in ct)
rels = zb["word/_rels/document.xml.rels"].decode("utf-8")
chk("rels 有 comments 关系", "relationships/comments\"" in rels and "commentsExtended\"" in rels)
chk("commentRange 与 comment 一致",
    db.count("<w:commentRangeStart") == cm.count("<w:comment "))

print("=== 13. document.xml 里批注 id 与 comments.xml 对齐 ===")
cids_doc = sorted(set(re.findall(r'<w:commentReference w:id="(\d+)"', db)))
cids_xml = sorted(re.findall(r'<w:comment w:id="(\d+)"', cm))
chk("批注 id 集合一致", cids_doc == cids_xml, "%s vs %s" % (cids_doc, cids_xml))

print("=== 14. 全部 XML 部件 well-formed ===")
from xml.etree import ElementTree as ET
badx = []
for n, data in zb.items():
    if n.endswith(".xml") or n.endswith(".rels"):
        try:
            ET.fromstring(data)
        except Exception as ex:
            badx.append((n, str(ex)[:60]))
chk("全部 XML well-formed", not badx, str(badx[:3]))

print("\n=== 汇总: %d PASS / %d FAIL ===" % (PASS, FAIL))
sys.exit(0 if FAIL == 0 else 2)
