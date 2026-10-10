# -*- coding: utf-8 -*-
"""对比两个 docx 的可见文本与结构：确认改动范围 = 指定的 37 处引用。"""
import zipfile, re, sys
from collections import Counter
import xml.etree.ElementTree as ET

A, B = sys.argv[1], sys.argv[2]
TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
TB = '<w:fldChar w:fldCharType="begin"/>'
TS = '<w:fldChar w:fldCharType="separate"/>'
TE = '<w:fldChar w:fldCharType="end"/>'
NUM = re.compile(r'[图表]\s*\d+')


def paras(x):
    res, depth, st = [], 0, None
    for m in TAGP.finditer(x):
        s = m.group(0)
        if s.endswith('/>'):
            if depth == 0:
                res.append((m.start(), m.end()))
        elif s == '</w:p>':
            depth -= 1
            if depth <= 0:
                if st is not None:
                    res.append((st, m.end()))
                st = None; depth = 0
        else:
            if depth == 0:
                st = m.start()
            depth += 1
    return res


def render(frag):
    out, pos = [], 0
    while True:
        a = frag.find(TB, pos)
        if a < 0:
            out.append(frag[pos:]); break
        sm = frag.find(TS, a); em = frag.find(TE, a)
        if sm < 0 or em < 0 or sm > em:
            out.append(frag[pos:a])
            if em < 0:
                out.append(frag[a:]); break
            out.append(frag[a:em]); pos = em; continue
        out.append(frag[pos:a]); out.append(frag[sm + len(TS):em])
        pos = em + len(TE)
    return ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', ''.join(out), re.S))


def load(p):
    z = zipfile.ZipFile(p)
    x = z.read('word/document.xml').decode('utf-8')
    names = z.namelist()
    z.close()
    return x, names


xa, na = load(A)
xb, nb = load(B)
print('zip 条目: %d vs %d  相同=%s' % (len(na), len(nb), na == nb))
ET.fromstring(xb.encode('utf-8'))
print('staged XML 可解析 ✓')

Pa, Pb = paras(xa), paras(xb)
print('段落数: %d vs %d' % (len(Pa), len(Pb)))

if len(Pa) == len(Pb):
    diffs = []
    for i, ((a1, b1), (a2, b2)) in enumerate(zip(Pa, Pb)):
        t1, t2 = render(xa[a1:b1]), render(xb[a2:b2])
        if t1 != t2:
            diffs.append((i, t1, t2))
    print('\n可见文本有差异的段落: %d' % len(diffs))
    sites, ok = 0, True
    for i, t1, t2 in diffs:
        r1 = NUM.findall(t1)
        r2 = NUM.findall(t2)
        flat1 = [re.sub(r'\s', '', x) for x in r1]
        flat2 = [re.sub(r'\s', '', x) for x in r2]
        same = flat1 == flat2
        # 去掉编号后的骨架文本应完全一致
        skel1 = NUM.sub('@', t1)
        skel2 = NUM.sub('@', t2)
        if not same or skel1 != skel2:
            ok = False
        sites += max(len(r1), len(r2))
        print('  p%-5d %s -> %s   骨架一致=%s'
              % (i, r1, r2, skel1 == skel2))
    print('\n差异处数 = %d ; 仅空格差异、骨架不变 = %s' % (sites, ok))

ca = Counter(re.findall(r'<w:fldChar w:fldCharType="(\w+)"/>', xa))
cb = Counter(re.findall(r'<w:fldChar w:fldCharType="(\w+)"/>', xb))
print('\nfldChar: %s -> %s' % (dict(ca), dict(cb)))
print('REF 域: %d -> %d' % (len(re.findall(r'REF\s+_Ref\d+', xa)),
                           len(re.findall(r'REF\s+_Ref\d+', xb))))
print('SEQ 域: %d -> %d' % (len(re.findall(r'SEQ\s+[图表]', xa)),
                           len(re.findall(r'SEQ\s+[图表]', xb))))
print('书签: %d/%d -> %d/%d' % (
    len(re.findall(r'<w:bookmarkStart\b', xa)), len(re.findall(r'<w:bookmarkEnd\b', xa)),
    len(re.findall(r'<w:bookmarkStart\b', xb)), len(re.findall(r'<w:bookmarkEnd\b', xb))))
# 表格/图片计数
for tag in ['<w:tbl>', '<w:drawing>', '<w:tblGrid>']:
    print('%s: %d -> %d' % (tag, xa.count(tag), xb.count(tag)))
print('document.xml 长度: %d -> %d (+%d)' % (len(xa), len(xb), len(xb) - len(xa)))
# 未闭合 fldChar 检查
print('begin/separate/end 配平:',
      cb['begin'] == cb['separate'] == cb['end'] == sum(1 for _ in re.finditer(TB, xb)))
