# -*- coding: utf-8 -*-
"""步骤5：校验。
用法: python 5_verify.py <src.docx> <dst.docx> <workdir>
"""
import re, zipfile, sys, collections
import xml.etree.ElementTree as ET

SRC, DST, W = sys.argv[1], sys.argv[2], sys.argv[3]

SUB = '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜᵢᵣᵤᵥⱼ'
SUP = '⁰¹²³⁴⁵⁶⁷⁸⁹ⁱⁿᵀᵃᵇᶜᵈᵉᶠᵍʰʲᵏˡᵐᵒᵖʳˢᵗᵘᵛʷˣʸᶻ'

a, b = zipfile.ZipFile(SRC), zipfile.ZipFile(DST)
ok = True

na = [(i.filename, i.file_size, i.compress_type) for i in a.infolist()]
nb = [(i.filename, i.file_size, i.compress_type) for i in b.infolist()]
diffs = [(x, y) for x, y in zip(na, nb) if x != y]
print('[1] 条目数 %d -> %d, 差异条目 %d (应仅 word/document.xml)' % (len(na), len(nb), len(diffs)))
for d in diffs:
    print('     ', d[0], '->', d[1])
if len(diffs) != 1 or diffs[0][1][0] != 'word/document.xml':
    ok = False

x = b.read('word/document.xml')
try:
    ET.fromstring(x)
    print('[2] document.xml XML 合法 ✓')
except Exception as e:
    print('[2] XML 非法 ✗', e); ok = False
t = x.decode('utf-8')

left = [(c) for c in re.findall(r'<w:t[^>]*>([^<]*)</w:t>', t) for c in c if c in SUB or c in SUP]
print('[3] 残留伪角标字符: %d (应为 0)' % len(left))
if left:
    print('     ', collections.Counter(left)); ok = False

print('[4] 剩余 <w:drawing>: %d ; <m:oMath>: %d ; 下标run %d ; 上标run %d'
      % (t.count('<w:drawing>'), t.count('<m:oMath>'),
         t.count('w:val="subscript"'), t.count('w:val="superscript"')))

paras = re.findall(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', t, re.S)
sec, per = '', collections.defaultdict(list)
for p in paras:
    if '<w:pStyle' in p:
        m = re.match(r'\s*(\d+(?:\.\d+)*)\s',
                     ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)))
        if m:
            sec = m.group(1)
    if 'w:tab w:val="center"' in p:
        lab = [v for v in re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p) if re.fullmatch(r'\(\d+\)', v)]
        per[sec].append(lab[0] if lab else '??')

total = sum(len(v) for v in per.values())
print('[5] 展示公式 %d 条，各节编号：' % total)
bad = []
for s in per:
    got = [int(v.strip('()')) for v in per[s]]
    good = got == list(range(1, len(got) + 1))
    if not good:
        bad.append(s)
    print('     %-10s %s  %s' % (s, 'OK' if good else 'BAD', got))
if bad:
    ok = False

emp = [1 for m in re.findall(r'<m:oMath>.*?</m:oMath>', t, re.S) if '<m:t>' not in m]
print('[6] 空公式: %d (应为 0)' % len(emp))
if emp:
    ok = False

print()
print('结论:', 'ALL CHECKS PASSED' if ok else 'HAS PROBLEMS —— 见上面标记 ✗ 的项')
print('建议再跑一次：pandoc -f docx -t markdown 新文件.docx，检查公式变成 $...$、角标变成 ~sub~ / ^sup~。')
