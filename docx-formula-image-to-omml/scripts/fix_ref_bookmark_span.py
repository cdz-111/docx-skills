# -*- coding: utf-8 -*-
"""把过宽的 _Ref 书签（覆盖整段题注）收窄为 Word 标准样式
   （只包 "图/表 " + SEQ 域），随后按书签语义重算全部 REF 缓存。

背景：为硬写题注补 _Ref 书签时若包住整段题注，缓存写对也没用 ——
      Word 全量更新域（F9 / 打印 / 导出 PDF）后会膨胀成「图 5 标题全文」。
      本脚本把 bookmarkEnd 移到 _Toc bookmarkStart 之前，使两者相邻不重叠。

判据：_Ref 跨度 XML 长度 > 350 且内部含 _Toc bookmarkStart  →  过宽。

用法: python fix_ref_bookmark_span.py <in.docx> <out.docx> [--dry]
"""

import zipfile, re, sys, os
from collections import Counter
import xml.etree.ElementTree as ET

DRY = '--dry' in sys.argv
SRC, DST = sys.argv[1], sys.argv[2]

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')


def paras(x):
    res, depth, start = [], 0, None
    for m in TAGP.finditer(x):
        s = m.group(0)
        if s.endswith('/>'):
            if depth == 0:
                res.append((m.start(), m.end()))
        elif s == '</w:p>':
            depth -= 1
            if depth <= 0:
                if start is not None:
                    res.append((start, m.end()))
                start = None
                depth = 0
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return res


T_BEGIN = '<w:fldChar w:fldCharType="begin"/>'
T_SEP = '<w:fldChar w:fldCharType="separate"/>'
T_END = '<w:fldChar w:fldCharType="end"/>'


def render(frag):
    """渲染可见文本：域取 separate..end 之间的缓存"""
    out, pos = [], 0
    while True:
        a = frag.find(T_BEGIN, pos)
        if a < 0:
            out.append(frag[pos:]); break
        sm = frag.find(T_SEP, a)
        em = frag.find(T_END, a)
        if sm < 0 or em < 0 or sm > em:
            out.append(frag[pos:a])
            if em < 0:
                out.append(frag[a:]); break
            out.append(frag[a:em])
            pos = em
            continue
        out.append(frag[pos:a])
        out.append(frag[sm + len(T_SEP):em])
        pos = em + len(T_END)
    return ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', ''.join(out), re.S))


def bm_index(x):
    st, en = {}, {}
    for m in re.finditer(r'<w:bookmarkStart\b[^>]*?/>', x):
        t = m.group(0)
        i = re.search(r'w:id="(\d+)"', t); n = re.search(r'w:name="([^"]*)"', t)
        if i and n:
            st[i.group(1)] = (n.group(1), m.start(), m.end())
    for m in re.finditer(r'<w:bookmarkEnd\b[^>]*?/>', x):
        i = re.search(r'w:id="(\d+)"', m.group(0))
        if i:
            en[i.group(1)] = (m.start(), m.end())
    return st, en


# ---------- 读入 ----------
z = zipfile.ZipFile(SRC)
x = z.read('word/document.xml').decode('utf-8')
infos = z.infolist()
z.close()

starts, ends = bm_index(x)

# ---------- 1) 找出过宽 _Ref 书签 ----------
plan = []
for bid, (name, s0, s1) in starts.items():
    if not name.startswith('_Ref') or bid not in ends:
        continue
    e0, e1 = ends[bid]
    span = x[s1:e0]
    if 'SEQ' not in span:
        continue
    tm = re.search(r'<w:bookmarkStart\b[^>]*w:name="(_Toc[^"]*)"', span)
    if not tm:
        continue
    toc = s1 + tm.start()
    if toc >= e0:
        continue
    se = span.find('<w:fldChar w:fldCharType="end"/>')
    if se < 0 or (s1 + se) >= toc:
        continue
    plan.append(dict(name=name, bid=bid, s0=s0, s1=s1, e0=e0, e1=e1, toc=toc))

print('过宽书签(待收窄): %d' % len(plan))
for r in plan[:6]:
    print('   %-13s id=%-5s  span %d -> %d' % (r['name'], r['bid'], r['e0'] - r['s1'], r['toc'] - r['s1']))
if len(plan) > 6:
    print('   ...（共 %d）' % len(plan))

# ---------- 2) 长度守恒地移动 bookmarkEnd ----------
plan.sort(key=lambda r: r['s0'])
for r in plan:
    toc, e0, e1 = r['toc'], r['e0'], r['e1']
    assert x[e1 - 11:e1].endswith('>') and '<w:bookmarkEnd' in x[e0:e1], r['name']
    x = x[:toc] + x[e0:e1] + x[toc:e0] + x[e1:]
ET.fromstring(x.encode('utf-8'))
print('收窄后 XML 解析: OK')

# ---------- 3) 按书签语义重算 REF 缓存 ----------
st2, en2 = bm_index(x)
sp = {}
for k, (n, s0, s1) in st2.items():
    if k in en2 and n.startswith('_Ref'):
        sp[n] = (s1, en2[k][0])

P = paras(x)
changes = []
patches = []          # (start, end, newtext) 按段落区间替换，保留段外内容
for i, (a, b) in enumerate(P):
    q = x[a:b]
    fields = []
    for m in re.finditer(T_BEGIN, q):
        seg = q[m.end():]
        em = re.search(r'<w:fldChar w:fldCharType="end"/>', seg)
        if not em:
            continue
        body = seg[:em.start()]
        im = re.search(r'<w:instrText[^>]*>(.*?)</w:instrText>', body, re.S)
        if not im:
            continue
        instr = im.group(1).strip()
        if not instr.upper().startswith('REF'):
            continue
        tm = re.search(r'REF\s+(\S+)', instr)
        if not tm or tm.group(1) not in sp:
            continue
        sm = body.find(T_SEP)
        if sm < 0:
            continue
        fields.append(dict(tgt=tm.group(1),
                           cstart=m.end() + sm + len(T_SEP),
                           cend=m.end() + em.start()))
    orig = q
    for f in sorted(fields, key=lambda d: d['cstart'], reverse=True):
        want = render(x[sp[f['tgt']][0]:sp[f['tgt']][1]]).strip()
        region = q[f['cstart']:f['cend']]
        cur = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', region, re.S))
        if cur == want:
            continue
        ts = list(re.finditer(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', region))
        if not ts:
            continue
        r2 = region
        for j, t in enumerate(ts):
            v = want if j == 0 else ''
            r2 = r2[:t.start()] + '<w:t xml:space="preserve">%s</w:t>' % v + r2[t.end():]
        q = q[:f['cstart']] + r2 + q[f['cend']:]
        changes.append((i, f['tgt'], cur, want))
    if q != orig:
        patches.append((a, b, q))

for a, b, q in sorted(patches, key=lambda t: t[0], reverse=True):
    x = x[:a] + q + x[b:]

print('REF 缓存改写: %d 处' % len(changes))
for c in changes[:30]:
    print('   p%-5d %-13s %r -> %r' % (c[0], c[1], c[2], c[3]))

xml_bytes = x.encode('utf-8')
ET.fromstring(xml_bytes)
bms = len(re.findall(r'<w:bookmarkStart\b', x))
bme = len(re.findall(r'<w:bookmarkEnd\b', x))
fld = Counter(re.findall(r'<w:fldChar w:fldCharType="(\w+)"/>', x))
print('校验: bookmarkStart=%d bookmarkEnd=%d  fldChar=%s' % (bms, bme, dict(fld)))
assert bms == bme
assert fld['begin'] == fld['separate'] == fld['end']

if DRY:
    print('[dry] 未写文件'); sys.exit(0)

zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for it in infos:
    data = zin.read(it.filename)
    if it.filename == 'word/document.xml':
        data = xml_bytes
    ni = zipfile.ZipInfo(it.filename, date_time=it.date_time)
    ni.compress_type = it.compress_type
    ni.external_attr = it.external_attr
    ni.internal_attr = it.internal_attr
    ni.create_system = it.create_system
    ni.flag_bits = it.flag_bits
    zout.writestr(ni, data)
zout.close(); zin.close()
print('已写出: %s  (%d bytes)' % (DST, os.path.getsize(DST)))
