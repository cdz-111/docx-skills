# -*- coding: utf-8 -*-
"""交叉引用三合一修复器（对「被别人改过」的稿子做收尾）

  A. 过宽 _Ref 书签（包住整段题注）→ 收窄为 Word 标准（只包 "图/表 " + SEQ 域）   [见 §15]
  B. 按 REF 语义重算全部 REF 域缓存                                              [见 §15]
  C. 清单条目编号按「PAGEREF 锚点 → 题注号」反查校正                              [见 §12.1]

三个步骤都是「只改 word/document.xml、其余部件逐字节照抄」，可安全反复跑（幂等）。

用法: python fix_cross_refs.py <in.docx> <out.docx> [--dry]
"""

import zipfile, re, sys, os
from collections import Counter
import xml.etree.ElementTree as ET

DRY = '--dry' in sys.argv
SRC, DST = sys.argv[1], sys.argv[2]

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
T_BEGIN = '<w:fldChar w:fldCharType="begin"/>'
T_SEP = '<w:fldChar w:fldCharType="separate"/>'
T_END = '<w:fldChar w:fldCharType="end"/>'
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.S)


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
                st = None
                depth = 0
        else:
            if depth == 0:
                st = m.start()
            depth += 1
    return res


def render(frag):
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
    return ''.join(m.group(1) for m in WT.finditer(''.join(out)))


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


z = zipfile.ZipFile(SRC)
x = z.read('word/document.xml').decode('utf-8')
infos = z.infolist()
z.close()

# ==================== A. 收窄过宽 _Ref 书签 ====================
starts, ends = bm_index(x)
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
    se = span.find(T_END)
    if se < 0 or (s1 + se) >= toc:
        continue
    plan.append(dict(name=name, s0=s0, e0=e0, e1=e1, toc=toc))
plan.sort(key=lambda r: r['s0'])
for r in plan:
    x = x[:r['toc']] + x[r['e0']:r['e1']] + x[r['toc']:r['e0']] + x[r['e1']:]
ET.fromstring(x.encode('utf-8'))
print('A. 过宽书签 %d 个 → 已收窄（XML OK）' % len(plan))

# ==================== B. 按语义重算 REF 缓存 ====================
st2, en2 = bm_index(x)
sp = {n: (s, en2[k][0]) for k, (n, s, _e) in st2.items()
      if k in en2 and n.startswith('_Ref')}

P = paras(x)
patches, changes = [], 0
for i, (a, b) in enumerate(P):
    q = x[a:b]
    fields = []
    for m in re.finditer(re.escape(T_BEGIN), q):
        seg = q[m.end():]
        em = re.search(re.escape(T_END), seg)
        if not em:
            continue
        body = seg[:em.start()]
        im = re.search(r'<w:instrText[^>]*>(.*?)</w:instrText>', body, re.S)
        if not im:
            continue
        instr = im.group(1).strip()
        if not instr.upper().startswith('REF'):
            continue
        tgt = re.search(r'REF\s+(\S+)', instr).group(1)
        if tgt not in sp:
            continue
        sm = body.find(T_SEP)
        if sm < 0:
            continue
        fields.append(dict(tgt=tgt, cstart=m.end() + sm + len(T_SEP),
                           cend=m.end() + em.start()))
    if not fields:
        continue
    orig = q
    for f in sorted(fields, key=lambda d: d['cstart'], reverse=True):
        want = render(x[sp[f['tgt']][0]:sp[f['tgt']][1]]).strip()
        region = q[f['cstart']:f['cend']]
        cur = ''.join(m.group(1) for m in WT.finditer(region))
        if cur == want:
            continue
        # 增量重建（多 run 时不能靠旧偏移逐个替换）
        out, last, j = [], 0, 0
        for t in WT.finditer(region):
            out.append(region[last:t.start()])
            out.append('<w:t xml:space="preserve">%s</w:t>' % (want if j == 0 else ''))
            last = t.end(); j += 1
        if j == 0:
            continue
        out.append(region[last:])
        q = q[:f['cstart']] + ''.join(out) + q[f['cend']:]
        changes += 1
    if q != orig:
        patches.append((a, b, q))
for a, b, q in sorted(patches, key=lambda t: t[0], reverse=True):
    x = x[:a] + q + x[b:]
ET.fromstring(x.encode('utf-8'))
print('B. REF 缓存改写 %d 处（XML OK）' % changes)

# ==================== C. 清单编号：锚点反查题注 ====================
# 1) _Toc####  →  所属题注的 (kind, num)
toc2cap = {}
for a, b in paras(x):
    q = x[a:b]
    ins = [t.strip() for t in re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', q)]
    if not any(t.startswith('SEQ 图') or t.startswith('SEQ 表') for t in ins):
        continue
    r = render(q)
    m = re.match(r'^\s*([图表])\s*(\d+)', r)
    if not m:
        continue
    for nm in re.findall(r'<w:bookmarkStart\b[^>]*w:name="(_Toc[^"]*)"', q):
        toc2cap[nm] = (m.group(1), int(m.group(2)))

# 2) 清单条目：PAGEREF 锚点 → 应显示的编号
edits = []
miss = 0
stat = Counter()
for i, (a, b) in enumerate(P):
    q = x[a:b]
    if 'PAGEREF _Toc' not in q:
        continue
    tgts = re.findall(r'PAGEREF\s+(_Toc\w+)', q)
    if len(tgts) != 1:
        continue
    tgt = tgts[0]
    if tgt not in toc2cap:
        continue                       # 正文目录条目，跳过
    kind, num = toc2cap[tgt]
    stat[kind] += 1
    head = q[:q.find('PAGEREF _Toc')]
    ts = list(WT.finditer(head))
    if not ts:
        miss += 1; continue
    joined = ''.join(t.group(1) for t in ts)
    m = re.search(r'([图表])\s*(\d+)', joined)
    if not m:
        miss += 1; continue
    if m.group(1) != kind:
        print('   !! p%d 锚点 %s 题注是 %s，清单写的却是 %s' % (i, tgt, kind, m.group(1)))
    gs, ge = m.start(2), m.end(2)
    # 定位数字落在哪个 <w:t>
    base = 0
    for t in ts:
        c = t.group(1)
        if base <= gs and ge <= base + len(c):
            ls, le = gs - base, ge - base
            if c[ls:le] == str(num):
                break
            newc = c[:ls] + str(num) + c[le:]
            edits.append((a + t.start(1), a + t.end(1), newc, i, kind, c, newc))
            break
        base += len(c)
print('C. 清单条目按锚点命中 图=%d 表=%d ；无法定位编号 %d'
      % (stat['图'], stat['表'], miss))
print('   需改编号: %d 条' % len(edits))
for s, e, v, i, kind, old, new in edits[:8]:
    print('      p%-5d %s  %r -> %r' % (i, kind, old, new))
if len(edits) > 8:
    print('      ... 共 %d' % len(edits))
for s, e, v, i, kind, old, new in sorted(edits, key=lambda r: r[0], reverse=True):
    x = x[:s] + v + x[e:]

xml_bytes = x.encode('utf-8')
ET.fromstring(xml_bytes)
bms = len(re.findall(r'<w:bookmarkStart\b', x)); bme = len(re.findall(r'<w:bookmarkEnd\b', x))
fld = Counter(re.findall(r'<w:fldChar w:fldCharType="(\w+)"/>', x))
print('\n校验: bookmarkStart=%d bookmarkEnd=%d  fldChar=%s' % (bms, bme, dict(fld)))
assert bms == bme and fld['begin'] == fld['separate'] == fld['end']

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
