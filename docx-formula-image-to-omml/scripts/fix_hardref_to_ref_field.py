# -*- coding: utf-8 -*-
r"""把正文残留的纯文本「图N/表N」升级为真正的 REF 交叉引用域。

生成: <w:fldChar begin> REF _Ref#### \h <separate> 缓存 <end>
  · 目标书签 (kind,num)->_Ref#### 由题注段（含 SEQ 图/表）反查，唯一、无歧义
  · 缓存文本取题注自身的渲染文本（如「图 16」），保证 F9 前后一致
  · 按「受影响 run 区间」整体重建，保留原 run 的 rPr（字体/加粗等不变）
  · 区间内若含域/图形/换行等非文本成分则拒绝该处，改为人工处理

用法: python fx_hardref.py <in.docx> <out.docx> [--dry]
"""
import zipfile, re, sys, os, json
from collections import Counter
import xml.etree.ElementTree as ET

DRY = '--dry' in sys.argv
argv = [a for a in sys.argv[1:] if not a.startswith('--')]
SRC, DST = argv[0], argv[1]

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.S)
FLDC = re.compile(r'<w:fldChar w:fldCharType="(begin|separate|end)"/>')
RPRRE = re.compile(r'<w:rPr>.*?</w:rPr>', re.S)
NUMRE = re.compile(r'([图表])\s*(\d+)(?!\d)')
RUNRE = re.compile(r'<w:r(?: [^>]*)?>.*?</w:r>|<w:r(?: [^>]*)?/>', re.S)
BADRE = re.compile(r'<w:fldChar|<w:instrText|<w:drawing|<w:pict|<w:object|'
                   r'<w:br\b|<w:tab\b|<w:cr\b|<w:sym\b|<w:noBreakHyphen|'
                   r'<w:softHyphen|<w:footnoteReference|<w:endnoteReference|'
                   r'<w:commentReference|<w:proofErr|<w:bookmarkStart|'
                   r'<w:bookmarkEnd')


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


def cached_regions(q):
    stack, out = [], []
    for m in FLDC.finditer(q):
        k = m.group(1)
        if k == 'begin':
            stack.append(None)
        elif k == 'separate':
            if stack and stack[-1] is None:
                stack[-1] = m.end()
        else:
            if stack:
                s = stack.pop()
                if s is not None:
                    out.append((s, m.start()))
    return out


def visible_segs(q):
    regs = cached_regions(q)
    def inside(p):
        return any(s <= p < e for s, e in regs)
    return [(m.start(), m.end(), m.group(1), m.start(1), m.end(1))
            for m in WT.finditer(q) if not inside(m.start())]


def run_of(q, pos):
    best = None
    for m in RUNRE.finditer(q):
        if m.start() <= pos < m.end():
            if best is None or (m.end() - m.start()) < (best[1] - best[0]):
                best = (m.start(), m.end(), m.group(0))
    return best


def rpr_of(runxml):
    m = RPRRE.search(runxml)
    return m.group(0) if m else ''


def ref_block(refname, rpr, cache):
    r = rpr
    return ('<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r>%s<w:instrText xml:space="preserve"> REF %s \\h </w:instrText></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r>'
            % (r, r, refname, r, r, cache, r))


z = zipfile.ZipFile(SRC)
x = z.read('word/document.xml').decode('utf-8')
infos = z.infolist()
z.close()

# ---------------- 1. (kind,num) -> (_Ref####, 渲染文本) ----------------
ends = {}
for m in re.finditer(r'<w:bookmarkEnd\b[^>]*?/>', x):
    i = re.search(r'w:id="(\d+)"', m.group(0))
    if i:
        ends[i.group(1)] = m.start()

CMAP = {}
for a, b in paras(x):
    q = x[a:b]
    if not re.search(r'SEQ\s+[图表]', q):
        continue
    vis = ''.join(WT.findall(q))
    mm = re.match(r'^\s*([图表])\s*(\d+)', vis)
    if not mm:
        continue
    kind, num = mm.group(1), int(mm.group(2))
    for m2 in re.finditer(r'<w:bookmarkStart\b[^>]*?/>', q):
        n = re.search(r'w:name="([^"]*)"', m2.group(0))
        i = re.search(r'w:id="(\d+)"', m2.group(0))
        if not (n and i) or not n.group(1).startswith('_Ref'):
            continue
        e0 = ends.get(i.group(1))
        if e0 is None:
            continue
        inner = x[a + m2.end():e0]
        CMAP[(kind, num)] = (n.group(1), ''.join(WT.findall(inner)), len(inner))
print('题注映射 %d 条' % len(CMAP))
NREF0 = len(re.findall(r'REF\s+_Ref\d+', x))
print('原有 REF 域 %d 个' % NREF0)

# ---------------- 2. 逐段定位并重建 ----------------
P = paras(x)
patches, report, skipped = [], [], []
for pi, (pa, pb) in enumerate(P):
    q = x[pa:pb]
    ins = re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', q)
    if any(t.strip().startswith('SEQ 图') or t.strip().startswith('SEQ 表') for t in ins):
        continue                                    # 题注段跳过
    segs = visible_segs(q)
    if not segs:
        continue
    jbase, tot = [], 0
    for s in segs:
        jbase.append(tot); tot += len(s[2])
    joined = ''.join(s[2] for s in segs)
    matches = list(NUMRE.finditer(joined))
    if not matches:
        continue
    seg_of = []
    for k, s in enumerate(segs):
        seg_of.extend([k] * len(s[2]))

    spans = []
    for m in matches:
        gs, ge = m.start(), m.end()
        k0, k1 = seg_of[gs], seg_of[ge - 1]
        p0 = segs[k0][3] + (gs - jbase[k0])
        p1 = segs[k1][3] + (ge - jbase[k1])
        r0 = run_of(q, p0); r1 = run_of(q, p1 - 1)
        if not r0 or not r1:
            skipped.append((pi, m.group(0), 'no_run')); continue
        spans.append(dict(gs=gs, ge=ge, txt=m.group(0), kind=m.group(1),
                          num=int(m.group(2)), rs=r0[0], re=r1[1]))
    if not spans:
        continue

    spans.sort(key=lambda d: d['rs'])
    groups, cur = [], None
    for sp in spans:
        if cur and sp['rs'] <= cur['re']:
            cur['re'] = max(cur['re'], sp['re']); cur['items'].append(sp)
        else:
            cur = dict(rs=sp['rs'], re=sp['re'], items=[sp]); groups.append(cur)

    pieces_all, bad = [], False
    for g in groups:
        reg = q[g['rs']:g['re']]
        if BADRE.search(reg):
            skipped.append((pi, [t['txt'] for t in g['items']], 'bad_content'))
            bad = True; continue
        ts = list(WT.finditer(reg))
        # 区间内应恰好覆盖若干完整可见 seg
        inner_segs = [k for k, s in enumerate(segs)
                      if s[0] >= g['rs'] and s[1] <= g['re']]
        if not inner_segs:
            skipped.append((pi, [t['txt'] for t in g['items']], 'no_seg'))
            bad = True; continue
        f0, l0 = inner_segs[0], inner_segs[-1]
        jb = jbase[f0]
        want_len = jbase[l0] + len(segs[l0][2]) - jb
        regtext = ''.join(t.group(1) for t in ts)
        if len(regtext) != want_len or len(ts) != len(inner_segs):
            skipped.append((pi, [t['txt'] for t in g['items']], 'seg_mismatch'))
            bad = True; continue
        owner = []
        for ti, t in enumerate(ts):
            owner.extend([ti] * len(t.group(1)))

        loc = sorted(g['items'], key=lambda d: d['gs'])
        outs = sorted((it['gs'] - jb, it['ge'] - jb, it) for it in loc)
        if any(not (0 <= lo < hi <= len(regtext)) for lo, hi, _ in outs):
            skipped.append((pi, [t['txt'] for t in loc], 'offset'))
            bad = True; continue

        out, cur_pos, baditem = [], 0, False
        for lo, hi, it in outs:
            if lo > cur_pos:
                rpr = rpr_of(run_of(reg, ts[owner[cur_pos]].start())[2])
                out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                           % (rpr, regtext[cur_pos:lo]))
            ref = CMAP.get((it['kind'], it['num']))
            if not ref:
                skipped.append((pi, it['txt'], 'no_bookmark')); baditem = True; break
            rprr = rpr_of(run_of(reg, ts[owner[lo]].start())[2])
            out.append(ref_block(ref[0], rprr, ref[1]))
            report.append((pi, it['txt'], ref[0], ref[1], rprr))
            cur_pos = hi
        if baditem:
            bad = True; continue
        if cur_pos < len(regtext):
            rpr = rpr_of(run_of(reg, ts[owner[cur_pos]].start())[2])
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                       % (rpr, regtext[cur_pos:]))
        pieces_all.append((g['rs'], g['re'], ''.join(out)))
    if bad or not pieces_all:
        continue
    nq = q
    for rs, re_, new in sorted(pieces_all, key=lambda t: t[0], reverse=True):
        nq = nq[:rs] + new + nq[re_:]
    if nq != q:
        patches.append((pa, pb, nq))

for pa, pb, nq in sorted(patches, key=lambda t: t[0], reverse=True):
    x = x[:pa] + nq + x[pb:]

xml_bytes = x.encode('utf-8')
ET.fromstring(xml_bytes)

print('\n改写 %d 处 REF 域：' % len(report))
for pi, txt, ref, cache, rp in report:
    print('  p%-5d %-8s -> %-14s cache=%-7s rPr=%s' % (pi, txt, ref, cache, rp or '(无)'))
if skipped:
    print('\n!! 跳过 %d 处（需人工）:' % len(skipped))
    for s in skipped:
        print('   ', s)

bms = len(re.findall(r'<w:bookmarkStart\b', x)); bme = len(re.findall(r'<w:bookmarkEnd\b', x))
fld = Counter(re.findall(r'<w:fldChar w:fldCharType="(\w+)"/>', x))
print('\n校验: bookmarkStart=%d bookmarkEnd=%d  fldChar=%s' % (bms, bme, dict(fld)))
assert bms == bme and fld['begin'] == fld['separate'] == fld['end']
nref = len(re.findall(r'REF\s+_Ref\d+', x))
print('REF 域总数 = %d（期望 %d）' % (nref, NREF0 + len(report)))
assert nref == NREF0 + len(report)

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
