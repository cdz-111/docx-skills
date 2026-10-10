# -*- coding: utf-8 -*-
"""侦察正文硬写「图N/表N」的 run 结构，分类决定改写策略。"""
import zipfile, re, sys

DOC = sys.argv[1] if len(sys.argv) > 1 else None
if not DOC:
    sys.exit('用法: recon_hardref.py <in.docx>')

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.S)
FLD = re.compile(r'<w:fldChar w:fldCharType="(begin|separate|end)"/>')


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


def cached_regions(q):
    """返回域缓存区间列表 [(s,e)]（separate..end 之间，含嵌套）"""
    stack, out = [], []
    for m in FLD.finditer(q):
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
    """可见 <w:t> 段：不在任何域缓存区间内。返回 [(t_start, t_end, text, abs_inner_s, abs_inner_e)]"""
    regs = cached_regions(q)

    def inside(pos):
        return any(s <= pos < e for s, e in regs)

    out = []
    for m in WT.finditer(q):
        if inside(m.start()):
            continue
        out.append((m.start(), m.end(), m.group(1), m.start(1), m.end(1)))
    return out


def run_of(q, pos):
    """返回包含 pos 的最内层 <w:r ...>...</w:r> 区间"""
    best = None
    for m in re.finditer(r'<w:r(?: [^>]*)?>.*?</w:r>|<w:r(?: [^>]*)?/>', q, re.S):
        if m.start() <= pos < m.end():
            if best is None or (m.end() - m.start()) < (best[1] - best[0]):
                best = (m.start(), m.end(), m.group(0))
    return best


z = zipfile.ZipFile(DOC)
x = z.read('word/document.xml').decode('utf-8')
z.close()
print('文件: %s  (doc.xml %d chars)' % (DOC.split('\\')[-1], len(x)))

P = paras(x)
stats = {'A_整段就是编号': [], 'B_段内子串': [], 'C_跨多段': [], 'D_未找到': []}
n = 0
for i, (a, b) in enumerate(P):
    q = x[a:b]
    segs = visible_segs(q)
    if not segs:
        continue
    joined = ''.join(s[2] for s in segs)
    if not re.search(r'[图表]\s*\d+', joined):
        continue
    # 是否题注段（含 SEQ）
    ins = re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', q)
    if any(t.strip().startswith('SEQ 图') or t.strip().startswith('SEQ 表') for t in ins):
        continue
    # 建 char → seg 索引
    idx = []
    for k, s in enumerate(segs):
        for _ in range(len(s[2])):
            idx.append(k)
    for m in re.finditer(r'([图表])\s*(\d+)(?!\d)', joined):
        n += 1
        cs, ce = m.start(), m.end()
        k0, k1 = idx[cs], idx[ce - 1]
        if k0 == k1:
            s = segs[k0]
            if m.group(0) == s[2]:
                cls = 'A_整段就是编号'
            else:
                cls = 'B_段内子串'
        else:
            cls = 'C_跨多段'
        r = run_of(q, segs[k0][0])
        stats[cls].append(dict(pi=i, txt=m.group(0), cs=cs, ce=ce,
                               seg0=segs[k0][2], k0=k0, k1=k1,
                               nseg=len(segs), segs=[t[2] for t in segs],
                               ctx=joined[max(0, cs - 18):ce + 14],
                               runlen=(r[1] - r[0]) if r else None))
print('\n共 %d 处' % n)
for cls, lst in stats.items():
    print('  %-14s %d' % (cls, len(lst)))
print('\n明细（前 40）：')
for cls, lst in stats.items():
    for r in lst[:40]:
        print('  [%s] p%-5d %-6r  seg0=%r  段内可见段数=%d  所在run长=%s'
              % (cls[:1], r['pi'], r['txt'], r['seg0'], r['nseg'], r['runlen']))
        print('        ctx=%r' % r['ctx'])
