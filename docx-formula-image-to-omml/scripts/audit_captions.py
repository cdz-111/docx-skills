# -*- coding: utf-8 -*-
"""9-28 全文图表题注结构化审计 + 交叉引用核对。
产出：report_caption928.txt（文本报告） / 图表题注审计.html（可读报告）
"""
import os
import sys
import zipfile, re, io, html as H

# ⚠️ 原来这里把输入/输出路径写死在「9-28 那一次」的文件上，导致本脚本
#    **无法被编排调用**（`docx-unify-pipeline` 的 S2 会直接 FileNotFoundError），
#    而且输出目录早已不存在。
#    现在改为三级回退，**不传参时行为与从前一致**：
#      ① 命令行第一个参数  ② 环境变量 DOCXFMT_DOC  ③ 原默认路径
DOC = (sys.argv[1] if len(sys.argv) > 1
       else os.environ.get('DOCXFMT_DOC')
       or r'C:\Users\14711\Desktop\中期科技报告9-28-公式规范化.docx')
OUTDIR = os.environ.get('DOCXFMT_OUTDIR') or os.path.dirname(
    os.path.abspath(DOC))
os.makedirs(OUTDIR, exist_ok=True)

z = zipfile.ZipFile(DOC)
xml = z.read('word/document.xml').decode('utf-8')

TAG = re.compile(r'<w:tbl>|</w:tbl>|<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')


def top_blocks(x):
    res, i, d, ts, cp = [], 0, 0, None, None
    while True:
        m = TAG.search(x, i)
        if not m:
            break
        s, i = m.group(0), m.end()
        if s == '<w:tbl>':
            d += 1
            if d == 1:
                ts = m.start()
        elif s == '</w:tbl>':
            d -= 1
            if d == 0:
                res.append(('TBL', ts, i)); ts = None
        elif s == '</w:p>':
            if cp is not None and d == 0:
                res.append(('P', cp, i))
            cp = None
        elif s.endswith('/>'):
            if d == 0:
                res.append(('P', m.start(), i))
        else:
            if cp is None:
                cp = m.start()
    res.sort(key=lambda r: r[1])
    return res


blks = top_blocks(xml)
kinds = [b[0] for b in blks]
blocks = [xml[s:e] for _, s, e in blks]

TOK = re.compile(
    r'<w:fldChar w:fldCharType="(\w+)"/>'
    r'|<w:instrText[^>]*>(.*?)</w:instrText>'
    r'|<w:t(?: [^>]*)?>(.*?)</w:t>'
    r'|<w:bookmarkStart[^>]*w:name="([^"]+)"'
    r'|<w:drawing>', re.S)


def segments(q):
    """按文档顺序返回 token 流：('T',文本) ('F',域指令,缓存) ('B',书签名) ('IMG',)"""
    out, stack, cur_instr, cur_disp, in_field = [], [], None, '', False
    for m in TOK.finditer(q):
        if m.group(1):                      # fldChar
            t = m.group(1)
            if t == 'begin':
                in_field, cur_instr, cur_disp = True, None, ''
            elif t == 'separate':
                pass
            elif t == 'end':
                out.append(('F', (cur_instr or '').strip(), cur_disp.strip()))
                in_field, cur_instr, cur_disp = False, None, ''
        elif m.group(2) is not None and in_field:
            cur_instr = (cur_instr or '') + m.group(2)
        elif m.group(3) is not None:
            if in_field:
                cur_disp += m.group(3)
            else:
                out.append(('T', m.group(3)))
        elif m.group(4):
            out.append(('B', m.group(4)))
        elif m.group(0) == '<w:drawing>':
            out.append(('IMG', None, None))
    return out


def vtext(q):
    return ''.join(t[1] for t in segments(q) if t[0] == 'T')


def vtext_full(q):
    """含域缓存显示值的完整可见文本"""
    s = ''
    for t in segments(q):
        if t[0] == 'T':
            s += t[1]
        elif t[0] == 'F':
            s += t[2]
    return s


def pstyle(q):
    m = re.search(r'<w:pStyle w:val="([^"]+)"', q)
    return m.group(1) if m else '-'


def bms(q):
    return [t[1] for t in segments(q) if t[0] == 'B']


# ---------- 逐块走，记录所属标题 ----------
heads = []
rows = []          # 题注
seq_counter = {'图': 0, '表': 0}
cur_head = '(正文之前)'
for i, q in enumerate(blocks):
    if kinds[i] == 'P':
        st = pstyle(q)
        if st in ('1', '2', '3', '4', '5'):
            t = vtext(q).strip()
            if t:
                cur_head = t
        segs = segments(q)
        flds = [s for s in segs if s[0] == 'F']
        capt = None
        for kk, s in enumerate(segs):
            if s[0] == 'F' and ('SEQ 图' in s[1] or 'SEQ 表' in s[1]):
                capt = (kk, s)
                break
        kind = None
        if capt:
            kind = 'SEQ' if 'SEQ 图' in capt[1][1] else 'SEQ'
            which = re.search(r'SEQ\s*([图表])', capt[1][1]).group(1)
        else:
            # 无 SEQ 域，看是否硬写
            m = re.match(r'^\s*([图表])\s*(\d+)', vtext(q))
            if m:
                which = m.group(1)
        if not capt and not (re.match(r'^\s*([图表])\s*\d+', vtext(q))):
            continue
        # ---- 组装 ----
        if capt:
            kk, f = capt
            cache = f[2]
            pre = ''.join(t[1] for t in segs[:kk] if t[0] == 'T')
            suf = ''
            for t in segs[kk + 1:]:
                if t[0] == 'T':
                    suf += t[1]
                elif t[0] == 'F':
                    suf += t[2]
            seq_counter[which] += 1
            after_f9 = seq_counter[which]
        else:
            pre, cache, suf, after_f9 = '', '', '', None
        vis = vtext_full(q).strip()
        m = re.match(r'^[图表]\s*(\d+)', vis)
        visnum = m.group(1) if m else None
        bmref = [b for b in bms(q) if b.startswith('_Ref')]
        bmtoc = [b for b in bms(q) if b.startswith('_Toc')]
        # 分类
        if not capt:
            cls = 'HARD'      # 纯手打编号，无域
        elif re.match(r'^\s*\d', suf):
            cls = 'LEFTOVER'  # 域后面还留着原来的数字
        elif cache != (visnum or ''):
            cls = 'CACHE'
        else:
            cls = 'OK'
        rows.append(dict(bi=i, kind=which, head=cur_head, cls=cls, pre=pre, cache=cache,
                         suf=suf, vis=vis, visnum=visnum, after_f9=after_f9,
                         style=pstyle(q), refbm=bmref, tocbm=bmtoc,
                         hasimg='<w:drawing>' in q))

# ---------- 交叉引用 ----------
bm2i = {}
for i, q in enumerate(blocks):
    for b in bms(q):
        bm2i[b] = i
refs = []
for i, q in enumerate(blocks):
    for t in segments(q):
        if t[0] == 'F' and t[1].startswith('REF'):
            m = re.search(r'REF\s+(\S+)', t[1])
            refs.append((i, m.group(1) if m else '?', t[2]))

out = io.StringIO()
P = lambda *a: print(*a, file=out)

P('=' * 90)
P('9-28 图表题注 + 交叉引用 审计')
P('=' * 90)
P('顶层块 %d（段落 %d / 表格 %d）' % (len(blocks), kinds.count('P'), kinds.count('TBL')))

for k in ('图', '表'):
    rr = [r for r in rows if r['kind'] == k]
    cnt = {}
    for r in rr:
        cnt[r['cls']] = cnt.get(r['cls'], 0) + 1
    P('\n' + '-' * 90)
    P('【%s】共 %d 个题注   OK=%d  域后残留数字=%d  缓存不符=%d  纯手打无域=%d'
      % (k, len(rr), cnt.get('OK', 0), cnt.get('LEFTOVER', 0), cnt.get('CACHE', 0), cnt.get('HARD', 0)))
    P('-' * 90)
    P('  %-6s %-5s %-9s %-8s %-6s %-8s %s' % ('块', '样式', '类别', '可见号', '缓存', 'F9后', '可见文本 / 结构'))
    for r in rr:
        flag = {'OK': '  ', 'LEFTOVER': '❗', 'CACHE': '⚠', 'HARD': '·'}[r['cls']]
        P('  %s%-5d %-5s %-9s %-8s %-6s %-8s %s'
          % (flag, r['bi'], r['style'], r['cls'], r['visnum'], r['cache'] or '-',
             r['after_f9'] or '-', r['vis'][:58]))
        if r['cls'] == 'LEFTOVER':
            P('          结构: %r + 【SEQ %s = %s】 + %r' % (r['pre'], r['kind'], r['cache'], r['suf'][:30]))
    P('  无 _Ref 书签的题注: %d 个 → %s'
      % (sum(1 for r in rr if not r['refbm']),
         [r['bi'] for r in rr if not r['refbm']][:24]))

P('\n' + '=' * 90)
P('交叉引用（REF 域）')
P('=' * 90)
P('  REF 域共 %d 个' % len(refs))
ok = mism = 0
broken = []
for i, tgt, disp in refs:
    j = bm2i.get(tgt)
    if j is None:
        broken.append((i, tgt, disp)); continue
    vis = vtext_full(blocks[j])
    m = re.match(r'^[图表]\s*(\d+)', vis.strip())
    capnum = m.group(1) if m else None
    shown = re.search(r'(\d+)', disp)
    shown = shown.group(1) if shown else None
    if capnum and shown and capnum != shown:
        mism += 1
        P('  ⚠ 显示号≠题注号: REF %s 显示 %r → 题注「%s」' % (tgt, disp, vis.strip()[:40]))
    else:
        ok += 1
P('  ✓ 一致 %d 个 | ⚠ 号不符 %d 个 | ❌ 断链 %d 个' % (ok, mism, len(broken)))
for i, tgt, disp in broken:
    P('     ❌ block %-5d REF %-12s 显示=%r' % (i, tgt, disp))
    P('        所在段: %s' % vtext(blocks[i]).strip()[-64:])

P('\n' + '=' * 90)
P('正文里「硬写」的图/表号（非交叉引用，改图号时不会被同步）')
P('=' * 90)
hard = []
for i, q in enumerate(blocks):
    if kinds[i] != 'P' or any(r['bi'] == i for r in rows):
        continue
    s = ''.join(('␣' if t[0] == 'F' else (t[1] or '')) for t in segments(q))
    for m in re.finditer(r'[图表]\s*\d+', s):
        hard.append((i, m.group(0), s[max(0, m.start() - 16):m.end() + 16]))
P('  共 %d 处' % len(hard))
for i, mm, ctx in hard:
    P('     block %-5d %-6s … %s …' % (i, mm, ctx.replace('\n', ' ')))

rep = out.getvalue()
open(os.path.join(OUTDIR, 'report_caption928.txt'), 'w',
     encoding='utf-8').write(rep)
import json
json.dump(dict(rows=rows, refs=[dict(bi=i, tgt=t, disp=d) for i, t, d in refs],
               broken=[dict(bi=i, tgt=t, disp=d) for i, t, d in broken],
               hard=[dict(bi=i, n=n, ctx=c) for i, n, c in hard],
               blocks=len(blocks), tables=kinds.count('TBL')),
          open(os.path.join(OUTDIR, 'cap928.json'), 'w',
               encoding='utf-8'), ensure_ascii=False, indent=1)
print('-> %s' % os.path.join(OUTDIR, 'cap928.json'))
print(rep)
