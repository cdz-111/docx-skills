# -*- coding: utf-8 -*-
"""校验 中期科技报告9-28-图表修复.docx

被编排调用时（docx-unify-pipeline 的 S2）按三级回退取入参，**不传参时行为与从前一致**：
    SRC = ① sys.argv[1]  ② 环境变量 DOCXFMT_DOC  ③ 原历史默认路径
    OUT = ① sys.argv[2]  ② 环境变量 DOCXFMT_OUT  ③ 原历史默认路径
"""
import os
import zipfile, re, sys, collections
import xml.etree.ElementTree as ET

SRC = (sys.argv[1] if len(sys.argv) > 1
       else os.environ.get('DOCXFMT_DOC')
       or r'C:\Users\14711\Desktop\中期科技报告9-28-公式规范化.docx')
OUT = (sys.argv[2] if len(sys.argv) > 2
       else os.environ.get('DOCXFMT_OUT')
       or r'C:\Users\14711\Desktop\中期科技报告9-28-图表修复.docx')

TAG = re.compile(r'<w:tbl>|</w:tbl>|<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
INSTR = re.compile(r'<w:instrText[^>]*>(.*?)</w:instrText>', re.S)
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.S)

zi = zipfile.ZipFile(OUT)
xml = zi.read('word/document.xml').decode('utf-8')


def spans(x):
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


sp = spans(xml)
kinds = [a for a, b, c in sp]
blocks = [xml[b:c] for a, b, c in sp]


def fields_off(q):
    res, pos = [], 0
    while True:
        b = q.find('<w:fldChar w:fldCharType="begin"/>', pos)
        if b < 0:
            break
        e = q.find('<w:fldChar w:fldCharType="end"/>', b)
        if e < 0:
            break
        body = q[b:e]
        im = INSTR.search(body)
        sep = body.find('<w:fldChar w:fldCharType="separate"/>')
        disp, wts = '', []
        if sep >= 0:
            ds = b + sep + len('<w:fldChar w:fldCharType="separate"/>')
            for w in WT.finditer(q, ds, e):
                disp += w.group(1); wts.append((w.start(1), w.end(1)))
        res.append(dict(instr=INSTR.sub(r'\1', im.group(0)).strip() if im else '',
                        disp=disp, wts=wts,
                        fspan=(b, e + len('<w:fldChar w:fldCharType="end"/>'))))
        pos = e + 30
    return res


def plain_text(q):
    return ''.join(WT.findall(re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S)))


def visible(q):
    out, pos = [], 0
    for f in fields_off(q):
        out.append(plain_text(q[pos:f['fspan'][0]])); out.append(f['disp'])
        pos = f['fspan'][1]
    out.append(plain_text(q[pos:]))
    return ''.join(out)


ok = True
def chk(cond, msg):
    global ok
    print(('  [OK] ' if cond else '  [!!] ') + msg)
    if not cond:
        ok = False


print('=== 1. XML 合法性 ===')
try:
    ET.fromstring(xml)
    chk(True, 'word/document.xml 解析通过 (%d 字符)' % len(xml))
except Exception as e:
    chk(False, 'XML 解析失败: %s' % e)
    sys.exit(1)

print('\n=== 2. 其他部件是否与源文件字节一致 ===')
zs = zipfile.ZipFile(SRC)
diff = []
for it in zs.infolist():
    if it.filename == 'word/document.xml':
        continue
    a = zs.read(it.filename); b = zi.read(it.filename)
    if a != b:
        diff.append(it.filename)
chk(not diff, '除 document.xml 外 %d 个部件全部字节一致%s'
    % (len(zs.infolist()) - 1, ('；差异: %s' % diff) if diff else ''))
chk(zs.namelist() == zi.namelist(), '部件清单一致')

print('\n=== 3. 书签平衡 / 标签嵌套 ===')
chk(len(re.findall(r'<w:t(?: [^>]*)?>[^<]*<w:t[ >/]', xml)) == 0, '无嵌套 <w:t> 结构')
chk(xml.count('<w:t>') + xml.count('<w:t ') == xml.count('</w:t>') + xml.count('<w:t/>'),
    '<w:t> 开闭计数一致')
bs = re.findall(r'<w:bookmarkStart w:id="(\d+)" w:name="([^"]+)"', xml)
be = re.findall(r'<w:bookmarkEnd w:id="(\d+)"', xml)
cs, ce = collections.Counter(i for i, n in bs), collections.Counter(be)
chk(cs == ce, 'bookmarkStart/End 配对一致 (start %d / end %d)' % (len(bs), len(be)))
dupname = [n for n, c in collections.Counter(n for i, n in bs).items() if c > 1]
chk(not dupname, '书签名唯一%s' % ('，重复: %s' % dupname if dupname else ''))

print('\n=== 4. 题注编号连续性 ===')
caps = []
for i, q in enumerate(blocks):
    if kinds[i] != 'P':
        continue
    fs = fields_off(q)
    seq = next((f for f in fs if f['instr'].startswith('SEQ') and
                ('SEQ 图' in f['instr'] or 'SEQ 表' in f['instr'])), None)
    vis = visible(q).strip()
    m0 = re.match(r'^([图表])\s*(\d+)', vis)
    sty = (re.search(r'<w:pStyle w:val="([^"]+)"', q) or [None, '-'])[1]
    if seq is None and not (sty == '12' and m0):
        continue
    bm = re.search(r'<w:bookmarkStart w:id="(\d+)" w:name="(_Ref\w+)"', q)
    caps.append(dict(bi=i, kind=(m0.group(1) if m0 else '?'),
                     num=(m0.group(2) if m0 else None),
                     fmt=('SEQ' if seq else '硬写'),
                     bm=(bm.group(2) if bm else None),
                     title=re.sub(r'^[图表]\s*\d+\s*', '', vis)[:30]))
n图 = sum(1 for c in caps if c['kind'] == '图')
n表 = sum(1 for c in caps if c['kind'] == '表')
chk(n图 == 44 and n表 == 52, '题注数 图%d/表%d (应为 44/52)' % (n图, n表))
bad = [c for c in caps if c['num'] is None]
chk(not bad, '全部题注都带编号%s' % ('，异常: %s' % bad if bad else ''))
for k, tot in (('图', 44), ('表', 52)):
    seq = [int(c['num']) for c in caps if c['kind'] == k]
    chk(seq == list(range(1, tot + 1)), '%s 编号连续 1..%d（实得 %s）'
        % (k, tot, seq if seq != list(range(1, tot + 1)) else 'OK'))
nobm = [c['bi'] for c in caps if not c['bm']]
chk(not nobm, '全部题注都有 _Ref 书签%s' % ('，缺: %s' % nobm if nobm else ''))
chk(len({c['bm'] for c in caps}) == len(caps), '题注书签互不重复')

cap_by_bm = {c['bm']: c for c in caps}

print('\n=== 5. 交叉引用（REF）指向与显示 ===')
nref = 0; err = []
for i, q in enumerate(blocks):
    for f in fields_off(q):
        if not f['instr'].startswith('REF'):
            continue
        nref += 1
        tgt = re.search(r'REF\s+(\S+)', f['instr']).group(1)
        if tgt not in {n for x, n in bs}:
            err.append('block %d 断链 -> %s' % (i, tgt)); continue
        c = cap_by_bm.get(tgt)
        if c is None:
            continue                       # 指向非题注书签（如章节），跳过
        want = '%s %d' % (c['kind'], int(c['num']))
        if f['disp'] != want:
            err.append('block %d %s 显示「%s」应为「%s」' % (i, tgt, f['disp'], want))
chk(nref == 53, 'REF 域数 %d (应为 53)' % nref)
chk(not err, '全部 REF 指向有效且显示与题注一致%s' % ('；问题: %s' % err if err else ''))

print('\n=== 6. 正文硬写号残留 ===')
def vis_map(q):
    s, pm, pos = [], [], 0
    for f in fields_off(q):
        pre = q[pos:f['fspan'][0]]
        for m in WT.finditer(pre):
            for k in range(m.start(1), m.end(1)):
                s.append(q[k]); pm.append(k)
        for ch in f['disp']:
            s.append(ch); pm.append(-1)
        pos = f['fspan'][1]
    tail = q[pos:]
    for m in WT.finditer(tail):
        for k in range(m.start(1), m.end(1)):
            s.append(tail[k]); pm.append(pos + k)
    return ''.join(s), pm

cap_bi = {c['bi']: c for c in caps}
res = []
for i, q in enumerate(blocks):
    if kinds[i] != 'P' or i in cap_bi:
        continue
    s, pm = vis_map(q)
    for m in re.finditer(r'([图表])\s*(\d{1,2})(?!\d)', s):
        ds, de = m.start(2), m.end(2)
        if any(pm[k] < 0 for k in range(ds, de)):
            continue
        near = min(caps, key=lambda c: abs(c['bi'] - i))
        res.append((i, m.group(1) + m.group(2), '%s%s' % (near['kind'], near['num']),
                    abs(near['bi'] - i), visible(q).strip()[:40]))
print('  正文硬写号共 %d 处：' % len(res))
badnum = []
for i, g, near, d, txt in res:
    flag = '' if g == near else '  <-- 与最近题注(%s)不一致' % near
    if flag:
        badnum.append((i, g, near))
    print('    block %-5d %-6s 最近题注=%s(d=%d)%s' % (i, g, near, d, flag))
print('\n  说明：硬写号与「最近题注」不一致并非一定错误（可能是前向引用），此处仅作人工复核清单。')

print('\n=== 7. 公式/图片未受损 ===')
for pat, name in ((r'<w:drawing>', 'drawing'), (r'<m:oMath>', 'oMath'),
                  (r'<m:oMathPara>', 'oMathPara')):
    a = len(re.findall(pat, zs.read('word/document.xml').decode('utf-8')))
    b = len(re.findall(pat, xml))
    chk(a == b, '%s: 源 %d -> 新 %d' % (name, a, b))

print('\n===== 结论: %s =====' % ('全部通过' if ok else '存在问题，见上 [!!]'))
