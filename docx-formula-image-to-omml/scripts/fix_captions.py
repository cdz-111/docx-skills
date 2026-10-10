# -*- coding: utf-8 -*-
"""9-28 图表题注修复 + 交叉引用修复 + 正文硬写号同步。
采用「手术式定点修改」：只改题注里的编号数字/域缓存，不重写整段，
从而保住题注内的图片、_Toc 书签、跨段 bookmarkEnd 等结构。

用法: fix_captions.py                          -> dry run，打印全部映射
      fix_captions.py apply                    -> 落盘到 <out.docx>（旧式）
      fix_captions.py <in.docx> <out.docx> [--apply]   -> 被编排调用时用这个（推荐）

被编排调用时（docx-unify-pipeline 的 S2）**必须**走带 <in.docx> <out.docx> 的写法或用环境变量
`DOCXFMT_DOC` / `DOCXFMT_OUT`，缺失则报错。
"""
import zipfile, re, sys, os, collections

_rest = sys.argv[1:]
APPLY = ('apply' in _rest) or ('--apply' in _rest)
_files = [a for a in _rest if not a.startswith('-') and a != 'apply']

WORK = os.environ.get('DOCXFMT_WORKDIR')
SRC = _files[0] if _files else os.environ.get('DOCXFMT_DOC')
OUT = _files[1] if len(_files) > 1 else os.environ.get('DOCXFMT_OUT')
if not SRC or not OUT:
    sys.exit('用法: fix_captions.py <in.docx> <out.docx> [--apply]'
             '（或设 DOCXFMT_DOC / DOCXFMT_OUT）')

z = zipfile.ZipFile(SRC)
xml = z.read('word/document.xml').decode('utf-8')

TAG = re.compile(r'<w:tbl>|</w:tbl>|<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
INSTR = re.compile(r'<w:instrText[^>]*>(.*?)</w:instrText>', re.S)
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.S)
PH = '\u2423'          # 占位符，用于按「与段落等长」的方式屏蔽域区间


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
    """每个域 -> dict(instr, disp, ins_span, wts, fspan)，坐标相对 q"""
    res, pos = [], 0
    while True:
        b = q.find('<w:fldChar w:fldCharType="begin"/>', pos)
        if b < 0:
            break
        e = q.find('<w:fldChar w:fldCharType="end"/>', b)
        if e < 0:
            break
        body = q[b:e]
        ins_m = INSTR.search(body)
        sep = body.find('<w:fldChar w:fldCharType="separate"/>')
        disp, wts = '', []
        if sep >= 0:
            dstart = b + sep + len('<w:fldChar w:fldCharType="separate"/>')
            for w in WT.finditer(q, dstart, e):
                disp += w.group(1)
                wts.append((w.start(1), w.end(1)))
        res.append(dict(instr=INSTR.sub(r'\1', ins_m.group(0)).strip() if ins_m else '',
                        ins_span=(b + ins_m.start(), b + ins_m.end()) if ins_m else None,
                        disp=disp, wts=wts,
                        fspan=(b, e + len('<w:fldChar w:fldCharType="end"/>'))))
        pos = e + 30
    return res


def plain_text(q):
    return ''.join(WT.findall(re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S)))


def visible(q):
    """完整可见文本（域用缓存值代替）"""
    out, pos = [], 0
    for f in fields_off(q):
        out.append(plain_text(q[pos:f['fspan'][0]]))
        out.append(f['disp'])
        pos = f['fspan'][1]
    out.append(plain_text(q[pos:]))
    return ''.join(out)


def masked(q):
    """与 q 等长，域区间用同长度占位符替换 -> 匹配坐标可与 q 一一对应"""
    out, pos = [], 0
    for f in fields_off(q):
        out.append(q[pos:f['fspan'][0]])
        out.append(PH * (f['fspan'][1] - f['fspan'][0]))
        pos = f['fspan'][1]
    out.append(q[pos:])
    return ''.join(out)


def vis_map(q):
    """返回 (s, pm)：s = 可见文本（域用缓存值填充），pm[k] = s[k] 在 q 中的坐标（-1 表示域内、不可编辑）。
    这样就得到「可见文本 -> 原始 XML」的精确一一映射，且不含任何假空格。"""
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


def to_chunks(pm, ds, de):
    """把 s 上的 [ds,de) 映射为 q 上的连续区间列表；含不可编辑位则返回 None"""
    qs = [pm[k] for k in range(ds, de)]
    if any(x < 0 for x in qs):
        return None
    chunks, cur = [], [qs[0], qs[0] + 1]
    for x in qs[1:]:
        if x == cur[1]:
            cur[1] = x + 1
        else:
            chunks.append((cur[0], cur[1])); cur = [x, x + 1]
    chunks.append((cur[0], cur[1]))
    return chunks


# ================= 1. 题注 =================
caps = []
for i, q in enumerate(blocks):
    if kinds[i] != 'P':
        continue
    fs = fields_off(q)
    seq = next((f for f in fs if f['instr'].startswith('SEQ') and
                ('SEQ 图' in f['instr'] or 'SEQ 表' in f['instr'])), None)
    vis = visible(q).strip()
    m0 = re.match(r'^\s*([图表])\s*(\d+)', vis)
    sty = (re.search(r'<w:pStyle w:val="([^"]+)"', q) or [None, '-'])[1]
    if seq is None and not (sty == '12' and m0):
        continue
    kind = re.search(r'SEQ\s*([图表])', seq['instr']).group(1) if seq else m0.group(1)
    bm = re.search(r'<w:bookmarkStart w:id="(\d+)" w:name="(_Ref\w+)"', q)
    caps.append(dict(bi=i, kind=kind, seq=seq,
                     visnum=(m0.group(2) if m0 else None),
                     seqdisp=(seq['disp'] if seq else None),
                     title=re.sub(r'^[图表]\s*\d+\s*', '', vis),
                     bm=bm.group(2) if bm else None,
                     bid=int(bm.group(1)) if bm else None))

cnt = collections.Counter()
for c in caps:
    cnt[c['kind']] += 1
    c['new'] = cnt[c['kind']]
print('题注 %d 个 -> 图 1-%d / 表 1-%d' % (len(caps), cnt['图'], cnt['表']))

cap_by_bi = {c['bi']: c for c in caps}
cap_by_bm = {c['bm']: c for c in caps if c['bm']}

# 给没有 _Ref 书签的题注分配新书签（原编号一律保留）
used_ids = {int(x) for x in re.findall(r'<w:bookmarkStart w:id="(\d+)"', xml)}
used_names = set(re.findall(r'<w:bookmarkStart[^>]*w:name="([^"]+)"', xml))
nid = (max(used_ids) + 1) if used_ids else 1
for c in caps:
    if c['bm']:
        continue
    while ('_Ref9%04d' % nid) in used_names or nid in used_ids:
        nid += 1
    c['bm'] = '_Ref9%04d' % nid
    c['bid'] = nid
    used_names.add(c['bm']); used_ids.add(nid)
    nid += 1

# ================= 2. 断链 REF 重指向（逐个按段落语义人工核定）=================
REPOINT = {
    ('_Ref2431',  840): 842,   # “…DOMOO…如图X所示” -> 图17 基于嵌套Pareto集学习的离线多目标优化框架
    ('_Ref28658', 856): 857,   # “表X~表Y列出主结果” 4 张 MOO 表之首 -> 表23
    ('_Ref29069', 856): 865,   # 同上，4 张表之末 -> 表26
    ('_Ref2591',  906): 905,   # “图X所示流程…数据筛选—状态感知—本地处理—指导反馈” -> 图19 边缘智能指导流程
    ('_Ref2627',  912): 914,   # “…协同关系如图X所示” -> 图20 边缘—中心协同方法架构
    ('_Ref2627',  915): 914,   # “图X所示方法首先…” -> 图20
    ('_Ref2702',  915): 916,   # “各项边缘计算机制与导学任务之间的关系如表X所示” -> 表28
    ('_Ref2757',  924): 923,   # “边缘—中心协同处理过程如图X所示” -> 图21 分层协同计算范式
    ('_Ref29291', 931): 933,   # “整体技术流程如图X所示” -> 图22 MuEmo-R1
    ('_Ref2855',  959): 960,   # “代表性任务数据来源如表X所示” -> 表31
    ('_Ref29393', 967): 969,   # “图X是强化学习阶段的训练奖励变化情况” -> 图23
    ('_Ref29575', 971): 972,   # “见表X和表Y” -> 表32 DFEW/MAFW
    ('_Ref29598', 971): 974,   # 同上 -> 表33 RAVDESS
    ('_Ref29713', 977): 978,   # “消融实验情况见表X” -> 表34
    ('_Ref29784', 981): 982,   # “表X是…物理特征的对齐度分析” -> 表35
}

print('\n--- REF 域 ---')
ref_log = []
for i, q in enumerate(blocks):
    for f in fields_off(q):
        if not f['instr'].startswith('REF'):
            continue
        tgt = re.search(r'REF\s+(\S+)', f['instr']).group(1)
        nb = REPOINT.get((tgt, i))
        newcap = cap_by_bi.get(nb) if nb is not None else None
        cur = cap_by_bm.get(tgt)
        t = newcap or cur
        if not t:
            print('  !! 未解析 %-14s @block %-5d disp=%r' % (tgt, i, f['disp']))
            continue
        want = '%s %d' % (t['kind'], t['new'])
        if newcap is None and f['disp'] == want:
            continue
        ref_log.append(dict(i=i, f=f, want=want, newbm=t['bm'], oldbm=tgt,
                            disp=f['disp'], mode=('重指向' if newcap else '只改显示')))

# ================= 3. 正文硬写号 =================
HARD_OVERRIDE = {                      # 按段落语义人工核定（作者手写号与题注号存在偏移）
    (788, '表18'): 789,                # Geo3K 表 -> 表17(新29)
    (788, '表19'): 792,                # K12 表   -> 表18(新30)
    (816, '表23'): 812,                # 消融表   -> 表22(新34)
    (903, '图65'): 905,                # 全文无图65 -> 边缘智能指导流程(新37)
}
hard_log = []
for i, q in enumerate(blocks):
    if kinds[i] != 'P' or i in cap_by_bi:
        continue
    s, pm = vis_map(q)
    for m in re.finditer(r'([图表])\s*(\d{1,2})(?!\d)', s):
        ds, de = m.start(2), m.end(2)
        kind, onum = m.group(1), m.group(2)
        chunks = to_chunks(pm, ds, de)
        if chunks is None:
            continue                                  # 编号落在域缓存里，不是硬写文本
        nb = HARD_OVERRIDE.get((i, '%s%s' % (kind, onum)))
        t = cap_by_bi.get(nb) if nb is not None else None
        if t is not None:
            conf = '人工指定'
        else:
            same = [c for c in caps if c['kind'] == kind and c['visnum'] == onum]
            if not same:
                print('  ~ 跳过（题注中无此编号）: block %-5d %s%s' % (i, kind, onum))
                continue
            t = min(same, key=lambda c: abs(c['bi'] - i))
            d = abs(t['bi'] - i)
            conf = '旧号唯一' if len(same) == 1 else ('就近' if d <= 12 else '!!远(%d)' % d)
        hard_log.append(dict(i=i, chunks=chunks, old='%s%s' % (kind, onum),
                             t=t, conf=conf))

print('\n--- 正文硬写号 (%d 处) ---' % len(hard_log))
for h in hard_log:
    t = h['t']
    print('  block %-5d %-7s -> %s%-3d [%s]  %s'
          % (h['i'], h['old'], t['kind'], t['new'], h['conf'], t['title'][:34]))

print('\n--- REF 改写明细 (%d 处) ---' % len(ref_log))
for r in ref_log:
    print('  block %-5d %-14s->%-14s 「%s」->「%s」 [%s]'
          % (r['i'], r['oldbm'], r['newbm'], r['disp'], r['want'], r['mode']))

if len(sys.argv) > 1 and sys.argv[1] == 'dump':
    import json
    json.dump(dict(
        caps=[dict(bi=c['bi'], kind=c['kind'], old=c['visnum'], new=c['new'],
                   seq=bool(c['seq']), seqdisp=c['seqdisp'], bm=c['bm'],
                   title=c['title']) for c in caps],
        refs=[dict(i=r['i'], old=r['oldbm'], new=r['newbm'], disp=r['disp'],
                   want=r['want'], mode=r['mode']) for r in ref_log],
        hards=[dict(i=h['i'], old=h['old'], new='%s%d' % (h['t']['kind'], h['t']['new']),
                    conf=h['conf'], near_bi=h['t']['bi'], near_title=h['t']['title'])
               for h in hard_log],
    ), open(os.path.join(WORK, 'fix928_dump.json'), 'w', encoding='utf-8'),
        ensure_ascii=False, indent=1)
    print('dump 已写出 fix928_dump.json')
    sys.exit(0)

if not APPLY:
    print('\n--- 题注新旧号对照（%d 个）---' % len(caps))
    for c in caps:
        print('  block %-5d %s%-4s -> %-4s %-12s %s  %s'
              % (c['bi'], c['kind'], c['visnum'], '%s%d' % (c['kind'], c['new']),
                 c['bm'], 'SEQ' if c['seq'] else '硬写', c['title'][:38]))
    print('\n(dry run，未写文件。加 apply 参数落盘)')
    sys.exit(0)

# ================= 应用 =================
edits = []
cap_edit_n = 0

for c in caps:
    i = c['bi']; q = blocks[i]; off = sp[i][1]
    N = c['new']
    if c['seq']:
        # a) 域显示数字 -> N（wts 是 <w:t> 的「内容」区间，只替内容，不能塞标签）
        if c['seq']['wts']:
            a, b = c['seq']['wts'][0]
            edits.append((off + a, off + b, str(N)))
            for a2, b2 in c['seq']['wts'][1:]:
                edits.append((off + a2, off + b2, ''))
        else:                                   # 无显示区：在 instrText 后补 separate+显示
            a = c['seq']['ins_span'][1]
            edits.append((off + a, off + a,
                          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                          '<w:r><w:t>%d</w:t></w:r>' % N))
        # b/c) 域后第一个 <w:t>：剥离粘连数字；编号与标题之间补空格
        e = c['seq']['fspan'][1]
        mo = re.search(r'<w:t( [^>]*)?>', q[e:])
        if mo:
            cstart = e + mo.end()
            cont = q[cstart:q.find('<', cstart)]
            glue_done = False
            mg = re.match(r'(\d+)', cont)
            if mg and c['visnum'] == (c['seqdisp'] or '') + mg.group(1):
                n = len(mg.group(1))
                edits.append((off + cstart, off + cstart + n,
                              '' if cont[n:n + 1] == ' ' else ' '))
                cap_edit_n += 1
                glue_done = True
            if not glue_done and not cont.startswith(' '):
                if 'xml:space' not in mo.group(0):
                    edits.append((off + e + mo.start(), off + e + mo.end(),
                                  '<w:t xml:space="preserve">'))
                edits.append((off + cstart, off + cstart, ' '))
                cap_edit_n += 1
    else:
        # 硬写题注：把可见数字替换为 N；若「图/表」与编号在 XML 里紧挨着，一并补一个空格
        s, pm = vis_map(q)
        done = False
        for m in re.finditer(r'([图表])\s*(\d{1,2})(?!\d)', s):
            ds, de = m.start(2), m.end(2)
            if m.group(2) != c['visnum']:
                continue
            if any(pm[k] < 0 for k in range(ds, de)):
                continue
            k0, k1 = m.start(1), m.end(1)
            if pm[k0] >= 0 and pm[k1 - 1] + 1 == pm[ds]:
                edits.append((off + pm[k0], off + pm[de - 1] + 1, '%s %d' % (c['kind'], N)))
            else:
                for k, (x, y) in enumerate(to_chunks(pm, ds, de)):
                    edits.append((off + x, off + y, str(N) if k == 0 else ''))
            done = True
            break
        assert done, '硬写题注未定位数字: block %d' % i
    # c) 补 _Ref 书签（原本没有的）
    if c['bid'] is not None:
        had = bool(re.search(r'<w:bookmarkStart w:id="%d" w:name="%s"' % (c['bid'], c['bm']), q))
        if not had:
            mp = re.search(r'<w:pPr>.*?</w:pPr>|<w:pPr/>', q, re.S)
            ins = mp.end() if mp else re.match(r'<w:p[^>]*>', q).end()
            edits.append((off + ins, off + ins,
                          '<w:bookmarkStart w:id="%d" w:name="%s"/>' % (c['bid'], c['bm'])))
            ie = q.rfind('</w:p>')
            edits.append((off + ie, off + ie, '<w:bookmarkEnd w:id="%d"/>' % c['bid']))

for r in ref_log:
    i = r['i']; f = r['f']
    if r['mode'] == '重指向':
        a, b = f['ins_span']
        edits.append((sp[i][1] + a, sp[i][1] + b,
                      '<w:instrText xml:space="preserve"> REF %s \\h </w:instrText>' % r['newbm']))
    if not f['wts']:
        continue
    if f['disp'] != r['want']:
        a0, b0 = f['wts'][0]
        edits.append((sp[i][1] + a0, sp[i][1] + b0, r['want']))
        for a, b in f['wts'][1:]:
            edits.append((sp[i][1] + a, sp[i][1] + b, ''))

for h in hard_log:
    off = sp[h['i']][1]
    for k, (a, b) in enumerate(h['chunks']):
        edits.append((off + a, off + b, str(h['t']['new']) if k == 0 else ''))

edits.sort(key=lambda e: (e[0], e[1]), reverse=True)
prev = len(xml) + 1
for a, b, t in edits:
    assert b <= prev, '区间重叠 %d-%d' % (a, b)
    xml = xml[:a] + t + xml[b:]
    prev = a
print('\n应用编辑 %d 处（题注 %d 段 + 粘连数字 %d + 引用 %d + 硬写 %d）'
      % (len(edits), len(caps), cap_edit_n, len(ref_log), len(hard_log)))

zo = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
for it in z.infolist():
    zo.writestr(it, xml.encode('utf-8') if it.filename == 'word/document.xml'
                else z.read(it.filename))
zo.close()
print('已写出 %s (%.1f MB)' % (OUT, os.path.getsize(OUT) / 1048576))
