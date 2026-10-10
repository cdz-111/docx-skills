# -*- coding: utf-8 -*-
"""对 9-26 版报告做公式规范化：
  1) 所有行间公式统一成「居中制表位 + 公式 + 右对齐制表位 + 编号」
     - A 型（已有制表位）→ 重建，补/改编号
     - B 型（Word 原生 oMathPara 显示公式）→ 剥掉 oMathPara 包裹，套同一版式
  2) 全文从 (1) 连续编号
  3) 同步正文「公式（n）」引用（旧编号 -> 新编号）
用法: python 6_rebuild_926.py <src.docx> <dst.docx> <workdir>
"""
import zipfile, re, os, sys, json, shutil, datetime, collections

SRC, DST, W = sys.argv[1], sys.argv[2], sys.argv[3]

PARA_RE = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
HEAD_STYLES = {'2', '3', '4'}
LABEL_RE = re.compile(r'\((\d+)\)')
T_RE = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')
REF_RE = re.compile(r'公式\s*（(\d+)）')

ts = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
base, ext = os.path.splitext(SRC)
import glob
_exist = sorted(glob.glob('%s.backup-*%s' % (base, ext)))
if _exist:
    bak = _exist[0]
    print('备份已存在，复用:', bak)
else:
    bak = '%s.backup-%s%s' % (base, ts, ext)
    shutil.copy2(SRC, bak)
    print('备份:', bak)

xml = zipfile.ZipFile(SRC).read('word/document.xml').decode('utf-8')
paras_all = PARA_RE.findall(xml)

# ---------- 段落 -> 小节 ----------
sec_of, cur = {}, ''
for i, p in enumerate(paras_all):
    st = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p, re.S))
    if st and st.group(1) in HEAD_STYLES and txt.strip():
        cur = txt.strip().split()[0]
    sec_of[i] = cur

# ---------- 制表位（按 sectPr） ----------
# 语义：段落级 sectPr 是「该节的最后一个段落」；其后的内容属于下一节。
# 最后一节的页属性来自 body 级 sectPr（不在任何 w:p 里），必须单独补上。
def stops_of(s):
    pg = re.search(r'<w:pgSz[^>]*w:w="(\d+)"', s)
    ml = re.search(r'<w:pgMar[^>]*w:left="(\d+)"', s)
    mr = re.search(r'<w:pgMar[^>]*w:right="(\d+)"', s)
    w = int(pg.group(1)) if pg else 11906
    l = int(ml.group(1)) if ml else 1800
    r = int(mr.group(1)) if mr else 1800
    return (w - l - r, (w - l - r) // 2)


sec_stops, end_idx = [], []
for i, p in enumerate(paras_all):
    m = re.search(r'<w:sectPr.*?</w:sectPr>', p, re.S)
    if m:
        sec_stops.append(stops_of(m.group(0)))
        end_idx.append(i)

# body 级 sectPr（最后一个 </w:p> 之后的那个）覆盖收尾部分
_tail = xml[xml.rfind('</w:p>'):] if '</w:p>' in xml else ''
_m = re.search(r'<w:sectPr.*?</w:sectPr>', _tail, re.S)
body_stops = stops_of(_m.group(0)) if _m else (8306, 4153)
print('节数=%d  body 级制表位=%s' % (len(sec_stops), body_stops))


def stops_for(pi):
    for k, e in enumerate(end_idx):
        if pi <= e:
            return sec_stops[k]
    return body_stops


def insert_tabs(ppr, tw, half):
    ppr = re.sub(r'<w:jc [^>]*/>', '', ppr)
    ppr = re.sub(r'<w:tabs>.*?</w:tabs>', '', ppr, flags=re.S)
    tabs = ('<w:tabs><w:tab w:val="center" w:pos="%d"/>'
            '<w:tab w:val="right" w:pos="%d"/></w:tabs>' % (half, tw))
    anchors = ['<w:spacing', '<w:ind', '<w:jc', '<w:rPr', '<w:sectPr']
    cands = [ppr.find(a) for a in anchors if ppr.find(a) >= 0]
    idx = min(cands + [len(ppr) - len('</w:pPr>')])
    ppr = ppr[:idx] + tabs + ppr[idx:]
    j = ppr.find('<w:rPr')
    if j < 0:
        j = len(ppr) - len('</w:pPr>')
    return ppr[:j] + '<w:jc w:val="left"/>' + ppr[j:]


def is_display(p):
    if '<w:drawing>' in p:
        return False
    pm = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
    ppr = pm.group(0) if pm else ''
    if '<m:oMathPara>' in p:
        return True
    if not re.search(r'<w:tab w:val="center" w:pos="4153"/>', ppr):
        return False
    if '<m:oMath>' not in p:
        return False
    vt = ''.join(T_RE.findall(p))
    return re.sub(r'\(\d+\)', '', vt).strip() == ''


# ================= 第 1 遍：编排新编号 + 建引用映射 =================
disp_idx = [i for i, p in enumerate(paras_all) if is_display(p)]
NEWNUM, OLD2NEW, rows = {}, {}, []
gnum = 0
for i in disp_idx:
    p = paras_all[i]
    gnum += 1
    NEWNUM[i] = gnum
    vt = ''.join(T_RE.findall(p))
    old = LABEL_RE.search(vt)
    old_lab = int(old.group(1)) if old else None
    if old_lab:
        OLD2NEW[old_lab] = gnum
    inner = ''.join(re.findall(r'<m:oMath>.*?</m:oMath>', p, re.S))
    rows.append({'para': i, 'sec': sec_of.get(i, ''),
                 'type': 'B(oMathPara)' if '<m:oMathPara>' in p else 'A(制表位)',
                 'old': old_lab, 'new': gnum,
                 'math': ''.join(re.findall(r'<m:t[^>]*>([^<]*)</m:t>', inner))[:70]})

print('行间公式总数 = %d' % gnum)
print('  原先无编号 = %d' % sum(1 for r in rows if r['old'] is None))
print('  A 型 %d / B 型 %d' % (
    sum(1 for r in rows if r['type'][0] == 'A'),
    sum(1 for r in rows if r['type'][0] == 'B')))
print('旧编号 -> 新编号（有变化的）：',
      {k: v for k, v in sorted(OLD2NEW.items()) if k != v})

# ================= 第 2 遍：重建 XML =================
stats = collections.Counter()
ref_log = []
pieces, last, pi = [], 0, -1
for m in PARA_RE.finditer(xml):
    pi += 1
    p = m.group(0)
    pieces.append(xml[last:m.start()])
    if pi in NEWNUM:
        n = NEWNUM[pi]
        pm = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
        ppr = pm.group(0) if pm else ('<w:pPr><w:widowControl/>'
                                     '<w:rPr><w:lang w:eastAsia="zh"/></w:rPr></w:pPr>')
        inner = ''.join(re.findall(r'<m:oMath>.*?</m:oMath>', p, re.S))
        am = re.match(r'<w:p( [^>]*)?>', p)
        attrs = (am.group(1) or '') if am else ''
        tw, half = stops_for(pi)
        p = ('<w:p%s>' % attrs) + insert_tabs(ppr, tw, half) + \
            '<w:r><w:tab/></w:r>' + inner + \
            '<w:r><w:tab/><w:t>(%d)</w:t></w:r>' % n + '</w:p>'
        stats['display'] += 1
    else:
        ms = list(T_RE.finditer(p))
        if ms:
            texts = [x.group(1) for x in ms]
            full = ''.join(texts)
            if '公式' in full:
                edits = []
                for mm in REF_RE.finditer(full):
                    o = int(mm.group(1))
                    nn = OLD2NEW.get(o)
                    if nn and nn != o:
                        edits.append((mm.start(), mm.end(), '公式（%d）' % nn))
                        ref_log.append((pi, o, nn))
                if edits:
                    starts, c = [], 0
                    for t in texts:
                        starts.append(c); c += len(t)
                    new_texts = []
                    for k, t in enumerate(texts):
                        s0, e0 = starts[k], starts[k] + len(t)
                        buf, pos = [], s0
                        for es, ee, nw in edits:
                            if ee <= s0 or es >= e0:
                                continue
                            if es > pos:
                                buf.append(full[pos:es])
                            if s0 <= es < e0:
                                buf.append(nw)
                            pos = max(pos, ee)
                        if pos < e0:
                            buf.append(full[pos:e0])
                        new_texts.append(''.join(buf))
                    out, ll = [], 0
                    for x, nt in zip(ms, new_texts):
                        out.append(p[ll:x.start(1)]); out.append(nt); ll = x.end(1)
                    out.append(p[ll:])
                    p = ''.join(out)
                    stats['refs'] += len(edits)
    pieces.append(p)
    last = m.end()
pieces.append(xml[last:])
new_xml = ''.join(pieces)

if os.path.exists(DST):
    os.remove(DST)
zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'word/document.xml':
        data = new_xml.encode('utf-8')
    zout.writestr(item, data)
zout.close()

print('正文引用改写: %s' % ref_log)
print('输出:', DST, os.path.getsize(DST) // 1024, 'KB')
json.dump(rows, open(os.path.join(W, 'renumber_926.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
