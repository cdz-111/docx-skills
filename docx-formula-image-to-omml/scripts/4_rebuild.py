# -*- coding: utf-8 -*-
"""步骤4：重建 document.xml。
用法: python 4_rebuild.py <src.docx> <dst.docx> <workdir> [section|global]
依赖: queue.json / disp_list.json / inline_list.json / omml_display.json / omml_inline.json
做的事:
  1) 展示公式段（行间公式）-> 居中制表位 + 原生 OMML + 右对齐编号 (n)
     编号模式 section = 每小节重新从 (1) 起；global = 全文从 (1) 连续排到 (N)；默认 section
  2) 行内公式图 -> 裸 <m:oMath>（行内公式**不编号**）
  3) 正文 Unicode 伪角标 -> 真正的 vertAlign 上下标
  4) 全文重新编号时，同步改写正文里的「公式（n）」引用（按引用所在小节做局部->全局映射）
制表位按每个公式所在节自动计算（文档常有多个 sectPr，含横向页）。

HEAD_STYLES 是本文档里标题所用的 pStyle 编号，换文档时可能要改；
它只用于「把某段归属到哪个小节」，从而决定编号与正文引用的映射。
"""
import zipfile, re, json, os, sys, collections

SRC, DST, W = sys.argv[1], sys.argv[2], sys.argv[3]
MODE = sys.argv[4] if len(sys.argv) > 4 else 'section'
HEAD_STYLES = {'2', '3', '4'}

PARA_RE = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
DRW_RE = re.compile(r'<w:drawing>.*?</w:drawing>', re.S)
TOK_RE = re.compile(r'<w:drawing>.*?</w:drawing>|<w:t(?: [^>]*)?>.*?</w:t>|<w:tab/>|<w:br/>', re.S)

SUB_MAP = {'₀':'0','₁':'1','₂':'2','₃':'3','₄':'4','₅':'5','₆':'6','₇':'7','₈':'8','₉':'9',
           '₊':'+','₋':'-','₌':'=','₍':'(','₎':')','ₐ':'a','ₑ':'e','ₒ':'o','ₓ':'x','ₕ':'h',
           'ₖ':'k','ₗ':'l','ₘ':'m','ₙ':'n','ₚ':'p','ₛ':'s','ₜ':'t','ᵢ':'i','ᵣ':'r','ᵤ':'u',
           'ᵥ':'v','ⱼ':'j'}
SUP_MAP = {'⁰':'0','¹':'1','²':'2','³':'3','⁴':'4','⁵':'5','⁶':'6','⁷':'7','⁸':'8','⁹':'9',
           'ⁱ':'i','ⁿ':'n','ᵀ':'T','ᵃ':'a','ᵇ':'b','ᶜ':'c','ᵈ':'d','ᵉ':'e','ᶠ':'f','ᵍ':'g',
           'ʰ':'h','ʲ':'j','ᵏ':'k','ˡ':'l','ᵐ':'m','ᵒ':'o','ᵖ':'p','ʳ':'r','ˢ':'s','ᵗ':'t',
           'ᵘ':'u','ᵛ':'v','ʷ':'w','ˣ':'x','ʸ':'y','ᶻ':'z'}

z = zipfile.ZipFile(SRC)
xml = z.read('word/document.xml').decode('utf-8')
rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))

disp_omml = json.load(open(os.path.join(W, 'omml_display.json'), encoding='utf-8'))
inl_omml = json.load(open(os.path.join(W, 'omml_inline.json'), encoding='utf-8'))
disp_list = json.load(open(os.path.join(W, 'disp_list.json'), encoding='utf-8'))
inline_list = json.load(open(os.path.join(W, 'inline_list.json'), encoding='utf-8'))

disp_n, disp_sec = {}, {}
for d in disp_list:
    disp_n[d['file']] = len([k for k in disp_omml if k.rsplit('#', 1)[0] == d['file']])
    disp_sec[d['file']] = d['sec']
inline_files = {u['file'] for u in inline_list}

paras_all = PARA_RE.findall(xml)

# ---- 段 -> 所属小节；小节 -> 展示公式数；小节内序号 -> 全文序号 ----
sec_of, cur = {}, ''
for i, p in enumerate(paras_all):
    st = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.S))
    if st and st.group(1) in HEAD_STYLES and txt.strip():
        cur = txt.split()[0]
    sec_of[i] = cur

sec_order, sec_count = [], {}
for d in disp_list:                       # 已是文档顺序
    s = d['sec']
    if s not in sec_count:
        sec_count[s] = 0
        sec_order.append(s)
    sec_count[s] += disp_n[d['file']]     # 按「公式数」累计（一张图里可能有多个公式）
L2G, _off = {}, 0
for s in sec_order:
    L2G[s] = {k: _off + k for k in range(1, sec_count[s] + 1)}
    _off += sec_count[s]
ref_stat = collections.Counter()

# 正文「公式（n）」引用：n 是「所在小节内的局部序号」，映射到全文序号
REF_RE = re.compile(r'公式\s*（(\d+)）')
# 手工核对过的特殊引用（标签里的章节号与所在小节不一致，经语义核对后显式指定）
REF_OVERRIDE = [
    # 2.2.2 正文写「式（2.6-6）」，实际指该小节第 6 个公式 Ltotal（原标签章节号有误）
    ('2.2.2', '式（2.6-6）', '公式（%d）' % L2G.get('2.2.2', {}).get(6, 6)),
]


def compute_edits(text, sec):
    """返回 [(start, end, new_text), ...]；无改动则为空列表。"""
    edits = []
    if sec not in L2G:
        return edits
    m1 = L2G[sec]
    for m in REF_RE.finditer(text):
        n, g = int(m.group(1)), m1.get(int(m.group(1)))
        if g and g != n:
            edits.append((m.start(), m.end(), '公式（%d）' % g))
    for s0, old, new in REF_OVERRIDE:
        if s0 == sec and old in text:
            p = 0
            while True:
                j = text.find(old, p)
                if j < 0:
                    break
                edits.append((j, j + len(old), new))
                p = j + len(old)
    # 去掉互相重叠的改写（保留靠前者）
    edits.sort()
    keep, last_end = [], -1
    for e in edits:
        if e[0] >= last_end:
            keep.append(e); last_end = e[1]
    ref_stat['ref_edits'] += len(keep)
    return keep


T_RE = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')


def rewrite_para_refs(para, sec):
    """跨 run 改写正文引用：把各 <w:t> 文本拼起来匹配，再按 run 切片回填。

    只在改动区间覆盖到的 run 上重建文本，其余 run 原样保留（保格式）。
    """
    if sec not in L2G:
        return para
    ms = list(T_RE.finditer(para))
    if not ms:
        return para
    texts = [m.group(1) for m in ms]
    full = ''.join(texts)
    if '公式' not in full and '式（' not in full:
        return para
    edits = compute_edits(full, sec)
    if not edits:
        return para

    starts, cur = [], 0
    for t in texts:
        starts.append(cur); cur += len(t)

    new_texts = []
    for i, t in enumerate(texts):
        s0, e0 = starts[i], starts[i] + len(t)
        buf, pos = [], s0
        for es, ee, new in edits:
            if ee <= s0 or es >= e0:
                continue
            if es > pos:
                buf.append(full[pos:es])
            if s0 <= es < e0:
                buf.append(new)
            pos = max(pos, ee)
        if pos < e0:
            buf.append(full[pos:e0])
        new_texts.append(''.join(buf))

    out, last = [], 0
    for m, nt in zip(ms, new_texts):
        out.append(para[last:m.start(1)]); out.append(nt)
        last = m.end(1)
    out.append(para[last:])
    return ''.join(out)

# ---- 按节算制表位 ----
sec_stops, end_idx = [], []
for i, p in enumerate(paras_all):
    m = re.search(r'<w:sectPr.*?</w:sectPr>', p, re.S)
    if m:
        s = m.group(0)
        pg = re.search(r'<w:pgSz[^>]*w:w="(\d+)"', s)
        ml = re.search(r'<w:pgMar[^>]*w:left="(\d+)"', s)
        mr = re.search(r'<w:pgMar[^>]*w:right="(\d+)"', s)
        w = int(pg.group(1)) if pg else 11906
        l = int(ml.group(1)) if ml else 1800
        r = int(mr.group(1)) if mr else 1800
        tw = w - l - r
        sec_stops.append((tw, tw // 2))
        end_idx.append(i)
    # 记录 index 位置
    pass
# 段索引 -> 用 finditer 拿真实下标，这里用顺序对齐即可
idx_of = list(range(len(paras_all)))


def stops_for(pi):
    for k, e in enumerate(end_idx):
        if pi <= e:
            return sec_stops[k]
    return sec_stops[-1] if sec_stops else (8306, 4153)


stats = collections.Counter()
num_counter, nums, num_log = {}, {}, []
gnum = [0]                       # 全文连续编号计数器


def media_of(draw):
    m = re.search(r'r:embed="(rId\d+)"', draw)
    if not m:
        return None
    t = relmap.get(m.group(1), '')
    return os.path.basename(t) if t else None


def split_script(text):
    out, cur, curk = [], '', 'n'
    for c in text:
        k = 'sub' if c in SUB_MAP else ('sup' if c in SUP_MAP else 'n')
        ch = SUB_MAP.get(c) or SUP_MAP.get(c) or c
        if k != curk and cur:
            out.append((curk, cur)); cur = ''
        curk = k
        cur += ch
    if cur:
        out.append((curk, cur))
    return out


def mk_run(text, rpr='', align=None):
    r = rpr
    if align:
        if r and '</w:rPr>' in r:
            r = r.replace('</w:rPr>', '<w:vertAlign w:val="%s"/></w:rPr>' % align)
        elif r:
            r += '<w:vertAlign w:val="%s"/>' % align
        else:
            r = '<w:rPr><w:vertAlign w:val="%s"/></w:rPr>' % align
    return '<w:r>' + r + '<w:t xml:space="preserve">' + text + '</w:t></w:r>'


def has_script(t):
    return any(c in SUB_MAP or c in SUP_MAP for c in t)


def rebuild_run(run):
    om = re.match(r'<w:r(?: [^>]*)?>', run)
    if not om:
        return run
    inner = run[len(om.group(0)):-len('</w:r>')]
    if '<w:drawing>' not in inner and not has_script(re.sub(r'<[^>]*>', '', inner)):
        return run
    rpm = re.search(r'<w:rPr>.*?</w:rPr>', inner, re.S)
    rpr = rpm.group(0) if rpm else ''
    out = []
    for t in TOK_RE.finditer(inner):
        tok = t.group(0)
        if tok.startswith('<w:drawing'):
            f = media_of(tok)
            if f in inline_files:
                out.append(inl_omml[f]); stats['inline'] += 1
            else:
                out.append('<w:r>' + tok + '</w:r>')
        elif tok.startswith('<w:t'):
            body = tok[tok.find('>') + 1:-len('</w:t>')]     # 已是 XML 转义文本，勿再 escape
            if has_script(body):
                for kind, chunk in split_script(body):
                    if kind == 'n':
                        out.append(mk_run(chunk, rpr))
                    else:
                        out.append(mk_run(chunk, rpr, 'subscript' if kind == 'sub' else 'superscript'))
                        stats[kind] += 1
            else:
                out.append(mk_run(body, rpr))
        else:
            out.append('<w:r>' + rpr + tok + '</w:r>')
    return ''.join(out)


def insert_tabs(pPr, tw, half):
    # pPr 内元素顺序有 schema 要求：tabs 必须在 spacing/ind/jc 之前
    pPr = re.sub(r'<w:jc [^>]*/>', '', pPr)
    pPr = re.sub(r'<w:tabs>.*?</w:tabs>', '', pPr, flags=re.S)
    tabs = ('<w:tabs><w:tab w:val="center" w:pos="%d"/>'
            '<w:tab w:val="right" w:pos="%d"/></w:tabs>' % (half, tw))
    anchors = ['<w:spacing', '<w:ind', '<w:jc', '<w:rPr', '<w:sectPr']
    cands = [pPr.find(a) for a in anchors if pPr.find(a) >= 0]
    idx = min(cands + [len(pPr) - len('</w:pPr>')])
    pPr = pPr[:idx] + tabs + pPr[idx:]
    j = pPr.find('<w:rPr')
    if j < 0:
        j = len(pPr) - len('</w:pPr>')
    return pPr[:j] + '<w:jc w:val="left"/>' + pPr[j:]


def build_display(base, om, num, tw, half):
    pPr = insert_tabs(base, tw, half)
    return ('<w:p>' + pPr + '<w:r><w:tab/></w:r>' + om +
            '<w:r><w:tab/><w:t>' + num + '</w:t></w:r></w:p>')


CORPUS = r'<w:r(?: [^>]*)?>.*?</w:r>'


def process(pi, para):
    sec = sec_of.get(pi, '')
    if '<w:drawing>' in para:
        files = [media_of(d) for d in DRW_RE.findall(para)]
        has_text = any(t.strip() for t in re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para))
        if not has_text and files and all(f in disp_n for f in files):
            pm = re.search(r'<w:pPr>.*?</w:pPr>', para, re.S)
            base = pm.group(0) if pm else '<w:pPr><w:widowControl/><w:rPr><w:lang w:eastAsia="zh"/></w:rPr></w:pPr>'
            tw, half = stops_for(pi)
            out = []
            for f in files:
                for i in range(disp_n[f]):
                    if MODE == 'global':
                        gnum[0] += 1
                        lab = '(%d)' % gnum[0]
                        num_counter[disp_sec[f]] = num_counter.get(disp_sec[f], 0)
                    else:
                        n = num_counter.get(disp_sec[f], 0) + 1
                        num_counter[disp_sec[f]] = n
                        lab = '(%d)' % n
                    nums['%s#%d' % (f, i)] = lab
                    num_log.append((f, i, disp_sec[f], lab))
                    out.append(build_display(base, disp_omml['%s#%d' % (f, i)], lab, tw, half))
                    stats['display_eq'] += 1
            stats['display_para'] += 1
            return ''.join(out)
    para = re.sub(CORPUS, lambda m: rebuild_run(m.group(0)), para, flags=re.S)
    if MODE == 'global':
        para = rewrite_para_refs(para, sec)      # 行内公式只换符号，不编号
    return para


pieces, last = [], 0
pi = -1
for m in PARA_RE.finditer(xml):
    pi += 1
    pieces.append(xml[last:m.start()])
    pieces.append(process(pi, m.group(0)))
    last = m.end()
pieces.append(xml[last:])
new_xml = ''.join(pieces)

print('mode:', MODE)
print('stats:', dict(stats))
if MODE == 'global':
    print('global numbering: 1 .. %d' % gnum[0])
else:
    print('numbering per section:', num_counter)
print('ref rewrite:', dict(ref_stat))

if os.path.exists(DST):
    os.remove(DST)
zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'word/document.xml':
        data = new_xml.encode('utf-8')
    zout.writestr(item, data)          # 保留原 ZipInfo，媒体字节级不变
zout.close()
print('written:', DST, os.path.getsize(DST) // 1024, 'KB')

json.dump(nums, open(os.path.join(W, 'numbers.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
json.dump(num_log, open(os.path.join(W, 'number_log.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
