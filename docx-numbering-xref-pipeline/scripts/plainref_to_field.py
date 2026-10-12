# -*- coding: utf-8 -*-
"""
plainref_to_field.py —— 把正文里「硬写」的图/表引用（纯文本，如 图2-7 / 图 5-5）
转成正规交叉引用域 REF _RefXXX \\h；并按需给目标题注补 _Ref 书签（包编号区）。

用法:
    python plainref_to_field.py <in.docx> <out.docx> <plan.json>

plan.json:
{
  "check_md5": "<in.docx 的 md5>",
  "expected_hits": 20,
  "binds": [{"caption": "图 5-5", "ref": "_Ref4153", "bookmark_id": 4153}],
  "map":   {"图 5-5": "_Ref4153", "图 2-7": "_Ref1857208920"}
}

不变量：w:p / w:tbl / w:drawing 计数不变；只有被命中 run 会被拆分，书签只增不改。
任何断言失败 → 不落盘、退出码 2。
"""
import re, sys, json, hashlib, zipfile

def die(msg, code=2):
    print('[FAIL] ' + msg); sys.exit(code)

def escapes(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

if len(sys.argv) < 4:
    print(__doc__); sys.exit(6)
IN, OUT, PLAN = sys.argv[1], sys.argv[2], sys.argv[3]
plan = json.load(open(PLAN, encoding='utf-8'))

raw = open(IN, 'rb').read()
got = hashlib.md5(raw).hexdigest()
if got != plan['check_md5']:
    die('源文件 md5 不符：得到 %s，期望 %s' % (got, plan['check_md5']))

zin = zipfile.ZipFile(IN)
x = zin.read('word/document.xml').decode('utf-8')
N0 = len(x)

# ---------- 基础：域栈 / 段落 / run ----------
FC = re.compile(r'<w:fldChar w:fldCharType="(begin|separate|end)"\s*/>')
stack, fields = [], []
for m in FC.finditer(x):
    k = m.group(1)
    if k == 'begin': stack.append([m.start(), None])
    elif k == 'separate':
        if stack and stack[-1][1] is None: stack[-1][1] = m.start()
    else:
        if stack:
            b, s = stack.pop(); fields.append((b, s, m.end()))
res_regions = sorted([(s, e) for (b, s, e) in fields if s is not None])
def in_res(pos):
    for a, b in res_regions:
        if a <= pos < b: return True
        if a > pos: return False
    return False

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
def paras(s):
    r, d, st = [], 0, None
    for m in TAGP.finditer(s):
        t = m.group(0)
        if t.endswith('/>'):
            if d == 0: r.append((m.start(), m.end()))
        elif t == '</w:p>':
            d -= 1
            if d <= 0:
                if st is not None: r.append((st, m.end()))
                st = None; d = 0
        else:
            if d == 0: st = m.start()
            d += 1
    return r
P = paras(x)
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')
INSTR = re.compile(r'<w:instrText[^>]*>([^<]*)</w:instrText>')
REFRE = re.compile(r'([图表])\s*(\d+)\s*[-\u2013\u2014\uff0d]\s*(\d+)')

def run_open_at(pos):
    s = x.rfind('<w:r', 0, pos)
    while s != -1 and x[s+4:s+5] not in ('>', ' '):
        s = x.rfind('<w:r', 0, s)
    return s

edits = []   # (start, end, replacement)

# ---------- STEP B: 正文硬写 → REF 域 ----------
runs = {}     # rs -> {re, open, rpr, text, text_off, matches:[(s,e,ref,disp)]}
hits = 0
for i, (a, b) in enumerate(P):
    q = x[a:b]
    instrs = [t.strip() for t in INSTR.findall(q)]
    if any(t.startswith('SEQ 图') or t.startswith('SEQ 表') for t in instrs): continue
    if any(t.startswith('TOC ') or t.startswith('PAGEREF') for t in instrs): continue
    for m in WT.finditer(q):
        gpos = a + m.start()
        if in_res(gpos): continue
        txt = m.group(1)
        for mm in REFRE.finditer(txt):
            canon = '%s %s-%s' % (mm.group(1), mm.group(2), mm.group(3))
            if canon not in plan['map']:
                die('命中 %r（段p%d）在 map 中无目标，中止' % (canon, i))
            rs = run_open_at(gpos)
            if rs == -1: die('无法定位 run（段p%d）' % i)
            re_ = x.find('</w:r>', gpos) + 6
            key = rs
            if key not in runs:
                run_xml = x[rs:re_]
                mo = re.match(r'^(<w:r(?: [^>]*)?>)', run_xml)
                opentag = mo.group(1); rest = run_xml[len(opentag):]
                mr = re.match(r'^(<w:rPr>.*?</w:rPr>)', rest, re.S)
                rpr = mr.group(1) if mr else ''
                rest2 = rest[len(rpr):]
                mt = re.match(r'^<w:t(?: [^>]*)?>([^<]*)</w:t></w:r>$', rest2)
                if not mt:
                    die('run 结构非预期（段p%d）：%s' % (i, run_xml[:120]))
                runs[key] = {'re': re_, 'open': opentag, 'rpr': rpr,
                             'text': mt.group(1), 'matches': []}
            # 相对 w:t 文本的偏移：
            rels = mm.start()
            runs[key]['matches'].append((rels, rels + len(mm.group(0)),
                                         plan['map'][canon], canon))
            hits += 1

if hits != plan['expected_hits']:
    die('命中数 %d ≠ 期望 %d' % (hits, plan['expected_hits']))
print('STEP B 命中 %d 处，涉及 %d 个 run' % (hits, len(runs)))

for rs, r in runs.items():
    text, op, rpr = r['text'], r['open'], r['rpr']
    ms = sorted(r['matches'])
    # 校验相对偏移落在文本内、且互不重叠
    prev = 0
    for (s, e, rf, dp) in ms:
        if not (0 <= s < e <= len(text)): die('偏移越界: %s' % text)
        if s < prev: die('重叠匹配: %s' % text)
        prev = e
    out = ''
    pos = 0
    for (s, e, rf, dp) in ms:
        if s > pos:
            seg = text[pos:s]
            out += op + rpr + '<w:t xml:space="preserve">' + escapes(seg) + '</w:t></w:r>'
        out += op + rpr + '<w:fldChar w:fldCharType="begin"/></w:r>'
        out += op + rpr + '<w:instrText xml:space="preserve"> REF ' + rf + ' \\h </w:instrText></w:r>'
        out += op + rpr + '<w:fldChar w:fldCharType="separate"/></w:r>'
        out += op + rpr + '<w:t>' + escapes(dp) + '</w:t></w:r>'
        out += op + rpr + '<w:fldChar w:fldCharType="end"/></w:r>'
        pos = e
    if pos < len(text):
        seg = text[pos:]
        out += op + rpr + '<w:t xml:space="preserve">' + escapes(seg) + '</w:t></w:r>'
    edits.append((rs, r['re'], out))

# ---------- STEP A: 给题注补 _Ref 书签（包编号区） ----------
for bd in plan['binds']:
    cap, nm, bid = bd['caption'], bd['ref'], bd['bookmark_id']
    found = None
    for i, (a, b) in enumerate(P):
        q = x[a:b]
        instrs = [t.strip() for t in INSTR.findall(q)]
        if not any(t.startswith('SEQ 图') or t.startswith('SEQ 表') for t in instrs): continue
        t = ''.join(WT.findall(q))
        if t.startswith(cap + ' ') or t.startswith(cap + '\u3000'):
            found = (i, a, b); break
    if not found: die('未找到题注 %s' % cap)
    i, a, b = found
    q = x[a:b]
    if re.search(r'<w:bookmarkStart[^>]*w:name="_Ref', q):
        die('题注 %s 已有 _Ref 书签，冲突' % cap)
    mr1 = re.search(r'<w:r(?: [^>]*)?>', q)
    if not mr1: die('题注 %s 找不到首个 w:r' % cap)
    ins1 = a + mr1.start()
    ppr_end = q.find('</w:pPr>')
    if ppr_end == -1 or ins1 < a + ppr_end:
        die('题注 %s 起点异常（无 pPr 或 r 在 pPr 内）' % cap)
    pos2 = q.find('<w:bookmarkStart')
    if pos2 == -1: die('题注 %s 找不到编号区终点' % cap)
    ins2 = a + pos2
    if not x[:ins2].endswith('</w:r>'):
        die('题注 %s 编号区终点前不是 </w:r>' % cap)
    edits.append((ins1, ins1, '<w:bookmarkStart w:id="%d" w:name="%s"/>' % (bid, nm)))
    edits.append((ins2, ins2, '<w:bookmarkEnd w:id="%d"/>' % bid))
print('STEP A 补书签 %d 个' % len(plan['binds']))

# ---------- 统一应用（从后往前） ----------
edits.sort(key=lambda e: (e[0], e[1]))
for k in range(1, len(edits)):
    if edits[k][0] < edits[k-1][1]:
        die('编辑区间重叠: %s / %s' % (edits[k-1][:2], edits[k][:2]))
for (s, e, rep) in reversed(edits):
    x = x[:s] + rep + x[e:]

# ---------- 不变量检查 ----------
z0 = zipfile.ZipFile(IN)
d0 = z0.read('word/document.xml').decode('utf-8')
def cnt(s, t): return s.count(t)
for t in ('<w:tbl>', '<w:drawing>'):
    if cnt(d0, t) != cnt(x, t): die('不变量破坏: %s' % t)
def npara(s):
    TAG = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>')
    return len(TAG.findall(s))
if npara(d0) != npara(x): die('段落数变化')

# ---------- 打包 ----------
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'word/document.xml':
        data = x.encode('utf-8')
    zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
    zi.compress_type = item.compress_type
    zi.external_attr = item.external_attr
    zout.writestr(zi, data)
zout.close()
print('OUT =', OUT)
print('document.xml 长度 %d -> %d (%+d)' % (N0, len(x), len(x)-N0))
print('w:p %d, w:tbl %d, w:drawing %d' % (npara(x), cnt(x,'<w:tbl>'), cnt(x,'<w:drawing>')))
print('fldChar 总数 %d -> %d' % (cnt(d0,'<w:fldChar'), cnt(x,'<w:fldChar')))
print('OK')
