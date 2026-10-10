# -*- coding: utf-8 -*-
"""建 (kind, num) -> (ref_name, rendered_text) 映射；rendered_text 为 _Ref 书签内可见文本。"""
import zipfile, re, sys, json

DOC = sys.argv[1] if len(sys.argv) > 1 else None
if not DOC:
    sys.exit('用法: build_caption_ref_map.py <in.docx>')

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
BM = re.compile(r'<w:bookmarkStart\b[^>]*?/>')
BME = re.compile(r'<w:bookmarkEnd\b[^>]*?/>')
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
                st = None; depth = 0
        else:
            if depth == 0:
                st = m.start()
            depth += 1
    return res


z = zipfile.ZipFile(DOC)
x = z.read('word/document.xml').decode('utf-8')
z.close()
P = paras(x)

# 书签 id -> end 偏移
ends = {}
for m in BME.finditer(x):
    i = re.search(r'w:id="(\d+)"', m.group(0))
    if i:
        ends[i.group(1)] = m.start()

M = {}
for k, (a, b) in enumerate(P):
    q = x[a:b]
    if not re.search(r'SEQ\s+[图表]', q):
        continue
    vis = ''.join(WT.findall(q))
    mm = re.match(r'^\s*([图表])\s*(\d+)', vis)
    if not mm:
        continue
    kind, num = mm.group(1), int(mm.group(2))
    for m2 in BM.finditer(q):
        t = m2.group(0)
        i = re.search(r'w:id="(\d+)"', t)
        n = re.search(r'w:name="([^"]*)"', t)
        if not (i and n):
            continue
        nm = n.group(1)
        if not nm.startswith('_Ref'):
            continue
        s0, e0 = a + m2.end(), ends.get(i.group(1))
        if e0 is None:
            continue
        inner = x[s0:e0]
        rendered = ''.join(WT.findall(inner))
        M['%s%d' % (kind, num)] = (nm, rendered, e0 - s0)

print('映射条目 %d' % len(M))
json.dump(M, open('_map_ref.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
for kk in ['图16', '图17', '图21', '图24', '图32', '表17', '表18', '表22', '表29', '表48']:
    print('  %-6s -> %s' % (kk, M.get(kk)))
