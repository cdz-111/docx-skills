# -*- coding: utf-8 -*-
"""步骤1：摸底 + 分类。
用法: python 1_classify.py <src.docx> <workdir>
产出: doc_dump.txt / items.json / queue.json / uniq.json / disp_list.json / inline_list.json / img/
"""
import zipfile, re, json, os, sys, hashlib

SRC, W = sys.argv[1], sys.argv[2]
IMG = os.path.join(W, 'img')
os.makedirs(IMG, exist_ok=True)

z = zipfile.ZipFile(SRC)
xml = z.read('word/document.xml').decode('utf-8')
rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))

PARA = re.compile(r'<w:p [^>]*>.*?</w:p>|<w:p/>', re.S)
paras = PARA.findall(xml)

# 段落大纲：段号 / pStyle / 可见文本 / 图数 / oMath 数
dump, sec, cur = [], {}, ''
for i, p in enumerate(paras):
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.S))
    st = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    st = st.group(1) if st else ''
    ni, nm = p.count('<w:drawing>'), p.count('<m:oMath')
    line = '[%d] (%s) %s' % (i, st, txt)
    if ni:
        line += '  <<IMG x%d>>' % ni
    if nm:
        line += '  <<MATH x%d>>' % nm
    dump.append(line)
    if st in ('1', '2', '3', '4', '5', '6', '7') and txt.strip():
        cur = txt.split()[0]
    sec[i] = cur
open(os.path.join(W, 'doc_dump.txt'), 'w', encoding='utf-8').write('\n'.join(dump))

# 逐个 drawing 提取
items = []
for pi, p in enumerate(paras):
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.S)).strip()
    for d in re.finditer(r'<w:drawing>.*?</w:drawing>', p, re.S):
        blk = d.group(0)
        ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', blk)
        emb = re.search(r'r:embed="(rId\d+)"', blk)          # 第一个即为位图本体
        if not ext or not emb:
            continue
        cx, cy = int(ext.group(1)), int(ext.group(2))
        tgt = relmap.get(emb.group(1), '')
        base = os.path.basename(tgt) if tgt else ''
        if base:
            try:
                open(os.path.join(IMG, base), 'wb').write(z.read('word/' + tgt))
            except KeyError:
                pass
        items.append(dict(para=pi, w=round(cx / 360000, 2), h=round(cy / 360000, 2),
                          media=tgt, file=base, ptextlen=len(txt), ptext=txt[:60],
                          sec=sec.get(pi, '')))
json.dump(items, open(os.path.join(W, 'items.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

# md5 去重
for it in items:
    p = os.path.join(IMG, it['file'])
    it['md5'] = hashlib.md5(open(p, 'rb').read()).hexdigest() if os.path.exists(p) else ''

# 分类：段落无正文且够矮 = 展示公式；有正文且矮 = 行内公式；其余 = 插图
disp, inline, figure = [], [], []
for it in items:
    if it['ptextlen'] > 0:
        (inline if it['h'] < 1.5 else figure).append(it)
    elif it['h'] >= 2.6:
        figure.append(it)        # 高度 2.6~7cm 的"独段图"务必再人工过一遍联系表！
    else:
        disp.append(it)

json.dump({'display': disp, 'inline': inline, 'figure': figure},
          open(os.path.join(W, 'queue.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

uset = {}
for it in disp + inline:
    uset.setdefault(it['md5'], it)
uniq = list(uset.values())
json.dump(uniq, open(os.path.join(W, 'uniq.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
json.dump(disp, open(os.path.join(W, 'disp_list.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
json.dump(inline, open(os.path.join(W, 'inline_list.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

print('drawings:', len(items))
print('display eq images:', len(disp), ' inline images (occurrences):', len(inline),
      ' figures:', len(figure))
print('unique formula images to transcribe:', len(uniq))
import collections
print('display per section:', dict(collections.Counter(d['sec'] for d in disp)))
print('NOTE: 请人工过一遍 queue.json 里 figure 中 h 在 2.6~7cm 的条目，确认没有漏掉的公式块。')
