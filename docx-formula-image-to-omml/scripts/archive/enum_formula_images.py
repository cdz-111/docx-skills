# -*- coding: utf-8 -*-
"""（历史脚本，已归档；路径参数须由环境变量或显式传参提供）
枚举第 1 章（para<1720）里的公式图：区分行内/行间，计数、导出。
"""
import zipfile, re, json, os

WORK = os.environ.get('DOCXFMT_WORKDIR')
SRC = os.path.join(WORK, '9-28.backup-20260928-005343.docx')
z = zipfile.ZipFile(SRC)
xml = z.read('word/document.xml').decode('utf-8')
rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
REL = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
ps = PARA.findall(xml)
T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')

disp, inline = [], []
for i, q in enumerate(ps):
    if i >= 1720:
        break
    imgs = re.findall(r'<w:drawing>(.*?)</w:drawing>', q, re.S)
    if not imgs:
        continue
    other = re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S)
    txt = ''.join(T.findall(other)).strip()
    for d in imgs:
        em = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', d)
        rid = re.search(r'r:embed="(rId\d+)"', d)
        cx, cy = (int(em.group(1)), int(em.group(2))) if em else (0, 0)
        h = cy / 360000
        rec = dict(para=i, rid=rid.group(1) if rid else '?',
                   media=REL.get(rid.group(1), '?').replace('media/', '') if rid else '?',
                   w=round(cx / 360000, 2), h=round(h, 2), txt=txt)
        if h <= 2.8:
            (disp if not txt else inline).append(rec)

print('第 1 章 公式图：行间 %d 个 / 行内 %d 个' % (len(disp), len(inline)))
print('\n--- 行间（独立成段）---')
for k, r in enumerate(disp, 1):
    print('  %2d) para %-5d %-14s %5.2f x %4.2f' % (k, r['para'], r['media'], r['w'], r['h']))
print('\n--- 行内（夹在正文里）---')
for k, r in enumerate(inline, 1):
    print('  %2d) para %-5d %-14s %5.2f x %4.2f  ctx=%s' %
          (k, r['para'], r['media'], r['w'], r['h'], r['txt'][:34]))

json.dump(dict(disp=disp, inline=inline),
          open(os.path.join(WORK, 'ch1_formulas.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

# 导出所有第1章公式图
outdir = os.path.join(WORK, 'img928')
os.makedirs(outdir, exist_ok=True)
seen = set()
for r in disp + inline:
    m = 'word/media/' + r['media']
    if m in seen:
        continue
    seen.add(m)
    if m in z.namelist():
        open(os.path.join(outdir, r['media']), 'wb').write(z.read(m))
print('\n已导出 %d 个媒体文件 → img928/' % len(seen))

# 第1章是否已有原生公式/编号
n_math = 0
for i, q in enumerate(ps):
    if i >= 1720:
        break
    n_math += len(re.findall(r'<m:oMath>', q))
print('第 1 章原生 oMath 实例数: %d' % n_math)
