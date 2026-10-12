#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
docx 图片智能压缩器（生成"精简版"/可分发版）

逐张图片在  JPEG(q) / PNG-8(256色) / 原图  三者中取**体积最小**者：
  - 图形、截图、图表类（颜色数少）→ 多半 PNG-8 胜出，又小又比 JPEG 锐利
  - 照片、3D 渲染类（颜色数多）    → 多半 JPEG 胜出
  - 重编码反而变大            → 保留原图

保证：
  - **分辨率完全不变**（除非显式 --dpi）
  - 透明通道统一压平为白底（与 WPS「压缩图片」行为一致）
  - `word/document.xml` / `styles.xml` 等**非图片部件逐字节原样复制**
  - 仅在图片扩展名变化时同步更新 `.rels` 的 Target
  - 落盘后自检：document.xml 必须逐字节相同；w:p/w:tbl/w:drawing/fldChar 计数不变；
    图片 rel 无断链；`r:embed` 引用 rId 全部有映射 —— 任一失败即删除输出、退出 5

用法：
  python compress_images.py <src.docx> <out.docx> [jpeg质量=88] [--expect-md5 <hex>]
                           [--no-png8] [--dpi 200]
环境变量回退：DOCX_IN / DOCX_OUT / DOCX_IMG_QUALITY

退出码：0 成功 ｜ 5 断言失败（未落盘或已回滚）｜ 6 用法错误
"""
import sys, os, io, re, zipfile, hashlib, json

def usage(code):
    print(__doc__)
    sys.exit(code)

args = sys.argv[1:]
if not args:
    a, b = os.environ.get('DOCX_IN'), os.environ.get('DOCX_OUT')
    if not (a and b):
        usage(6)
    args = [a, b]
SRC, OUT = args[0], args[1]
rest = args[2:]
Q, EXPECT, USE_PNG8, DPI = 88, None, True, None
i = 0
while i < len(rest):
    t = rest[i]
    if t == '--expect-md5':
        i += 1; EXPECT = rest[i].lower()
    elif t == '--no-png8':
        USE_PNG8 = False
    elif t == '--dpi':
        i += 1; DPI = int(rest[i])
    elif re.fullmatch(r'\d+', t):
        Q = int(t)
    else:
        print('未知参数:', t); usage(6)
    i += 1
# 环境变量兜底
if Q == 88 and os.environ.get('DOCX_IMG_QUALITY'):
    Q = int(os.environ['DOCX_IMG_QUALITY'])

try:
    from PIL import Image
except ImportError:
    print('需要 Pillow：pip install pillow'); sys.exit(6)

def die(m):
    print('✖ 中止（未落盘）:', m); sys.exit(5)

src_bytes = open(SRC, 'rb').read()
md5 = hashlib.md5(src_bytes).hexdigest()
if EXPECT and md5 != EXPECT:
    die('源文件指纹不符：期望 %s，实际 %s（Word 可能重存过 → 重新出单）' % (EXPECT, md5))
print('源 %s  md5=%s  %.1f MB' % (os.path.basename(SRC), md5, len(src_bytes)/1048576))

z = zipfile.ZipFile(SRC)
names = z.namelist()
doc0 = z.read('word/document.xml')
d0s = doc0.decode('utf-8')

def stat(s):
    return {'w:p': len(re.findall(r'<w:p(?:\s[^>]*)?>', s)) + len(re.findall(r'<w:p/>', s)),
            'w:tbl': s.count('<w:tbl>'), 'w:drawing': s.count('<w:drawing>'),
            'fldChar': len(re.findall(r'<w:fldChar\b', s))}
BASE = stat(d0s)
print('不变量基线:', BASE)

media = [t for t in names if t.startswith('word/media/') and not t.endswith('/')]
if not media:
    die('该文档没有内嵌图片，无需压缩')

# 每张图的显示尺寸（EMU → cm），用于可选降采样
EMU = 914400.0
rels0 = z.read('word/_rels/document.xml.rels').decode('utf-8')
rid2t = dict(re.findall(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"', rels0))
use = {}
for cx, cy, body in re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"/>(.*?)</w:drawing>', d0s, re.S):
    m = re.search(r'r:embed="(rId\d+)"', body)
    if not m: continue
    t = rid2t.get(m.group(1))
    if not t: continue
    k = 'word/' + t.lstrip('/')
    w = int(cx)/EMU*2.54; h = int(cy)/EMU*2.54
    if k not in use or w > use[k][0]:
        use[k] = (w, h)

def flat(im):
    if im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info):
        r = im.convert('RGBA')
        bg = Image.new('RGB', im.size, (255, 255, 255))
        bg.paste(r, mask=r.split()[-1])
        return bg
    return im.convert('RGB')

def enc_jpg(im):
    b = io.BytesIO(); im.save(b, 'JPEG', quality=Q, optimize=True, progressive=True); return b.getvalue()
def enc_png8(im):
    q = im.quantize(colors=256, method=Image.MEDIANCUT, dither=Image.FLOYDSTEINBERG)
    b = io.BytesIO(); q.save(b, 'PNG', optimize=True); return b.getvalue()

new_media, rename, log = {}, {}, []
for t in media:
    raw = z.read(t)
    im = Image.open(io.BytesIO(raw)); size0, mode0 = im.size, im.mode
    rgb = flat(im)
    if DPI and t in use:
        mpx = max(1, int(round(use[t][0]/2.54*DPI)))
        if rgb.width > mpx*1.02:
            r = mpx/rgb.width
            rgb = rgb.resize((mpx, max(1, int(round(rgb.height*r)))), Image.LANCZOS)
    cand = [('orig', raw, t)]
    try: cand.append(('jpeg', enc_jpg(rgb), re.sub(r'\.[A-Za-z0-9]+$', '.jpeg', t)))
    except Exception as e: print('  [warn] %s JPEG 失败 %s' % (t, e))
    if USE_PNG8:
        try: cand.append(('png8', enc_png8(rgb), re.sub(r'\.[A-Za-z0-9]+$', '.png', t)))
        except Exception as e: print('  [warn] %s PNG8 失败 %s' % (t, e))
    kind, data, name = min(cand, key=lambda c: len(c[1]))
    if name != t: rename[t] = name
    new_media[name] = data
    log.append({'file': t.split('/')[-1], 'kind': kind, 'orig': len(raw), 'new': len(data),
                'new_file': name.split('/')[-1], 'px': '%dx%d' % size0, 'mode': mode0})

t0 = sum(len(z.read(x)) for x in media); t1 = sum(len(v) for v in new_media.values())
sel = {}
for r in log: sel[r['kind']] = sel.get(r['kind'], 0) + 1
print('图片 %d 张 %s：%.1f MB → %.1f MB' % (len(media), sel, t0/1048576, t1/1048576))

new_rels = {}
for t in names:
    if t.endswith('.rels'):
        r = z.read(t).decode('utf-8')
        for o, n in rename.items():
            r = r.replace('Target="%s"' % o.replace('word/', ''), 'Target="%s"' % n.replace('word/', ''))
        new_rels[t] = r.encode('utf-8')

MEDIA = set(media)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
    for t in names:
        if t.endswith('/'):
            zo.writestr(zipfile.ZipInfo(t), b'', compress_type=zipfile.ZIP_STORED); continue
        if t in MEDIA: continue
        if t in new_rels:
            zo.writestr(zipfile.ZipInfo(t), new_rels[t], compress_type=zipfile.ZIP_DEFLATED)
        else:
            zi = z.getinfo(t)
            zo.writestr(zipfile.ZipInfo(t, date_time=zi.date_time), z.read(t), compress_type=zipfile.ZIP_DEFLATED)
    for name, data in new_media.items():
        zo.writestr(zipfile.ZipInfo(name), data, compress_type=zipfile.ZIP_STORED)

# ---- 落盘后自检 ----
zz = zipfile.ZipFile(OUT)
if zz.read('word/document.xml') != doc0:
    os.remove(OUT); die('document.xml 被改动（应逐字节相同）')
if stat(zz.read('word/document.xml').decode('utf-8')) != BASE:
    os.remove(OUT); die('结构不变量变化')
for t in names:
    if t.endswith('.rels') and t != 'word/_rels/document.xml.rels' and z.read(t) != zz.read(t):
        os.remove(OUT); die('%s 被误改' % t)
r1 = zz.read('word/_rels/document.xml.rels').decode('utf-8')
have = set(zz.namelist())
bad = [t for t in re.findall(r'Target="([^"]*media/[^"]+)"', r1) if ('word/' + t) not in have]
if bad:
    os.remove(OUT); die('图片 rel 断链 %d: %s' % (len(bad), bad[:5]))
rl = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', r1))
miss = [e for e in set(re.findall(r'r:(?:embed|id|link)="(rId\d+)"', zz.read('word/document.xml').decode('utf-8'))) if e not in rl]
if miss:
    os.remove(OUT); die('缺 rId 映射 %s' % miss[:5])

print('✔ 落盘:', OUT)
print('  %.1f MB → %.1f MB (−%.1f%%)' % (len(src_bytes)/1048576, os.path.getsize(OUT)/1048576,
                                       (1 - os.path.getsize(OUT)/len(src_bytes))*100))
print('  输出 md5:', hashlib.md5(open(OUT, 'rb').read()).hexdigest())
print('  决策:', sel, ' 改名:', len(rename))
