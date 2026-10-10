# -*- coding: utf-8 -*-
"""步骤2：生成识读联系表（给多模态读图用）。
用法:
  python 2_make_sheets.py <workdir> display <start> <n>     # 单列，公式看清上下标
  python 2_make_sheets.py <workdir> inline  <start> <n>     # 3 列网格，符号放大
  python 2_make_sheets.py <workdir> tall                    # 列出 2.6~7cm 的独段图供人工确认
透明 PNG 必须合成到白底，否则会整张变黑。
"""
import json, os, sys
from PIL import Image, ImageDraw

W, kind = sys.argv[1], sys.argv[2]
IMG = os.path.join(W, 'img')


def load(f, cell_w, cell_h):
    im = Image.open(os.path.join(IMG, f))
    if im.mode in ('RGBA', 'LA', 'P'):
        im = im.convert('RGBA')
        bg = Image.new('RGBA', im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im).convert('RGB')
    else:
        im = im.convert('RGB')
    sc = min(cell_w / im.width, cell_h / im.height)
    sc = min(sc, 4) if sc > 1 else sc
    return im.resize((max(1, int(im.width * sc)), max(1, int(im.height * sc))), Image.LANCZOS)


if kind == 'tall':
    q = json.load(open(os.path.join(W, 'queue.json'), encoding='utf-8'))
    tall = [f for f in q['figure'] if f['ptextlen'] == 0 and 2.4 <= f['h'] < 7.5]
    tall.sort(key=lambda x: x['h'])
    print('需人工确认的独段高图（可能是多行公式块，也可能是插图）:', len(tall))
    for f in tall:
        print('  P%-6d %5.2f x %5.2f  %-16s %s' % (f['para'], f['w'], f['h'], f['file'], f['sec']))
    sys.exit(0)

lst = json.load(open(os.path.join(W, 'disp_list.json' if kind == 'display' else 'inline_list.json'),
                     encoding='utf-8'))
start, n = int(sys.argv[3]), int(sys.argv[4])
sel = lst[start:start + n]

if kind == 'display':
    COLS, CW, CH, LAB = 1, 1480, 78, 26
else:
    COLS, CW, CH, LAB = 3, 490, 78, 26

cells = []
for u in sel:
    cells.append((u, load(u['file'], CW - 16, CH - 8)))

ROWH = CH + LAB
rows = (len(cells) + COLS - 1) // COLS
sheet = Image.new('RGB', (CW * COLS, ROWH * rows), 'white')
d = ImageDraw.Draw(sheet)
for idx, (u, im) in enumerate(cells):
    r, c = divmod(idx, COLS)
    x, y = c * CW, r * ROWH
    d.rectangle([x, y, x + CW - 2, y + ROWH - 2], outline=(200, 200, 210))
    d.rectangle([x, y, x + CW - 2, y + LAB - 1], fill=(232, 232, 240))
    d.text((x + 5, y + 7), '%d  %s' % (start + idx, u['file']), fill=(0, 0, 170))
    sheet.paste(im, (x + 8 + (CW - 16 - im.width) // 2, y + LAB + (CH - 8 - im.height) // 2))

out = os.path.join(W, 'tr_%s_%d.png' % (kind, start))
sheet.save(out)
print(out, sheet.size, 'items', start, '..', start + len(cells) - 1)
print('按序号转写，再用脚本把序号映射回文件名（勿凭肉眼抄 image361/image381 这类文件名）。')
