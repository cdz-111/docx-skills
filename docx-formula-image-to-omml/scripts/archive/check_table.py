# -*- coding: utf-8 -*-
"""生成公式逐条对照核对表 HTML：原图 vs 转换后公式（MathML，浏览器原生渲染）。"""
import subprocess, re, os, base64, json

WORK = r'D:\appdata\workbuddy\2026-09-24-23-47-06\work'
IMG = os.path.join(WORK, 'img928')
PANDOC = r'C:\Users\14711\anaconda3\Scripts\pandoc.exe'
omml_src = open(os.path.join(WORK, 'x928_latex.py'), encoding='utf-8').read()
ns = {}
exec(omml_src.split("items = ")[0], ns)
DISP, INLINE = ns['DISP'], ns['INLINE']

# 行内出现顺序（按段落）
INLINE_ORDER = [
    (223, 'image4.jpeg'), (223, 'image5.jpeg'), (223, 'image6.jpeg'),
    (227, 'image9.jpeg'), (230, 'image12.jpeg'), (230, 'image5.jpeg'),
    (241, 'image16.jpeg'), (241, 'image17.jpeg'), (241, 'image18.jpeg'),
    (243, 'image20.jpeg'), (243, 'image21.jpeg'), (243, 'image22.jpeg'), (243, 'image23.jpeg'),
    (243, 'image24.jpeg'), (243, 'image25.jpeg'), (243, 'image26.jpeg'), (243, 'image25.jpeg'),
    (244, 'image16.jpeg'), (244, 'image27.jpeg'), (244, 'image17.jpeg'), (244, 'image28.jpeg'),
    (246, 'image30.jpeg'), (250, 'image32.jpeg'), (250, 'image33.jpeg'),
    (889, 'image42.jpeg'), (889, 'image43.jpeg'),
    (1711, 'image51.jpeg'), (1711, 'image52.jpeg'),
    (1716, 'image55.jpeg'), (1716, 'image56.jpeg'),
]

items = []
for k, (p, tex) in enumerate(DISP, 1):
    items.append(dict(kind='disp', no=k, para=p, img='image%d.jpeg' % {"224":7,"226":8,"228":10,"229":11,"231":13,"232":14,"242":19,"245":29,"248":31,"878":37,"880":38,"883":39,"885":40,"888":41,"1712":53,"1717":57}[str(p)], tex=tex))
for k, (p, f) in enumerate(INLINE_ORDER, 1):
    items.append(dict(kind='inl', no=k, para=p, img=f, tex=INLINE[f]))

# 批量 pandoc：用标记切分
mdp, htp = os.path.join(WORK, '_mk.md'), os.path.join(WORK, '_mk.html')
md = []
for it in items:
    md.append('MKR%s%dMKR' % ('d' if it['kind'] == 'disp' else 'i', it['no']))
    md.append('$$' + it['tex'] + '$$')
open(mdp, 'w', encoding='utf-8').write('\n\n'.join(md))
subprocess.run([PANDOC, '-f', 'markdown', '-t', 'html', '--mathml', '-o', htp, mdp], check=True)
html = open(htp, encoding='utf-8').read()
parts = re.split(r'MKR[di]\d+MKR', html)
html2 = re.sub(r'(MKR[di]\d+MKR)', r'<span class="mkr">\1</span>', html)
parts = re.split(r'<span class="mkr">MKR([di])(\d+)MKR</span>', html2)
mm = {}
for a, b, c in zip(parts[1::3], parts[2::3], parts[3::3]):
    mm[('%s%s' % (a, b))] = c.strip()
print('切分到 %d 段 MathML（预期 %d）' % (len(mm), len(items)))
assert len(mm) == len(items), '切分失败: %d' % len(mm)


def b64(f):
    return 'data:image/jpeg;base64,' + base64.b64encode(open(os.path.join(IMG, f), 'rb').read()).decode()


def row(it, ml):
    key = ('d%d' if it['kind'] == 'disp' else 'i%d') % it['no']
    tag = '(第 %d 条 · para %d)' % (it['no'], it['para']) if it['kind'] == 'disp' else '第 %d 处 · para %d' % (it['no'], it['para'])
    return ('<tr><td class="t">%s<br><code>%s</code></td>'
            '<td class="a"><img src="%s"></td><td class="b">%s</td></tr>'
            % (tag, it['img'], b64(it['img']), mm[key]))


disps = [it for it in items if it['kind'] == 'disp']
inls = [it for it in items if it['kind'] == 'inl']

CSS = """
body{background:#1b1d22;color:#e6e6e6;font-family:"Microsoft YaHei",Segoe UI,sans-serif;margin:0;padding:22px 26px}
h1{font-size:19px;margin:0 0 4px}
h2{font-size:16px;margin:26px 0 8px;color:#78d6ff;border-bottom:1px solid #2e3238;padding-bottom:5px}
p.sub{color:#9aa2ad;font-size:12.5px;margin:0 0 8px}
table{border-collapse:collapse;width:100%;font-size:13px}
td{border:1px solid #31353c;padding:7px 9px;vertical-align:middle}
td.t{width:132px;color:#9aa2ad;font-size:12px;white-space:nowrap}
td.a{width:250px;background:#fff}
td.b{background:#23262c}
img{max-width:100%;height:auto;display:block}
math{font-size:1.25em}
code{color:#5f6b7a;font-size:11px}
.badge{display:inline-block;background:#2b6b41;color:#c8f7d8;border-radius:3px;padding:1px 7px;font-size:11.5px;margin-left:6px}
.note{background:#2a2418;border-left:3px solid #b98b2e;padding:8px 12px;font-size:12.5px;color:#e2d3ab;margin:14px 0}
"""

out = ['<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">',
       '<title>9-28 公式转换核对表</title><style>%s</style></head><body>' % CSS,
       '<h1>中期科技报告 9-28 · 第 1 章公式转换核对表</h1>',
       '<p class="sub">左列 = 原文档里那张公式图（截自你的 docx），右列 = 转换后的可编辑公式。'
       '逐条比对即可确认识读无误。<span class="badge">%d 条行间 + %d 处行内</span></p>' % (len(disps), len(inls)),
       '<div class="note">行间公式已按全文连续编号排好：第 1 章这 16 条 = (1)–(16)；'
       '第 2/3 章原有 (20)–(71) 整体前移为 (17)–(68)；§3.4.2 那 7 条无编号的补为 (69)–(75)。'
       '正文里唯一的引用「公式（33）」已同步改为「公式（30）」。</div>',
       '<h2>一、行间公式（编号 1–16）</h2><table>',
       ''.join(row(it, mm['d%d' % it['no']]) for it in disps),
       '</table>',
       '<h2>二、行内符号（原位替换，30 处）</h2><table>',
       ''.join(row(it, mm['i%d' % it['no']]) for it in inls),
       '</table></body></html>']

open(os.path.join(WORK, '公式核对表.html'), 'w', encoding='utf-8').write('\n'.join(out))
print('已写出 公式核对表.html  (%d KB)'
      % (os.path.getsize(os.path.join(WORK, '公式核对表.html')) // 1024))
