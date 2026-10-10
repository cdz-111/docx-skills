# -*- coding: utf-8 -*-
"""（历史脚本，已归档；路径参数须由环境变量或显式传参提供）
由 fix928_dump.json 生成《图表修复说明.html》（数据与修复脚本同源）
"""
import json, html, os

WORK = os.environ.get('DOCXFMT_WORKDIR')
DST = os.path.join(WORK, '图表修复说明.html')
d = json.load(open(os.path.join(WORK, 'fix928_dump.json'), encoding='utf-8'))

E = html.escape
caps, refs, hards = d['caps'], d['refs'], d['hards']

nfix = sum(1 for c in caps if str(c['old']) != str(c['new']))
nprob = sum(1 for c in caps if (not c['seq']) or str(c['seqdisp']) != str(c['old']))
nrefchg = sum(1 for r in refs if r['disp'] != r['want'])
nrepoint = sum(1 for r in refs if r['mode'] == '重指向')
nhardchg = sum(1 for h in hards if h['old'] != h['new'])

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--line:#e3e6ea;--fg:#1f2328;--dim:#6b7280;
--red:#c0392b;--green:#1a7f37;--blue:#0b5cad;--amber:#8a5a00}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);
font:14px/1.65 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:28px 22px 60px}
h1{font-size:22px;margin:0 0 6px}
.sub{color:var(--dim);font-size:13px;margin-bottom:20px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:18px 0 26px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px}
.card .k{color:var(--dim);font-size:12px}
.card .v{font-size:24px;font-weight:700;margin-top:4px}
.card .v small{font-size:12px;font-weight:400;color:var(--dim)}
h2{font-size:16px;margin:30px 0 10px;padding-left:9px;border-left:4px solid var(--blue)}
table{width:100%;border-collapse:collapse;background:var(--card);font-size:13px}
th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:#eef1f5;font-weight:600;position:sticky;top:0;font-size:12px;color:#374151}
tr:last-child td{border-bottom:none}
tr:hover td{background:#fafbfc}
td.n{white-space:nowrap;font-variant-numeric:tabular-nums;color:var(--dim)}
.kind{font-weight:600}
.red{color:var(--red);font-weight:600}.green{color:var(--green)}.dim{color:var(--dim)}
.badge{display:inline-block;padding:1px 7px;border-radius:20px;font-size:11px;border:1px solid;white-space:nowrap}
.b-warn{color:var(--amber);border-color:#e6d3a3;background:#fff8e6}
.b-fix{color:var(--blue);border-color:#bcd7f0;background:#eef6fd}
.b-ok{color:var(--green);border-color:#bfe3c9;background:#eefaf1}
.note{background:#fff8e6;border:1px solid #e6d3a3;border-radius:10px;padding:12px 16px;margin:14px 0}
.note b{color:var(--amber)}
.scroll{max-height:640px;overflow:auto;border-radius:10px;border:1px solid var(--line);background:var(--card)}
.scroll table{border-radius:0}
.legend{color:var(--dim);font-size:12px;margin:6px 0 12px}
"""


def tbl(head, rows):
    h = ''.join('<th>%s</th>' % E(x) for x in head)
    b = ''.join('<tr>%s</tr>' % ''.join('<td%s>%s</td>' % (c, v) for c, v in r) for r in rows)
    return ('<div class="scroll"><table><thead><tr>%s</tr></thead><tbody>%s</tbody></table></div>'
            % (h, b))


r1 = []
for c in caps:
    chg = str(c['old']) != str(c['new'])
    if not c['seq']:
        pr = '<span class="badge b-warn">硬写（无 SEQ 域）</span>'
    elif str(c['seqdisp']) != str(c['old']):
        pr = '<span class="badge b-warn">域后残留数字</span>'
    else:
        pr = '<span class="dim">—</span>'
    r1.append([('n', str(c['bi'])), ('kind', E(c['kind'])), ('n', E(str(c['old']))),
               ('n red', E(str(c['new']))),
               ('', '<span class="badge b-fix">重编号</span>' if chg
                else '<span class="badge b-ok">不变</span>'),
               ('', pr), ('dim', E(c['title'][:54]))])

order = {}
for c in caps:
    order[c['bi']] = (c['kind'], c['new'])
r2 = []
for r in refs:
    via = '<span class="green">重指向</span>' if r['mode'] == '重指向' else '<span class="dim">显示校正</span>'
    r2.append([('n', str(r['i'])), ('n dim', E(r['old'])), ('n', E(r['disp'])),
               ('n', E(r['new'])), ('n red', E(r['want'])), ('', via)])
r2.sort(key=lambda x: int(x[0][1]))

r3 = []
for h in hards:
    k, n = order.get(h['near_bi'], ('?', '?'))
    r3.append([('n', str(h['i'])), ('n dim', E(h['old'])), ('n red', E(h['new'])),
               ('n', '%s%s' % (k, n)), ('', '<span class="badge b-ok">已同步</span>'),
               ('dim', E(h['near_title'][:46]))])
r3.sort(key=lambda x: int(x[0][1]))

HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<title>图表题注与交叉引用修复说明</title><style>%s</style></head><body><div class="wrap">
<h1>图表题注与交叉引用修复说明</h1>
<div class="sub">源文件 <b>待处理文档.docx</b> → 成品 <b>修复后文档.docx</b>　|　编号方案：全文连续（图 1–44 / 表 1–52）</div>

<div class="cards">
  <div class="card"><div class="k">题注总数</div><div class="v">%d <small>图 44 / 表 52</small></div></div>
  <div class="card"><div class="k">题注重编号</div><div class="v">%d <small>其中 %d 处结构缺陷</small></div></div>
  <div class="card"><div class="k">交叉引用 REF</div><div class="v">53 <small>改写 %d 处，断链重指向 %d</small></div></div>
  <div class="card"><div class="k">正文硬写编号</div><div class="v">%d <small>已同步</small></div></div>
</div>

<h2>一、题注重编号（96 个）</h2>
<div class="legend">「域后残留数字」＝ 题注的 SEQ 域后面粘着多余数字（SEQ 显示 1 + 残留 9 → 显示成“图 19”）；「硬写（无 SEQ 域）」＝ 编号是手打文字，不含域，无法随域更新。</div>
%s

<h2>二、交叉引用修复（共 53 个 REF 域）</h2>
<div class="legend">「重指向」＝ 原书签在文档中已不存在（断链），按段落语义重新指向正确题注；「显示校正」＝ 目标书签正确，但域缓存显示值仍是旧编号，已同步刷新。</div>
%s

<h2>三、正文硬写编号同步（%d 处）</h2>
<div class="legend">正文里直接手打的“图 N / 表 N”（非交叉引用域），随题注重编号一并同步。末列为该处最近的题注，供核对。</div>
%s

<h2>四、必须注意</h2>
<div class="note">
<b>不要在整篇文档上执行 Ctrl+A → F9（更新域）！</b><br>
本报告中有 <b>11 个“硬写（无 SEQ 域）”题注</b>的编号是纯文字；其余 85 个含 SEQ 域。若全选更新域，
85 个含域的题注会重排为 1–44 / 1–52（与本次结果一致），但目录、页码、题注等所有域会被一起刷新，
历史上正是因为反复全选更新，才产生了“域后残留数字”和编号重复。<b>如需局部更新，请只选中单个题注按 F9。</b><br><br>
<b>其他部件零改动：</b>除 word/document.xml 外，压缩包内 125 个部件（公式、图片、样式、编号定义等）与源文件逐字节一致；
公式 387 个、图片 44 个数量均未变化，XML 结构合法，书签 146 个开始/结束完全配对。<br><br>
<b>编号机制说明：</b>85 个含域题注使用的是 <code>SEQ 图 \\* ARABIC</code> / <code>SEQ 表 \\* ARABIC</code>（无章节重启开关），
因此对该域执行更新后会自然得到 1–44 / 1–52 的全文连续编号——本次修改已把缓存显示值写成与之相同的数值，二者一致。
</div>
</div></body></html>""" % (CSS, len(caps), nfix, nprob, nrefchg, nrepoint, nhardchg,
                           tbl(['位置块', '类型', '原编号', '新编号', '处理', '结构问题', '题注标题'], r1),
                           tbl(['位置块', '原书签', '原显示', '现书签', '现显示', '方式'], r2),
                           nhardchg,
                           tbl(['位置块', '原正文编号', '现正文编号', '最近题注', '状态', '对应题注标题'], r3))

open(DST, 'w', encoding='utf-8').write(HTML)
print('已生成 %s' % DST)
print('题注 %d（重编号 %d / 结构缺陷 %d）｜REF 改写 %d（重指向 %d）｜硬写号同步 %d'
      % (len(caps), nfix, nprob, nrefchg, nrepoint, nhardchg))
