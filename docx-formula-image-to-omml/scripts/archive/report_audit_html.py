# -*- coding: utf-8 -*-
"""把 cap928.json 渲染成可读的审计报告 HTML（深色主题）。"""
import json, os, html as H

WORK = r'D:\appdata\workbuddy\2026-09-24-23-47-06\work'
d = json.load(open(os.path.join(WORK, 'cap928.json'), encoding='utf-8'))
rows, refs, broken, hard = d['rows'], d['refs'], d['broken'], d['hard']

E = H.escape
CLS = {'OK': ('正常', '#3f7a52', '#c9f2d6'),
       'LEFTOVER': ('域后残留旧数字', '#8a3b3b', '#ffd6d6'),
       'CACHE': ('缓存与显示不符', '#8a6d2c', '#ffe9b8'),
       'HARD': ('纯手打、无编号域', '#7a5c33', '#f0dfc4')}

# 14 处断链 → 建议目标（按上下文/标题推断，人工核对过）
SUGGEST = {
    '_Ref2431': '图 16 场景认知融合的启发式推理智能指导总体框架',
    '_Ref28658': '表 23 不同离线MOO方法的平均HV排名：Synthetic、MO-NAS与MORL',
    '_Ref29069': '表 25 不同离线MOO方法的平均IGDoffline排名：Synthetic、MO-NAS',
    '_Ref2591': '图 20 面向导学任务的边缘—中心协同方法架构',
    '_Ref2627': '图 19 面向动态实验场景的边缘智能指导流程',
    '_Ref2702': '表 28 边缘计算关键机制与导学任务',
    '_Ref2757': '图 21 导学系统的分层协同计算范式',
    '_Ref29291': '图 22 多模态细粒度感知与显式证据链推理框架（MuEmo-R1）',
    '_Ref2855': '表 31 代表性在轨任务数据来源',
    '_Ref29393': '图 23 MuEmo-R1强化学习阶段的训练奖励变化',
    '_Ref29575': '表 32 不同方法在DFEW与MAFW上的情感识别结果',
    '_Ref29598': '表 33 不同方法在RAVDESS上的跨语料情感识别结果',
    '_Ref29713': '表 34 MuEmo-R1在DFEW上的消融实验结果',
    '_Ref29784': '表 35 模型与工具提取物理特征的一致性结果',
}

n_cap = {k: sum(1 for r in rows if r['kind'] == k) for k in ('图', '表')}
n_bad = {k: sum(1 for r in rows if r['kind'] == k and r['cls'] != 'OK') for k in ('图', '表')}
stale = [r for r in refs if r['bi'] not in {b['bi'] for b in broken}]
# 显示号 vs 题注号不一致的（复算）
bad_ref = []
bmi = {}
for r in rows:
    for b in r['refbm']:
        bmi[b] = r
for r in refs:
    cap = bmi.get(r['tgt'])
    if r['bi'] in {b['bi'] for b in broken} or not cap:
        continue
    if cap['visnum'] and str(cap['visnum']) != str(r['disp']).replace('图 ', '').replace('表 ', ''):
        bad_ref.append((r, cap))


def cap_table(kind):
    rr = [r for r in rows if r['kind'] == kind]
    out = ['<table><thead><tr><th>序</th><th>块</th><th>所属小节</th><th>状态</th>'
           '<th>可见号</th><th>编号域缓存</th><th>按F9后会变成</th><th>可见文本</th></tr></thead><tbody>']
    for i, r in enumerate(rr, 1):
        lab, bg, fg = CLS[r['cls']]
        warn = ''
        if r['cls'] != 'OK':
            warn = ' style="background:%s22"' % bg
        out.append('<tr%s>' % warn)
        out.append('<td class="n">%d</td><td class="n">%d</td>' % (i, r['bi']))
        out.append('<td class="h">%s</td>' % E(r['head'][:22]))
        out.append('<td><span class="tag" style="background:%s;color:%s">%s</span></td>'
                   % (bg, fg, lab))
        out.append('<td class="n b">%s</td>' % (r['visnum'] or '—'))
        out.append('<td class="n">%s</td>' % (r['cache'] or '—'))
        out.append('<td class="n">%s</td>' % (r['after_f9'] or '—'))
        out.append('<td>%s</td></tr>' % E(r['vis'][:64]))
        if r['cls'] == 'LEFTOVER':
            out.append('<tr%s><td></td><td colspan="7" class="fix">结构是 '
                       '<code>%s</code> + 【SEQ %s = %s】 + <code>%s</code>'
                       ' → 域后面还留着原来的数字</td></tr>'
                       % (warn, E(r['pre']), r['kind'], r['cache'], E(r['suf'][:40])))
    out.append('</tbody></table>')
    return ''.join(out)


CSS = """
body{background:#16181d;color:#e4e6ea;font-family:"Microsoft YaHei",Segoe UI,sans-serif;
     margin:0;padding:24px 30px;font-size:13.5px;line-height:1.6}
h1{font-size:20px;margin:0 0 3px}
h2{font-size:16px;margin:30px 0 10px;color:#7fd4ff;border-bottom:1px solid #2c3037;padding-bottom:6px}
p.sub{color:#98a0ab;font-size:12.5px;margin:0 0 16px}
.cards{display:flex;gap:12px;flex-wrap:wrap;margin:14px 0 6px}
.card{flex:1;min-width:150px;background:#1d2027;border:1px solid #2c3037;border-radius:7px;padding:11px 14px}
.card .k{color:#98a0ab;font-size:12px}
.card .v{font-size:23px;font-weight:700;margin-top:2px}
.card.bad .v{color:#ff8f8f}
.card.ok .v{color:#78e0a0}
.card.warn .v{color:#ffd479}
table{border-collapse:collapse;width:100%;font-size:12.5px;margin-top:6px}
th,td{border:1px solid #2c3037;padding:5px 8px;text-align:left;vertical-align:top}
th{background:#1d2027;color:#9fb4c8;font-weight:600;white-space:nowrap}
td.n{text-align:center;white-space:nowrap;color:#c6ccd4}
td.h{color:#8f97a2;font-size:11.5px;white-space:nowrap}
td.b{font-weight:700}
td.fix{background:#241f16;color:#e7c98a;font-size:12px}
code{background:#24272e;padding:1px 4px;border-radius:3px;color:#ffd479;font-size:11.5px}
.tag{display:inline-block;border-radius:3px;padding:1px 7px;font-size:11.5px;white-space:nowrap}
.note{background:#1d2027;border-left:3px solid #4a90d9;padding:10px 14px;border-radius:4px;margin:12px 0}
.note.warn{border-left-color:#c9863b;background:#241f16}
.note.bad{border-left-color:#c0504d;background:#251a1a}
ol,ul{margin:6px 0 6px 22px;padding:0}
li{margin:3px 0}
.broken{color:#98a0ab;font-size:12px}
"""

h = ['<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">',
     '<title>9-28 图表题注与交叉引用审计</title><style>%s</style></head><body>' % CSS,
     '<h1>中期科技报告 9-28 · 图表题注 与 交叉引用 审计</h1>',
     '<p class="sub">对象：<code>中期科技报告9-28-公式规范化.docx</code>　'
     '全文 %d 个顶层块（段落 %d / 表格 %d）</p>' % (d['blocks'], d['blocks'] - d['tables'], d['tables']),

     '<div class="cards">',
     '<div class="card ok"><div class="k">图 题注</div><div class="v">%d</div>'
     '<div class="k">其中 %d 个有问题</div></div>' % (n_cap['图'], n_bad['图']),
     '<div class="card ok"><div class="k">表 题注</div><div class="v">%d</div>'
     '<div class="k">其中 %d 个有问题</div></div>' % (n_cap['表'], n_bad['表']),
     '<div class="card bad"><div class="k">交叉引用断链</div><div class="v">%d</div>'
     '<div class="k">处（%d 个不同目标）</div></div>'
     % (len(broken), len({b['tgt'] for b in broken})),
     '<div class="card warn"><div class="k">引用显示号 ≠ 题注号</div><div class="v">%d</div>'
     '<div class="k">处</div></div>' % len(bad_ref),
     '<div class="card warn"><div class="k">正文硬写的图表号</div><div class="v">%d</div>'
     '<div class="k">处，改号时不会同步</div></div>' % len(hard),
     '</div>',

     '<div class="note bad"><b>最要紧的一条：</b>现在千万不要在 Word 里按 '
     '<code>Ctrl+A → F9</code> 更新域。全文 37 个 <code>SEQ 图</code> 域按文档顺序重排后会变成 1–37，'
     '而其中 2 个题注的域后面还残留着原编号（<code>图 </code>+【域=17】+<code>9 …</code>），'
     '按 F9 后会变成「<b>图 179</b>」「<b>图 180</b>」；剩下 11 个手打编号的题注又完全不参与重排，'
     '结果会彻底乱掉。</div>',

     '<h2>一、图 题注（%d 个，文档顺序）</h2>' % n_cap['图'],
     '<p class="sub">「按F9后会变成」列 = 该 SEQ 域在全文 SEQ 图 域中的序号，也就是更新域之后的显示值。'
     '可见号与它不同，说明现在的号是<b>假</b>的。</p>', cap_table('图'),

     '<h2>二、表 题注（%d 个，文档顺序）</h2>' % n_cap['表'], cap_table('表'),

     '<h2>三、交叉引用问题</h2>',
     '<h3 style="font-size:14px;color:#ff9b9b;margin:16px 0 6px">'
     '3.1 断链：%d 处引用指向不存在的书签</h3>' % len(broken),
     '<p class="sub">这些 REF 域在文档里找不到对应 bookmark，域缓存还留着旧号。'
     '按上下文可确定它们本该指向哪个题注（下表「建议重新指向」）。</p>',
     '<table><thead><tr><th>块</th><th>REF 目标</th><th>当前显示</th><th>建议重新指向</th></tr></thead><tbody>']
for b in broken:
    h.append('<tr><td class="n">%d</td><td><code>%s</code></td><td class="n">%s</td>'
             '<td class="fix">%s</td></tr>'
             % (b['bi'], b['tgt'], E(b['disp']), E(SUGGEST.get(b['tgt'], '（需人工确认）'))))
h.append('</tbody></table>')

h.append('<h3 style="font-size:14px;color:#ffd479;margin:18px 0 6px">'
         '3.2 号不符：%d 处引用的显示号与它指向的题注号不一致</h3>' % len(bad_ref))
h.append('<p class="sub">书签存在、能指向，但正文显示的号是旧的。'
         '更新域后会自愈——前提是题注本身的编号先修好。</p>')
h.append('<table><thead><tr><th>块</th><th>REF 目标</th><th>引用显示</th>'
         '<th>实际指向的题注</th></tr></thead><tbody>')
for r, cap in bad_ref:
    h.append('<tr><td class="n">%d</td><td><code>%s</code></td><td class="n">%s</td>'
             '<td>%s</td></tr>' % (r['bi'], r['tgt'], E(r['disp']), E(cap['vis'][:52])))
h.append('</tbody></table>')

h.append('<h2>四、正文里「硬写」的图表号（%d 处）</h2>' % len(hard))
h.append('<p class="sub">这些是直接敲在正文里的「如图 17 所示」「表 24」，'
         '<b>不是交叉引用</b>，改图号时必须一起改，否则会与题注脱节。</p>')
h.append('<table><thead><tr><th>块</th><th>写法</th><th>上下文</th></tr></thead><tbody>')
for x in hard:
    h.append('<tr><td class="n">%d</td><td class="n b">%s</td><td class="broken">%s</td></tr>'
             % (x['bi'], E(x['n']), E(x['ctx'])))
h.append('</tbody></table>')

h.append('<h2>五、建议的修复顺序</h2><ol>'
         '<li><b>重建题注编号</b>：44 个图题注、52 个表题注统一成 '
         '<code>图 N</code> + SEQ 域 + 标题的结构，去掉域后残留数字，'
         '给每个题注补 <code>_Ref</code> 书签。</li>'
         '<li><b>同步正文 37 处硬写号</b>。</li>'
         '<li><b>修交叉引用</b>：14 处断链按上表重新指向；其余显示号更新域后自动对齐。</li>'
         '<li>最后再 <code>Ctrl+A → F9</code> 全局更新一次目录与域。</li></ol>'
         '<div class="note">注意：题注编号方案（全文连续 <code>图1–44</code> / <code>表1–52</code>，'
         '还是按章 <code>图x-y</code>）需要你定一下再动，因为它会同时改写正文里那 37 处引用。</div>'
         '</body></html>')

p = os.path.join(WORK, '图表题注与交叉引用审计.html')
open(p, 'w', encoding='utf-8').write('\n'.join(h))
print('已写出', p, os.path.getsize(p) // 1024, 'KB')
