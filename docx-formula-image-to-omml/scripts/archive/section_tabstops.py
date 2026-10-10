# -*- coding: utf-8 -*-
"""（历史脚本，已归档；路径参数须由环境变量或显式传参提供）
把关键段落映射到所属分节，算出该节正确的制表位。
"""
import os
import zipfile, re

SRC = os.environ.get('DOCXFMT_SRC')
xml = zipfile.ZipFile(SRC).read('word/document.xml').decode('utf-8')
PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)

# 段落在 xml 中的起点偏移
offs, pos = [], 0
for m in PARA.finditer(xml):
    offs.append(m.start())
ps = [m.group(0) for m in PARA.finditer(xml)]

# 所有 sectPr 的位置
sects = []
for m in re.finditer(r'<w:sectPr(?: [^>]*)?>.*?</w:sectPr>', xml, re.S):
    s = m.group(0)
    pg = re.search(r'<w:pgSz[^/]*w:w="(\d+)"', s)
    mg_l = re.search(r'w:left="(\d+)"', s)
    mg_r = re.search(r'w:right="(\d+)"', s)
    land = 'landscape' in s
    W, L, R = int(pg.group(1)), int(mg_l.group(1)), int(mg_r.group(1))
    right = W - L - R
    sects.append(dict(pos=m.start(), W=W, land=land, right=right, center=right // 2,
                      raw='%d/%d' % (right // 2, right)))


def sec_of(pi):
    p = offs[pi]
    for s in sects:
        if s['pos'] > p:
            return s
    return sects[-1]


print('共 %d 个 sectPr' % len(sects))
for k, s in enumerate(sects, 1):
    print('  节%d pos=%-9d W=%-6d %-10s 制表位 center=%d right=%d' %
          (k, s['pos'], s['W'], '横向' if s['land'] else '纵向', s['center'], s['right']))

targets = {
    '第1章行间公式': [224, 226, 228, 229, 231, 232, 242, 245, 248, 878, 880, 883, 885, 888, 1712, 1717],
    '§3.4.2 oMathPara': [4417, 4419, 4424, 4426, 4428, 4431, 4433],
}
# 所有带编号的 ch2/3 公式段
num_paras = []
for i, q in enumerate(ps):
    if i < 1720:
        continue
    t = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S)))
    if re.fullmatch(r'\(\d+\)', t.strip()):
        num_paras.append(i)
print('\n带编号的 ch2/3 公式段: %d 个' % len(num_paras))
targets['ch2/3 已编号公式'] = num_paras

for name, idxs in targets.items():
    print('\n=== %s ===' % name)
    bad = 0
    for i in idxs:
        s = sec_of(i)
        p = ps[i]
        tb = re.findall(r'<w:tab w:val="(\w+)" w:pos="(\d+)"', p)
        cur = ' '.join('%s:%s' % t for t in tb) or '<无制表位>'
        flag = ''
        if tb and any(pos != str(s['center']) and pos != str(s['right']) for _, pos in tb):
            flag = '  ⚠ 与所属节不匹配'
            bad += 1
        print('  para %-5d 节W=%-6d %-4s 现制表位 %-26s 应为 %s%s' %
              (i, s['W'], '横' if s['land'] else '纵', cur, s['raw'], flag))
    if name == 'ch2/3 已编号公式':
        print('  … 不匹配合计: %d' % bad)
