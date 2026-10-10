# -*- coding: utf-8 -*-
"""精确列出 docx 里 图/表 SEQ 域位置，定位重复与样式异常。

被编排调用时（docx-unify-pipeline 的 S2）按两级回退取入参，**缺失则报错**：
    DOC = ① sys.argv[1]  ② 环境变量 DOCXFMT_DOC
"""
import os
import sys
import zipfile, re

PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')

DOC = sys.argv[1] if len(sys.argv) > 1 else os.environ.get('DOCXFMT_DOC')
if not DOC:
    sys.exit('用法: locate_dup.py <in.docx>（或设 DOCXFMT_DOC）')

X = zipfile.ZipFile(DOC).read('word/document.xml').decode('utf-8')
ps = PARA.findall(X)

for kind in ('图', '表'):
    rows = []
    for i, q in enumerate(ps):
        if 'SEQ %s' % kind not in q:
            continue
        st = re.search(r'<w:pStyle w:val="([^"]+)"', q)
        segs = re.findall(r'separate"/>(.*?)<w:fldChar w:fldCharType="end"/>', q, re.S)
        num = ''
        for sg in segs:
            n = ''.join(T.findall(sg)).strip()
            if re.fullmatch(r'\d+', n):
                num = n
        has_img = '<w:drawing' in q
        rows.append((i, st.group(1) if st else '-', num, has_img,
                     ''.join(T.findall(q)).strip()))
    print('\n' + '=' * 78)
    print('SEQ %s：共 %d 条' % (kind, len(rows)))
    print('=' * 78)
    prev = None
    for k, (i, sty, num, img, txt) in enumerate(rows, 1):
        flags = []
        if sty != '12':
            flags.append('样式=%s' % sty)
        if num == prev:
            flags.append('★与上条同号')
        elif prev and num.isdigit() and prev.isdigit() and int(num) != int(prev) + 1:
            flags.append('跳号')
        prev = num
        if flags:
            print('  [%2d] para %-5d 号=%-4s %s' % (k, i, num, ' | '.join(flags)))
            print('       文本: %s' % txt[:70])
