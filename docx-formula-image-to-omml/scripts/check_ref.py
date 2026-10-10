# -*- coding: utf-8 -*-
"""核对每个交叉引用(REF 域)与它指向题注的显示号是否一致。"""
import zipfile, re

PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')


def run(path, tag):
    X = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8')
    ps = PARA.findall(X)

    # 书签 -> 所在段落
    bm2para = {}
    for i, q in enumerate(ps):
        for n in re.findall(r'<w:bookmarkStart[^>]*w:name="([^"]+)"', q):
            bm2para[n] = i

    # 每个 REF 域：目标书签 + 缓存显示值
    refs = []
    for i, q in enumerate(ps):
        for m in re.finditer(
                r'<w:fldChar w:fldCharType="begin"/>.*?<w:instrText[^>]*>\s*REF\s+(_Ref\w+)[^<]*</w:instrText>'
                r'.*?<w:fldChar w:fldCharType="separate"/>(.*?)<w:fldChar w:fldCharType="end"/>',
                q, re.S):
            target = m.group(1)
            cached = ''.join(T.findall(m.group(2))).strip()
            refs.append((i, target, cached))

    print('\n' + '=' * 78)
    print('%s   交叉引用(REF) 共 %d 个' % (tag, len(refs)))
    print('=' * 78)
    bad = 0
    for i, target, cached in refs:
        pi = bm2para.get(target)
        cap = ''.join(T.findall(ps[pi])).strip() if pi is not None else '<无书签>'
        capnum = re.search(r'[图表]\s*(\d+)', cap)
        capnum = capnum.group(1) if capnum else '?'
        bad_mark = ''
        if str(capnum) not in cached:
            bad_mark = '  ❌显示号与题注号不一致'
            bad += 1
        print('  para %-5d -> %-11s 引用显示=%-4s 题注=%-4s %s%s' %
              (i, target, cached, capnum, cap[:34], bad_mark))
    print('\n  不一致条数: %d' % bad)


run(r'C:\Users\14711\Desktop\中期科技报告(09241800)-9-26.backup-20260926-222152.docx', 'A 我动手前(备份 22:12)')
run(r'C:\Users\14711\Desktop\中期科技报告(09241800)-9-26-公式规范化.docx', 'B 我的产出(22:24)')
run(r'C:\Users\14711\Desktop\中期科技报告(09241800)-9-26.docx', 'C 你的最新版(22:44)')
