# -*- coding: utf-8 -*-
"""检查每个编号段：制表位 / 公式前 Tab / 编号前 Tab 三件套是否齐全。"""
import sys, zipfile, re

TAGP = re.compile(r'<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>|</w:p>')
PPR_SCAN = re.compile(r'</?w:pPr>')
NUMEND = re.compile(r'[（(]\s*\d+(?:\s*[-–—]\s*\d+)?\s*[)）]\s*$')

def paras(xml):
    segs = []; depth = 0; start = None
    for mo in TAGP.finditer(xml):
        t = mo.group(0)
        if t.startswith('</'):
            depth -= 1
            if depth == 0:
                segs.append((start, mo.end()))
        elif t.endswith('/>'):
            if depth == 0:
                segs.append((mo.start(), mo.end()))
        else:
            if depth == 0:
                start = mo.start()
            depth += 1
    return segs

def ppr_span(seg):
    s = seg.find("<w:pPr>")
    if s < 0:
        return None
    depth = 0
    for mo in PPR_SCAN.finditer(seg, s):
        if mo.group(0) == "<w:pPr>":
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                return (s, mo.end())
    return None

def report(path, label):
    with zipfile.ZipFile(path) as z:
        doc = z.read('word/document.xml').decode('utf-8')
    sp = paras(doc)
    good = []; bad = []
    for i, (s, e) in enumerate(sp):
        seg = doc[s:e]
        if '<m:oMath' not in seg:
            continue
        txt = re.sub(r'<[^>]+>', '', seg).strip()
        if not NUMEND.search(txt):
            continue
        spn = ppr_span(seg)
        ppr = seg[spn[0]:spn[1]] if spn else ''
        has_tabs = '<w:tabs>' in ppr
        k2 = seg.find('<m:oMath')
        between = seg[spn[1]:k2] if (spn and 0 <= spn[1] < k2) else ''
        pre_tab = '<w:tab/>' in between
        cut = seg.rfind('</m:oMath>')
        tail = seg[cut:] if cut >= 0 else ''
        post_tab = '<w:tab/>' in tail
        mjc = re.search(r'<w:jc\s+w:val="([^"]+)"', ppr)
        rec = (i, has_tabs, pre_tab, post_tab, mjc.group(1) if mjc else '-')
        (good if (has_tabs and pre_tab and post_tab) else bad).append(rec)
    print('=== %s ===' % label)
    print('  编号段总数 %d ｜ 三件套齐全 %d ｜ 不合规 %d' % (len(good) + len(bad), len(good), len(bad)))
    for r in bad:
        print('    ✗ 段%-5d tabs=%-5s 公式前Tab=%-5s 编号前Tab=%-5s jc=%s' % r)
    return len(bad)

if __name__ == '__main__':
    n = report(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else 'doc')
    sys.exit(0 if n == 0 else 1)
