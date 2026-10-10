# -*- coding: utf-8 -*-
r"""按 Word 的 REF 语义重算每个 REF 域应显示什么，与文件里的域缓存比对。

Word 语义： `REF <bookmark> \h`  →  显示 bookmark 覆盖范围内的文本（内嵌域用其缓存值渲染）。
所以「缓存 == 目标题注号」只是必要不充分条件：若 bookmark 包住了整段题注，
F9 之后引用会膨胀成「图 5 标题全文」。本脚本专门抓这类隐性坑。

用法: python audit_ref_semantics.py <docx> [<docx2> ...]
"""
import zipfile, re, sys, os

TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
T_BEGIN = '<w:fldChar w:fldCharType="begin"/>'
T_SEP = '<w:fldChar w:fldCharType="separate"/>'
T_END = '<w:fldChar w:fldCharType="end"/>'


def paras(x):
    """深度感知切分 <w:p>（含表格单元格内段落）"""
    res, depth, start = [], 0, None
    for m in TAGP.finditer(x):
        s = m.group(0)
        if s.endswith('/>'):
            if depth == 0:
                res.append((m.start(), m.end()))
        elif s == '</w:p>':
            depth -= 1
            if depth <= 0:
                if start is not None:
                    res.append((start, m.end()))
                start = None
                depth = 0
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return res


def render(frag):
    """渲染可见文本：域取 separate..end 之间的缓存；无 separate 的域整体跳过"""
    out, pos = [], 0
    while True:
        a = frag.find(T_BEGIN, pos)
        if a < 0:
            out.append(frag[pos:]); break
        sm = frag.find(T_SEP, a)
        em = frag.find(T_END, a)
        if sm < 0 or em < 0 or sm > em:
            out.append(frag[pos:a])
            if em < 0:
                out.append(frag[a:]); break
            out.append(frag[a:em])
            pos = em
            continue
        out.append(frag[pos:a])
        out.append(frag[sm + len(T_SEP):em])
        pos = em + len(T_END)
    return ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', ''.join(out), re.S))


def bookmark_spans(x):
    st, en = {}, {}
    for m in re.finditer(r'<w:bookmarkStart\b[^>]*?/>', x):
        t = m.group(0)
        i = re.search(r'w:id="(\d+)"', t)
        n = re.search(r'w:name="([^"]*)"', t)
        if i and n:
            st[i.group(1)] = (n.group(1), m.end())
    for m in re.finditer(r'<w:bookmarkEnd\b[^>]*?/>', x):
        i = re.search(r'w:id="(\d+)"', m.group(0))
        if i:
            en[i.group(1)] = m.start()
    return {n: (s, en[k]) for k, (n, s) in st.items() if k in en}


def audit(path):
    z = zipfile.ZipFile(path)
    x = z.read('word/document.xml').decode('utf-8')
    z.close()
    sp = bookmark_spans(x)
    rows, wide = [], []
    for i, (a, b) in enumerate(paras(x)):
        q = x[a:b]
        for m in re.finditer(T_BEGIN, q):
            seg = q[m.end():]
            em = re.search(T_END, seg)
            if not em:
                continue
            body = seg[:em.start()]
            im = re.search(r'<w:instrText[^>]*>(.*?)</w:instrText>', body, re.S)
            if not im:
                continue
            instr = im.group(1).strip()
            if not instr.upper().startswith('REF'):
                continue
            tm = re.search(r'REF\s+(\S+)', instr)
            if not tm:
                continue
            tgt = tm.group(1)
            sm = body.find(T_SEP)
            cache = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>',
                                       body[sm:], re.S)).strip() if sm >= 0 else ''
            exp = render(x[sp[tgt][0]:sp[tgt][1]]).strip() if tgt in sp else None
            rows.append(dict(pi=i, tgt=tgt, cache=cache, exp=exp))
            if tgt in sp:
                inner = x[sp[tgt][0]:sp[tgt][1]]
                if 'SEQ' in inner and re.search(r'<w:bookmarkStart\b[^>]*w:name="_Toc', inner):
                    wide.append(tgt)

    bad = [r for r in rows if r['exp'] is not None and r['cache'] != r['exp']]
    gone = [r for r in rows if r['exp'] is None]
    print('=' * 78)
    print('%s' % path)
    print('=' * 78)
    print('  REF 域: %d   目标书签缺失: %d   缓存与语义不符: %d   过宽书签(被引用): %d'
          % (len(rows), len(gone), len(bad), len(set(wide))))
    for r in gone:
        print('     p%-5d %-13s 书签不存在' % (r['pi'], r['tgt']))
    for r in bad:
        print('     p%-5d %-13s\n        缓存 = %r\n        应为 = %r'
              % (r['pi'], r['tgt'], r['cache'], r['exp']))
    return len(bad) + len(gone)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    total = 0
    for p in sys.argv[1:]:
        print()
        total += audit(p)
    print('\n合计问题: %d' % total)
    sys.exit(0 if total == 0 else 2)
