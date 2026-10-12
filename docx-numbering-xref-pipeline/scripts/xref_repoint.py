# -*- coding: utf-8 -*-
r"""REF 交叉引用「指向修复」执行器 —— 一次完成三类修复（只改 word/document.xml）。

为什么单独一个脚本（与 siblings 的分工）：
    · `add_ref_bookmark.py`：题注**从未有过** _Ref 书签（正文只好硬写）→ **自动编号**补书签。
    · 本脚本：正文**已经有** REF 域，只是域里的书签名指向了不存在的书签 / 空书签 / 过宽书签。
      修复的是「域 → 书签」的指向关系，**不新增、不删除任何可见文字**。

三类动作（由一个 plan.json 驱动，可任选）：
  1) repoint   改 REF 域指令里的书签名（断链重指向）。只改 <w:instrText>，不动结果区缓存。
  2) replenish 给题注段补书签，**沿用正文引用的原名** —— 这样正文一侧零改动。
               （注意：别用自动编号，否则新书签与旧引用对不上，断链照旧。）
  3) narrow    收窄「过宽」书签：把 bookmarkEnd 挪到 SEQ 域 end 之后。
               过宽书签会让正文引用直接渲染成病句（如「如图 1-34 脚本缺陷问题示例图所示」）。

plan.json 形状：
{
  "check_md5": "207d…",                                   // 可选
  "repoint":   [{"from": "_Ref1244", "to": "_Ref9009"}],
  "replenish": [{"name": "_Ref1862923300", "anchor": "图 2-8"}],
  "narrow":    [{"name": "_Ref12095"}]
}

用法: python xref_repoint.py <in.docx> <out.docx> <plan.json> [--dry]
退出码: 0 成功 / 2 自检失败 / 5 指纹不符 / 6 用法错误
"""
import os
import re
import sys
import json
import hashlib
import zipfile

import xml.etree.ElementTree as ET

DOCXML = 'word/document.xml'
TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
WT = re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', re.S)
INSTR = re.compile(r'<w:instrText[^>]*>(.*?)</w:instrText>', re.S)
FCEND_RUN = '<w:fldChar w:fldCharType="end"/></w:r>'
BMSTART = re.compile(r'<w:bookmarkStart\b[^>]*?/>')
BMEND = re.compile(r'<w:bookmarkEnd\b[^>]*?/>')
ZERO = re.compile(r'[\u200b-\u200f\ufeff\u2060]')


def paras(x):
    res, depth, st = [], 0, None
    for m in TAGP.finditer(x):
        s = m.group(0)
        if s.endswith('/>'):
            if depth == 0:
                res.append((m.start(), m.end()))
        elif s == '</w:p>':
            depth -= 1
            if depth <= 0:
                if st is not None:
                    res.append((st, m.end()))
                st, depth = None, 0
        else:
            if depth == 0:
                st = m.start()
            depth += 1
    return res


def bm_index(x):
    st, en = {}, {}
    for m in BMSTART.finditer(x):
        t = m.group(0)
        i = re.search(r'w:id="(\d+)"', t)
        n = re.search(r'w:name="([^"]*)"', t)
        if i and n:
            st[i.group(1)] = (n.group(1), m.start(), m.end())
    for m in BMEND.finditer(x):
        i = re.search(r'w:id="(\d+)"', m.group(0))
        if i:
            en[i.group(1)] = (m.start(), m.end())
    return st, en


def domains(x):
    TOK = re.compile(r'<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"[^>]*/>'
                     r'|<w:instrText\b[^>]*>(.*?)</w:instrText>')
    stk, out = [], []
    for m in TOK.finditer(x):
        if m.group(1) is not None:
            k = m.group(1)
            if k == 'begin':
                stk.append({'i': '', 'n': 0})
            elif k == 'separate':
                if stk:
                    stk[-1]['n'] += 1
            else:
                if stk:
                    out.append(stk.pop())
        else:
            if stk:
                stk[-1]['i'] += m.group(2)
    return out, len(stk)


def ppr_end(q):
    i = q.find('<w:pPr>')
    if i < 0:
        return 0
    d = 0
    for m in re.finditer(r'<w:pPr>|</w:pPr>', q):
        if m.start() < i:
            continue
        d += 1 if m.group(0) == '<w:pPr>' else -1
        if d == 0:
            return m.end()
    raise AssertionError('pPr 未配平')


def main():
    argv = [a for a in sys.argv[1:] if not a.startswith('--')]
    dry = '--dry' in sys.argv
    if len(argv) < 3:
        print(__doc__)
        return 6
    SRC, DST, PLANF = argv[0], argv[1], argv[2]

    plan = json.load(open(PLANF, encoding='utf-8'))
    if plan.get('check_md5'):
        got = hashlib.md5(open(SRC, 'rb').read()).hexdigest()
        if got != plan['check_md5']:
            print('!! 指纹不符：期望 %s 实际 %s' % (plan['check_md5'], got))
            return 5
        print('[OK] 指纹校验通过')

    zin = zipfile.ZipFile(SRC)
    infos = zin.infolist()
    x = zin.read(DOCXML).decode('utf-8')

    def cnt(p, s):
        return len(re.findall(p, s))

    BASE = {'w:p': cnt(r'<w:p[ >]', x) + cnt(r'<w:p/>', x), 'w:tbl': cnt(r'<w:tbl[ >]', x),
            'w:drawing': cnt(r'<w:drawing[ >]', x), 'fldChar': cnt(r'<w:fldChar\b', x),
            'bmStart': cnt(r'<w:bookmarkStart\b', x), 'bmEnd': cnt(r'<w:bookmarkEnd\b', x)}
    print('基线:', BASE)

    rep = {'repoint': 0, 'replenish': 0, 'narrow': 0}

    # ---------------- 1) repoint ----------------
    for spec in plan.get('repoint', []):
        frm, to = spec['from'], spec['to']
        pat = re.compile(r'(<w:instrText[^>]*>\s*REF\s+)' + re.escape(frm) + r'(?=\s)')
        x, n = pat.subn(lambda m: m.group(1) + to, x)
        want = spec.get('count')
        if want is not None and n != want:
            print('!! repoint %s->%s 命中 %d，期望 %d' % (frm, to, n, want))
            return 2
        if n == 0:
            print('!! repoint %s->%s 一处未命中' % (frm, to))
            return 2
        rep['repoint'] += n
        print('  repoint %-16s -> %-12s x%d' % (frm, to, n))

    # ---------------- 2) replenish ----------------
    if plan.get('replenish'):
        st, en = bm_index(x)
        used = set(k for k in list(st) + list(en))
        names = set(v[0] for v in st.values())
        nid = max([int(k) for k in used] or [-1]) + 1
        P = paras(x)
        for spec in plan['replenish']:
            name, anchor = spec['name'], spec['anchor']
            if name in names:
                print('  replenish %s 已存在，跳过' % name)
                continue
            hit = None
            for a, b in P:
                q = x[a:b]
                if re.search(r'<w:bookmarkStart\b[^>]*w:name="_Ref', q):
                    continue
                ins = ' '.join(INSTR.findall(q))
                if not re.search(r'SEQ\s+[图表]', ins):
                    continue
                txt = ZERO.sub('', ''.join(WT.findall(q))).strip()
                if txt.startswith(anchor):
                    hit = (a, b, q)
                    break
            if hit is None:
                print('!! replenish 找不到锚点 %r 的题注段' % anchor)
                return 2
            a, b, q = hit
            q2 = q[:ppr_end(q)] + '<w:bookmarkStart w:id="%d" w:name="%s"/>' % (nid, name) + q[ppr_end(q):]
            idx = q2.rfind(FCEND_RUN)
            if idx < 0:
                print('!! replenish %s 段内找不到 SEQ 域 end' % name)
                return 2
            p = idx + len(FCEND_RUN)
            q2 = q2[:p] + '<w:bookmarkEnd w:id="%d"/>' % nid + q2[p:]
            x = x[:a] + q2 + x[b:]
            rep['replenish'] += 1
            print('  replenish %-18s @anchor=%r  id=%d' % (name, anchor, nid))
            nid += 1

    # ---------------- 3) narrow ----------------
    if plan.get('narrow'):
        for spec in plan['narrow']:
            name = spec['name']
            st, en = bm_index(x)
            sbid = next((k for k, v in st.items() if v[0] == name), None)
            if sbid is None or sbid not in en:
                print('!! narrow %s 不存在或不配对' % name)
                return 2
            s0 = st[sbid][1]
            cur_e0, cur_e1 = en[sbid]
            P = paras(x)
            a, b, q = None, None, None
            for pa, pb in P:
                if pa <= s0 < pb:
                    a, b, q = pa, pb, x[pa:pb]
                    break
            if a is None:
                print('!! narrow %s 定位不到所在段' % name)
                return 2
            # 段内最后一个 SEQ 域 end run 的结束位置
            idx = q.rfind(FCEND_RUN)
            if idx < 0:
                print('!! narrow %s 段内无 SEQ 域 end' % name)
                return 2
            want_pos = a + idx + len(FCEND_RUN)
            closing = '<w:bookmarkEnd w:id="%s"/>' % sbid
            if cur_e0 == want_pos:
                print('  narrow %-18s 已在正确位置，跳过' % name)
                continue
            q2 = q.replace(closing, '')
            idx2 = q2.rfind(FCEND_RUN)
            p = idx2 + len(FCEND_RUN)
            q2 = q2[:p] + closing + q2[p:]
            x = x[:a] + q2 + x[b:]
            rep['narrow'] += 1
            print('  narrow %-18s bookmarkEnd 由段末移至 SEQ 域 end 之后' % name)

    # ---------------- 不变量 ----------------
    NOW = {'w:p': cnt(r'<w:p[ >]', x) + cnt(r'<w:p/>', x), 'w:tbl': cnt(r'<w:tbl[ >]', x),
           'w:drawing': cnt(r'<w:drawing[ >]', x), 'fldChar': cnt(r'<w:fldChar\b', x),
           'bmStart': cnt(r'<w:bookmarkStart\b', x), 'bmEnd': cnt(r'<w:bookmarkEnd\b', x)}
    grown = rep['replenish']          # narrow/repoint 不改计数
    for k in BASE:
        exp = BASE[k] + (grown if k in ('bmStart', 'bmEnd') else 0)
        if NOW[k] != exp:
            print('!! 不变量破坏 %s: %d -> %d (期望 %d)' % (k, BASE[k], NOW[k], exp))
            return 2
        print('  %-9s %d -> %d OK' % (k, BASE[k], NOW[k]))
    if NOW['bmStart'] != NOW['bmEnd']:
        print('!! 书签不配平')
        return 2

    # ---------------- 断链自检 ----------------
    doms, unclosed = domains(x)
    st, en = bm_index(x)
    NAMES = set(v[0] for v in st.values())
    bad = [d['i'].strip()[:24] for d in doms
           if re.match(r'\s*REF\s+(\S+)', d['i'])
           and re.match(r'\s*REF\s+(\S+)', d['i']).group(1) not in NAMES]
    print('  域栈闭合=%s  REF 断链=%d %s' % (unclosed == 0, len(bad), bad[:5]))
    if unclosed or bad:
        print('!! 断链未清零，不落盘')
        return 2

    ET.fromstring(x.encode('utf-8'))
    if dry:
        print('[dry] 汇总 %r；未写文件' % rep)
        return 0

    xml_bytes = x.encode('utf-8')
    zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
    for it in infos:
        data = zin.read(it.filename)
        if it.filename == DOCXML:
            data = xml_bytes
        ni = zipfile.ZipInfo(it.filename, date_time=it.date_time)
        ni.compress_type = it.compress_type
        ni.external_attr = it.external_attr
        ni.internal_attr = it.internal_attr
        ni.create_system = it.create_system
        ni.flag_bits = it.flag_bits
        zout.writestr(ni, data)
    zout.close()
    zin.close()
    print('已写出 %s (%d bytes)  汇总=%r' % (DST, os.path.getsize(DST), rep))
    return 0


if __name__ == '__main__':
    sys.exit(main())
