# -*- coding: utf-8 -*-
r"""给「题注段内已有 SEQ 图 / SEQ 表 域、但缺 _Ref 书签」的题注段补一个 _Ref 书签。

背景（见 SKILL §15「交叉引用三合一」）：
    正文要用 REF 域指向某张图 / 表时，目标题注段必须有一个 `_Ref` 书签。
    本机这份稿子里有 3 个表题注只有 `_Toc` 书签、没有 `_Ref` 书签，于是正文
    无法用 REF 域指向它们（`build_caption_ref_map` 反查不到 → 硬写）。

    本脚本给这类题注段补一个 `_Ref` 书签，**只包住「图/表 N」+ SEQ 域**，
    绝不包住整段题注——否则 Word 全量更新域（F9 / 打印 / 导出 PDF）后，
    引用会膨胀成「图 5 标题全文」（与 `fix_ref_bookmark_span.py` 相反的方向）。

铁律（与 `fix_ref_bookmark_span.py` 的收窄判据对称）：
    · 新 `_Ref` 跨度的 XML 长度必须 **< 350**；
    · 跨度内部 **不得** 含 `_Toc` 开头书签。
    两者任一不满足 → **该段不出书签**；只要有任何一段不合法，**整体跳过不落盘并报警**
    （本脚本不做"把 bookmarkEnd 往前挪"的自动收窄，需人工收窄跨度后重跑）。

书签结构对齐文档内既有 `_Ref` 书签：
    `<w:bookmarkStart .../>` 是 `<w:p>` 的直接子元素（与 `<w:r>` 平级），
    落在「图 」/「表 」那个 run 之前；`<w:bookmarkEnd .../>` 落在 SEQ 域
    `<w:fldChar w:fldCharType="end"/>` 所在 run 之后。

命名：扫全文档既有 `_Ref\d+`，取当前最大编号 +1 起递增；`w:id` 取现有
      最大 bookmark id +1 起递增。

写盘：**只有 `word/document.xml` 变**；其余 zip 部件逐字节照抄，且复制
      `date_time` / `compress_type` / `external_attr` / `internal_attr` /
      `create_system` / `flag_bits`，保持条目顺序不变。

用法: python add_ref_bookmark.py <in.docx> <out.docx> [--dry]
"""
import os
import re
import sys
import zipfile

import xml.etree.ElementTree as ET

DRY = '--dry' in sys.argv
_argv = [a for a in sys.argv[1:] if not a.startswith('--')]
if len(_argv) < 2:
    print(__doc__)
    sys.exit(1)
SRC, DST = _argv[0], _argv[1]

DOCXML = 'word/document.xml'
TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
RUNRE = re.compile(r'<w:r(?: [^>]*)?>.*?</w:r>|<w:r(?: [^>]*)?/>', re.S)
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', re.S)
INSTR = re.compile(r'<w:instrText[^>]*>(.*?)</w:instrText>', re.S)
SEQ_INSTR = re.compile(r'SEQ\s+([图表])')
FBEGIN = '<w:fldChar w:fldCharType="begin"/>'
FEND = '<w:fldChar w:fldCharType="end"/>'
BMSTART = re.compile(r'<w:bookmarkStart\b[^>]*?/>')
BMEND = re.compile(r'<w:bookmarkEnd\b[^>]*?/>')
MAX_SPAN = 350           # 与 fix_ref_bookmark_span.py 的过宽判据同口径


# ------------------------------------------------------------------ 基础
def paras(x):
    """深度感知切分 <w:p>（含表格单元格内段落），返回 [(start, end), ...]。"""
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
                st = None
                depth = 0
        else:
            if depth == 0:
                st = m.start()
            depth += 1
    return res


def bm_index(x):
    """返回 (bookmarkStart 表, bookmarkEnd 表)，键都是 w:id 字符串。"""
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


def alloc_ids(x):
    """扫现有空间，给出 (下一个 bookmark id, 下一个 _Ref 编号)。"""
    st, en = bm_index(x)
    ids = [int(k) for k in list(st) + list(en)]
    next_id = (max(ids) + 1) if ids else 0
    nums = []
    for name, *_ in st.values():
        m = re.fullmatch(r'_Ref(\d+)', name)
        if m:
            nums.append(int(m.group(1)))
    next_ref = (max(nums) + 1) if nums else 1
    return next_id, next_ref


def run_text(runxml):
    """run 内可见文本（拼接所有 <w:t>）。"""
    return ''.join(WT.findall(runxml))


# ------------------------------------------------------------------ 找题注段
def candidates(x):
    """找「有 SEQ 图/表 域、但段内无 _Ref bookmarkStart」的题注段。

    返回 [(para_index, (a, b), kind, start_off, end_off, note), ...]
    start_off / end_off 是相对整篇 x 的插入点：
        bookmarkStart 插在 start_off；bookmarkEnd 插在 end_off。
    """
    out = []
    for pi, (a, b) in enumerate(paras(x)):
        q = x[a:b]
        # 段内已有 _Ref bookmarkStart → 不重复补
        if re.search(r'<w:bookmarkStart\b[^>]*w:name="_Ref', q):
            continue
        # 必须是 SEQ 图/表 域
        im = None
        for m in INSTR.finditer(q):
            if SEQ_INSTR.search(m.group(1)):
                im = m
                break
        if im is None:
            continue
        kind = SEQ_INSTR.search(im.group(1)).group(1)
        ipos = im.start()
        bm = None
        for m in re.finditer(re.escape(FBEGIN), q):
            if m.start() < ipos:
                bm = m            # 取最后一个在 instr 之前的 begin
            else:
                break
        if bm is None:
            continue
        em = q.find(FEND, im.end())
        if em < 0:
            continue
        bpos, epos = bm.start(), em

        # 找 run：begin 所在 run，以及它前一个「标签」run
        runs = list(RUNRE.finditer(q))
        rb = None
        for r in runs:
            if r.start() <= bpos < r.end():
                rb = r
                break
        if rb is None:
            continue
        label_off = rb.start()                    # 默认从 begin run 起
        prev = None
        for r in runs:
            if r.end() <= rb.start():
                prev = r
            else:
                break
        if prev is not None:
            txt = run_text(prev.group(0)).strip()
            if txt and txt[-1] in ('图', '表'):
                label_off = prev.start()          # 含「表 」/「图 」 → 从它起

        re_ = None
        for r in runs:
            if r.start() <= epos < r.end():
                re_ = r
                break
        if re_ is None:
            continue
        end_off = re_.end()                       # 落在 SEQ end run 之后

        out.append((pi, (a, b), kind, a + label_off, a + end_off, ''))
    return out


# ------------------------------------------------------------------ 主流程
def main():
    z = zipfile.ZipFile(SRC)
    x = z.read(DOCXML).decode('utf-8')
    infos = z.infolist()
    z.close()

    next_id, next_ref = alloc_ids(x)
    cands = candidates(x)

    print('候选题注段（有 SEQ 域、缺 _Ref 书签）: %d' % len(cands))
    print('起始编号:  bookmark id=%d   _Ref%d' % (next_id, next_ref))

    inserts = []          # (offset, kind, text)
    report = []           # (pi, name, bid, span_len, ok, note)
    new_bm = set()        # 新书签名，供落盘后自检

    for k, (pi, rng, kind, s_off, e_off, _note) in enumerate(cands):
        name = '_Ref%d' % (next_ref + k)
        bid = str(next_id + k)
        start_tag = '<w:bookmarkStart w:id="%s" w:name="%s"/>' % (bid, name)
        end_tag = '<w:bookmarkEnd w:id="%s"/>' % bid

        # 跨度 = bookmarkStart 与 bookmarkEnd 之间的内容（不含标签本身）；
        # 两标签都插在 run 边界，故其内容恰为 x[s_off:e_off]。
        span_txt = x[s_off:e_off]
        span_len = e_off - s_off
        has_toc = bool(re.search(r'<w:bookmarkStart\b[^>]*w:name="_Toc', span_txt))
        ok = (span_len < MAX_SPAN) and (not has_toc)
        note = ''
        if not ok:
            note = ('跨度 %d≥%d' % (span_len, MAX_SPAN)) if span_len >= MAX_SPAN \
                else '跨度内含 _Toc'
        inserts.append((s_off, 'S', start_tag))
        inserts.append((e_off, 'E', end_tag))
        new_bm.add(name)
        report.append((pi, name, bid, span_len, ok, note, kind))

    for r in report:
        pi, name, bid, span_len, ok, note, kind = r
        print('   para %-5d  %-13s id=%-5s %s 跨度=%d  %s'
              % (pi, name, bid, kind, span_len,
                 'OK' if ok else '!! ' + note))

    bad = [r for r in report if not r[4]]
    if bad:
        print('\n!! 有 %d 个书签跨度不合法，**整体跳过不落盘**（需人工）：'
              % len(bad))
        for r in bad:
            print('   para %d %s 跨度=%d %s' % (r[0], r[1], r[3], r[5]))
        if DRY:
            print('[dry] 未写文件')
            return 1
        return 1

    if DRY:
        print('\n[dry] 将插入 %d 个 _Ref 书签；未写文件。' % len(report))
        return 0

    # 应用插入：按 offset 降序，保证 offset 不失效
    for off, _k, txt in sorted(inserts, key=lambda t: t[0], reverse=True):
        x = x[:off] + txt + x[off:]

    # ---- 落盘前自检 ----
    ET.fromstring(x.encode('utf-8'))
    st2, en2 = bm_index(x)
    byname = {}
    for bid, (n, s0, s1) in st2.items():
        if bid in en2:
            byname[n] = (s1, en2[bid][0])
    for pi, name, bid, span_len, ok, note, kind in report:
        assert name in byname, name
        s1, e0 = byname[name]
        inner = x[s1:e0]
        assert len(inner) < MAX_SPAN, (name, len(inner))
        assert not re.search(r'<w:bookmarkStart\b[^>]*w:name="_Toc', inner), name
    bms = len(BMSTART.findall(x))
    bme = len(BMEND.findall(x))
    assert bms == bme, (bms, bme)
    print('自检通过: 新增 %d 个 _Ref；bookmarkStart=%d bookmarkEnd=%d'
          % (len(report), bms, bme))

    xml_bytes = x.encode('utf-8')

    zin = zipfile.ZipFile(SRC)
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
    print('已写出: %s  (%d bytes)' % (DST, os.path.getsize(DST)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
