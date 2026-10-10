# -*- coding: utf-8 -*-
"""终极核查：图表顺序 / 交叉引用 是否被改动过。
A = 基准文件（改动前）
B = 产出文件
C = 当前文件

被编排调用时（docx-unify-pipeline 的 S2）按下列顺序取入参，**缺失则报错**：
    A / B / C = ① --a/--b/--c 或位置参数  ② 环境变量 DOCXFMT_A/B/C
"""
import os
import sys
import zipfile, re

PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')

_argv = sys.argv[1:]
_pos = [a for a in _argv if not a.startswith('--')]


def _pick(flag, pos, env, default=None):
    """① --flag 的值 → ② 位置参数 → ③ 环境变量。"""
    if flag in _argv:
        k = _argv.index(flag)
        if k + 1 < len(_argv):
            return _argv[k + 1]
    if len(_pos) > pos:
        return _pos[pos]
    return os.environ.get(env) or default


_A = _pick('--a', 0, 'DOCXFMT_A')
_B = _pick('--b', 1, 'DOCXFMT_B')
_C = _pick('--c', 2, 'DOCXFMT_C')
if not (_A and _B and _C):
    sys.exit('用法: final_check.py --a <基准.docx> --b <产出.docx> --c <当前.docx>'
             '（或三个位置参数，或设 DOCXFMT_A / DOCXFMT_B / DOCXFMT_C）')

FILES = [
    ('A 基准文件', _A),
    ('B 产出文件', _B),
    ('C 当前文件', _C),
]


def load(p):
    return zipfile.ZipFile(p).read('word/document.xml').decode('utf-8')


def seqs(x):
    """返回 (图 SEQ 号列表, 表 SEQ 号列表, 每个 SEQ 的 para 序号)"""
    out = {}
    paras = PARA.findall(x)
    for kind in ('图', '表'):
        rows = []
        for i, q in enumerate(paras):
            if 'SEQ %s' % kind not in q:
                continue
            segs = re.findall(r'separate"/>(.*?)<w:fldChar w:fldCharType="end"/>', q, re.S)
            num = ''
            for sg in segs:
                n = ''.join(T.findall(sg)).strip()
                if re.fullmatch(r'\d+', n):
                    num = n
            rows.append((i, num))
        out[kind] = rows
    return out


def bms(x):
    return sorted(re.findall(r'<w:bookmarkStart[^>]*w:name="([^"]+)"', x))


def refs(x):
    return re.findall(r'REF\s+(_Ref\w+)', x)


def pagerefs(x):
    return re.findall(r'PAGEREF\s+(_Ref\w+)', x)


def drawings(x):
    return re.findall(r'r:(?:embed|link)="(rId\d+)"', x)


data = {}
for tag, p in FILES:
    x = load(p)
    data[tag] = x

print('=' * 78)
print('一、图片/题注顺序')
print('=' * 78)
for tag, _ in FILES:
    s = seqs(data[tag])
    f = [n for _, n in s['图']]
    t = [n for _, n in s['表']]
    print('\n%s' % tag)
    print('  图 SEQ 缓存值: %s' % f)
    print('  表 SEQ 缓存值: %s' % t)
    print('  图号重复: %s' % sorted(k for k in set(f) if f.count(k) > 1))
    print('  表号重复: %s' % sorted(k for k in set(t) if t.count(k) > 1))
    print('  图 严格1..%d: %s' % (len(f), f == [str(i) for i in range(1, len(f) + 1)]))
    print('  表 严格1..%d: %s' % (len(t), t == [str(i) for i in range(1, len(t) + 1)]))
    print('  drawing 引用顺序: %s' % drawings(data[tag]))

print('\n' + '=' * 78)
print('二、三版本对比（顺序是否被改）')
print('=' * 78)
tags = [t for t, _ in FILES]
for key in ('图', '表'):
    a = [n for _, n in seqs(data[tags[0]])[key]]
    b = [n for _, n in seqs(data[tags[1]])[key]]
    c = [n for _, n in seqs(data[tags[2]])[key]]
    print('\n%s 号序列:' % key)
    print('  A: %s' % a)
    print('  B: %s' % b)
    print('  C: %s' % c)
    print('  A==B: %s   B==C: %s   A==C: %s' % (a == b, b == c, a == c))

for tag, _ in FILES:
    d = drawings(data[tag])
    print('\n%s' % tag)
print('\ndrawing 序列 A==B: %s  B==C: %s' % (
    drawings(data[tags[0]]) == drawings(data[tags[1]]),
    drawings(data[tags[1]]) == drawings(data[tags[2]])))

print('\n' + '=' * 78)
print('三、交叉引用完整性')
print('=' * 78)
for tag, _ in FILES:
    x = data[tag]
    B = bms(x)
    R = refs(x)
    P = pagerefs(x)
    broken = [r for r in set(R) if r not in B]
    print('\n%s' % tag)
    print('  bookmark 数: %d' % len(B))
    print('  REF 域: %d   指向不存在书签: %s' % (len(R), broken or '无'))
    print('  PAGEREF 域: %d  指向不存在书签: %s' % (
        len(P), [r for r in set(P) if r not in B] or '无'))
    err = sum(len(T.findall(m)) for m in re.findall(r'错误[^<]*', x))
    print('  含"错误!未找到引用源"字样: %d' % x.count('未找到引用源'))

print('\n' + '=' * 78)
print('四、书签集合差异（A vs B vs C）')
print('=' * 78)
A, B, C = (set(bms(data[t])) for t in tags)
print('  B 比 A 多: %s' % sorted(B - A))
print('  B 比 A 少: %s' % sorted(A - B))
print('  C 比 B 多: %s' % sorted(C - B))
print('  C 比 B 少: %s' % sorted(B - C))

print('\n' + '=' * 78)
print('五、正文段落差异（A vs B）')
print('=' * 78)
pa = PARA.findall(data[tags[0]])
pb = PARA.findall(data[tags[1]])
print('  段落数 A=%d  B=%d' % (len(pa), len(pb)))
n = max(len(pa), len(pb))
diff = [i for i in range(n) if (pa[i] if i < len(pa) else None) != (pb[i] if i < len(pb) else None)]
print('  内容不同的段落: %d 个' % len(diff))
for i in diff:
    a = ''.join(T.findall(pa[i])) if i < len(pa) else ''
    b = ''.join(T.findall(pb[i])) if i < len(pb) else ''
    kind = '公式' if ('m:oMath' in (pa[i] + (pb[i] if i < len(pb) else ''))) else '文本'
    print('    para %-5d [%s]  A=%r  B=%r' % (i, kind, a[:40], b[:40]))
