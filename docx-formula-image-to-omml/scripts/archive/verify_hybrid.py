# -*- coding: utf-8 -*-
"""（历史脚本，已归档；路径参数须由环境变量或显式传参提供）
9-28 产出校验：结构 / 计数 / 编号 / 引用 / 媒体完整性 / 段落 diff。
"""
import zipfile, re, os, json

WORK = os.environ.get('DOCXFMT_WORKDIR')
SRC = os.path.join(WORK, '9-28.backup-20260928-005343.docx')
OUT = os.environ.get('DOCXFMT_OUT')

zs, zo = zipfile.ZipFile(SRC), zipfile.ZipFile(OUT)
xs = zs.read('word/document.xml').decode('utf-8')
xo = zo.read('word/document.xml').decode('utf-8')
PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')

print('=' * 72)
print('一、XML 合法性')
print('=' * 72)
import xml.etree.ElementTree as ET
try:
    ET.fromstring(xo.encode('utf-8'))
    print('  ✓ document.xml 解析通过')
except Exception as e:
    print('  ✗ 解析失败:', e)

print('\n' + '=' * 72)
print('二、结构计数')
print('=' * 72)
for k, f in (('原', xs), ('新', xo)):
    print('  %s: 段落=%-5d drawing=%-4d oMath=%-4d oMathPara=%-3d oMathParaPr=%d'
          % (k, len(PARA.findall(f)), len(re.findall(r'<w:drawing>', f)),
             len(re.findall(r'<m:oMath>', f)), len(re.findall(r'<m:oMathPara>', f)),
             len(re.findall(r'<m:oMathParaPr>', f))))
print('  预期: drawing 90→44, oMath 341→387, oMathPara 7→0')

print('\n' + '=' * 72)
print('三、媒体完整性（逐条目 CRC 比对）')
print('=' * 72)
crc_s = {i.filename: (i.CRC, i.file_size) for i in zs.infolist()}
crc_o = {i.filename: (i.CRC, i.file_size) for i in zo.infolist()}
print('  条目数: 原 %d / 新 %d' % (len(crc_s), len(crc_o)))
diff = [n for n in crc_s if crc_s[n] != crc_o.get(n)]
print('  CRC 有差异的条目: %s' % (diff or '无'))
media = [n for n in crc_s if n.startswith('word/media/')]
print('  word/media 条目 %d 个，全部字节一致: %s'
      % (len(media), all(crc_s[n] == crc_o[n] for n in media)))

print('\n' + '=' * 72)
print('四、公式编号（应严格 1..75）')
print('=' * 72)
labs = []
for q in PARA.findall(xo):
    t = ''.join(T.findall(re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S))).strip()
    m = re.fullmatch(r'\((\d+)\)', t)
    if m:
        labs.append(int(m.group(1)))
print('  编号个数: %d' % len(labs))
print('  全序列: %s' % labs)
print('  严格 1..75: %s' % (labs == list(range(1, 76))))
dup = sorted(k for k in set(labs) if labs.count(k) > 1)
print('  重复: %s' % (dup or '无'))

print('\n' + '=' * 72)
print('五、公式段制表位')
print('=' * 72)
bad = 0
for i, q in enumerate(PARA.findall(xo)):
    t = ''.join(T.findall(re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S))).strip()
    if not re.fullmatch(r'\(\d+\)', t):
        continue
    tb = re.findall(r'<w:tab w:val="(\w+)" w:pos="(\d+)"', q)
    if tb != [('center', '4153'), ('right', '8306')] or '<m:oMath>' not in q:
        bad += 1
        print('  ⚠ para %d 制表位=%s 有公式=%s' % (i, tb, '<m:oMath>' in q))
print('  制表位/结构异常: %d 个' % bad)

print('\n' + '=' * 72)
print('六、正文引用与残留')
print('=' * 72)
print('  残留 公式（33）: %d' % xo.count('公式（33）'))
print('  新   公式（30）: %d' % xo.count('公式（30）'))
print('  所有“公式（”引用: %s' % re.findall(r'公式（\d+）', xo))

# 被转换掉的公式图是否还有引用
INL = ['image4', 'image5', 'image6', 'image9', 'image12', 'image16', 'image17', 'image18',
       'image20', 'image21', 'image22', 'image23', 'image24', 'image25', 'image26', 'image27',
       'image28', 'image30', 'image32', 'image33', 'image42', 'image43', 'image51', 'image52',
       'image55', 'image56'] + ['image7', 'image8', 'image10', 'image11', 'image13', 'image14',
                                'image19', 'image29', 'image31', 'image37', 'image38', 'image39',
                                'image40', 'image41', 'image53', 'image57']
rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(media/[^"]+)"',
                       zo.read('word/_rels/document.xml.rels').decode('utf-8')))
used = set(re.findall(r'r:embed="(rId\d+)"', xo))
still = sorted({rels[r] for r in used if r in rels})
leftover = [m for m in still if os.path.basename(m).rsplit('.', 1)[0] in INL]
print('  仍被引用的公式图: %s' % (leftover or '无（全部已转为原生公式）'))
print('  仍被引用的图片总数: %d' % len(still))

print('\n' + '=' * 72)
print('七、段落文本级 diff（应只动公式段/引用段）')
print('=' * 72)
pa, pb = PARA.findall(xs), PARA.findall(xo)


def txt(q):
    return ''.join(T.findall(re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S)))


print('  段落数 原=%d 新=%d' % (len(pa), len(pb)))
n = max(len(pa), len(pb))
diff = [i for i in range(n) if (pa[i] if i < len(pa) else None) != (pb[i] if i < len(pb) else None)]
print('  差异段落: %d 个' % len(diff))
kinds = {'公式段': 0, '引用段': 0, '其他': 0}
others = []
for i in diff:
    a, b = txt(pa[i]), txt(pb[i])
    ismath = '<m:oMath>' in pa[i] or '<m:oMath>' in pb[i] or '<w:drawing>' in pa[i]
    if ismath:
        kinds['公式段'] += 1
    elif '公式（' in a or '公式（' in b:
        kinds['引用段'] += 1
    else:
        kinds['其他'] += 1
        others.append((i, a[:60], b[:60]))
print('  分类: %s' % kinds)
for i, a, b in others[:10]:
    print('    para %d  原=%r  新=%r' % (i, a, b))
