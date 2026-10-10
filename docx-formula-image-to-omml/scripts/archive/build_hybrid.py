# -*- coding: utf-8 -*-
"""9-28 重建：第1章公式图→原生OMML + 全文公式连续编号 (1)-(75)。
产出 C:\\Users\\14711\\Desktop\\中期科技报告9-28-公式规范化.docx
"""
import zipfile, re, json, os, shutil

WORK = r'D:\appdata\workbuddy\2026-09-24-23-47-06\work'
SRC = os.path.join(WORK, '9-28.backup-20260928-005343.docx')
OUT = r'C:\Users\14711\Desktop\中期科技报告9-28-公式规范化.docx'
z = zipfile.ZipFile(SRC)
xml = z.read('word/document.xml').decode('utf-8')
rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
REL = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
omml = json.load(open(os.path.join(WORK, 'omml_928.json'), encoding='utf-8'))

PARA = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>|<w:p/>', re.S)
PSTART = [m.start() for m in PARA.finditer(xml)]
PEND = [m.end() for m in PARA.finditer(xml)]
ps = [m.group(0) for m in PARA.finditer(xml)]
T = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')
RUN = re.compile(r'<w:r(?: [^>]*)?>.*?</w:r>', re.S)

# ---------- 1. 目标段落定位 ----------
CH1_DISP = [224, 226, 228, 229, 231, 232, 242, 245, 248, 878, 880, 883, 885, 888, 1712, 1717]

CH23_NUM = []
for i, q in enumerate(ps):
    if i < 1720:
        continue
    t = ''.join(T.findall(re.sub(r'<w:drawing>.*?</w:drawing>', '', q, flags=re.S))).strip()
    if re.fullmatch(r'\(\d+\)', t):
        CH23_NUM.append(i)

old_labels = [int(re.search(r'\((\d+)\)', ''.join(T.findall(ps[i]))).group(1)) for i in CH23_NUM]
OMATHPARA = [i for i, q in enumerate(ps) if i > 1720 and '<m:oMathPara>' in q]

print('第1章行间公式 %d 个 | ch2/3 已编号公式 %d 个 | §3.4.2 无编号 %d 个'
      % (len(CH1_DISP), len(CH23_NUM), len(OMATHPARA)))
print('现有编号序列首尾: %s … %s   严格20..71: %s'
      % (old_labels[:4], old_labels[-4:], old_labels == list(range(20, 72))))

total = len(CH1_DISP) + len(CH23_NUM) + len(OMATHPARA)
assert total == 75, '编号总数不是 75（实际 %d），请检查' % total
print('全文显示公式合计: %d → 编号 (1)-(75)' % total)

NEW_CH1 = {p: k + 1 for k, p in enumerate(CH1_DISP)}
NEW_CH23 = {p: 17 + k for k, p in enumerate(CH23_NUM)}
NEW_OP = {p: 69 + k for k, p in enumerate(OMATHPARA)}
OLD2NEW = dict(zip(old_labels, sorted(NEW_CH23.values())))
print('旧→新映射: 20→%d, 33→%d, 71→%d' % (OLD2NEW[20], OLD2NEW[33], OLD2NEW[71]))

# ---------- 2. 生成编辑 ----------
edits = []


def disp_para(open_tag, om, num):
    return ('%s<w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/>'
            '<w:tab w:val="right" w:pos="8306"/></w:tabs>'
            '<w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:tab/></w:r>%s<w:r><w:tab/></w:r><w:r><w:t>(%d)</w:t></w:r></w:p>'
            % (open_tag, om, num))


# (a) 第 1 章行间公式段：图片 → 制表位 + 原生公式 + 编号
for p in CH1_DISP:
    open_tag = re.match(r'<w:p(?: [^>]*)?>', ps[p]).group(0)
    edits.append((PSTART[p], PEND[p], disp_para(open_tag, omml['D%d' % p], NEW_CH1[p])))

# (b) §3.4.2 oMathPara → 制表位版式 + 编号
for p in OMATHPARA:
    open_tag = re.match(r'<w:p(?: [^>]*)?>', ps[p]).group(0)
    inner = re.search(r'<m:oMathPara>.*?(<m:oMath>.*?</m:oMath>).*?</m:oMathPara>',
                      ps[p], re.S).group(1)
    inner = re.sub(r'<w:sz w:val="22"/>', '', inner)      # 与全文公式统一字号
    edits.append((PSTART[p], PEND[p], disp_para(open_tag, inner, NEW_OP[p])))

# (c) ch2/3 标签改号（只改 <w:t> 里的纯数字）
cnt_lab = 0
for p in CH23_NUM:
    q = ps[p]
    ms = list(re.finditer(r'(<w:t(?: [^>]*)?>)\((\d+)\)(</w:t>)', q))
    assert len(ms) == 1, 'para %d 的标签数=%d' % (p, len(ms))
    m = ms[0]
    edits.append((PSTART[p] + m.start(2), PSTART[p] + m.end(2), str(NEW_CH23[p])))
    cnt_lab += 1

# (d) ch1 行内图片 → 行内 oMath
cnt_inl = 0
for i, q in enumerate(ps):
    if i >= 1720 or i in CH1_DISP or '<w:drawing>' not in q:
        continue
    base = PSTART[i]
    for rm in RUN.finditer(q):
        r = rm.group(0)
        rid = re.search(r'r:embed="(rId\d+)"', r)
        if not rid:
            continue
        if '<w:t' in r:
            raise SystemExit('run 同时含图片和文本，需手工处理: para %d' % i)
        media = REL.get(rid.group(1), '').replace('media/', '')
        key = 'I' + media
        if key not in omml:            # 真插图，保持不动
            continue
        edits.append((base + rm.start(), base + rm.end(), omml[key]))
        cnt_inl += 1
assert cnt_inl == 30, '行内公式替换数不是 30（实际 %d）' % cnt_inl
print('行内替换 %d 处 | 标签改号 %d 处' % (cnt_inl, cnt_lab))

# (e) 正文引用 公式（33）→ 公式（30）
REF_OLD, REF_NEW = 33, OLD2NEW[33]
cnt_ref = 0
for i, q in enumerate(ps):
    old, new = '公式（%d）' % REF_OLD, '公式（%d）' % REF_NEW
    if old not in ''.join(T.findall(q)):
        continue
    for m in re.finditer(r'(<w:t(?: [^>]*)?>)([^<]*)</w:t>', q):
        val = m.group(2)
        if old in val:
            edits.append((PSTART[i] + m.start(2), PSTART[i] + m.end(2),
                          val.replace(old, new)))
            cnt_ref += 1
assert cnt_ref == 1, '正文引用改写数量异常: %d' % cnt_ref
print('正文引用改写: 公式（%d）→ 公式（%d）' % (REF_OLD, REF_NEW))

# ---------- 3. 应用 ----------
edits.sort(key=lambda e: e[0], reverse=True)
prev = len(xml) + 1
for s, e, t in edits:
    assert e <= prev, '编辑区间重叠: %d-%d (prev=%d)' % (s, e, prev)
    xml = xml[:s] + t + xml[e:]
    prev = s
print('共应用 %d 处编辑' % len(edits))

# ---------- 4. 重新打包（逐条目复制，保留原 ZipInfo）----------
zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'word/document.xml':
        data = xml.encode('utf-8')
    zout.writestr(item, data)
zout.close()
print('\n已写出 %s  (%.1f MB)' % (OUT, os.path.getsize(OUT) / 1048576))
