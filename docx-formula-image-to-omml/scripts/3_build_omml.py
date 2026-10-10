# -*- coding: utf-8 -*-
"""步骤3：LaTeX 批量转 OMML。
用法: python 3_build_omml.py <workdir> <pandoc.exe>
约定（与 4_rebuild.py 对齐）：
  disp_latex.py  ->  DISP = [dict(file=, sec=, lines=[latex,...]), ...]   按文档顺序
  inline_latex.py ->  INLINE = [latex, ...]  与 inline_list.json 同序
产出: omml_display.json  (key '文件名#序号' -> <m:oMath>…)
      omml_inline.json   (key '文件名'      -> <m:oMath>…)
"""
import json, os, re, sys, subprocess, zipfile

W, PANDOC = sys.argv[1], sys.argv[2]
sys.path.insert(0, W)


def to_omml(pairs):
    keys, lats = zip(*pairs)
    md = '\n\n'.join('$' + l.strip().replace('\n', ' ') + '$' for l in lats)
    mdp, dxp = os.path.join(W, '_math.md'), os.path.join(W, '_math.docx')
    open(mdp, 'w', encoding='utf-8').write(md)
    r = subprocess.run([PANDOC, '-f', 'markdown', '-t', 'docx', '-o', dxp, mdp],
                       capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit('pandoc failed: ' + r.stderr[:2000])
    x = zipfile.ZipFile(dxp).read('word/document.xml').decode('utf-8')
    x = re.sub(r'</?m:oMathPara[^>]*>', '', x)          # 只保留裸 oMath
    found = re.findall(r'<m:oMath>.*?</m:oMath>', x, re.S)
    if len(found) != len(keys):
        print('WARN count mismatch: got %d want %d' % (len(found), len(keys)))
    return dict(zip(keys, found))


from disp_latex import DISP
from inline_latex import INLINE

dp = []
for d in DISP:
    for i, l in enumerate(d['lines']):
        dp.append(('%s#%d' % (d['file'], i), l))
m1 = to_omml(dp)
json.dump(m1, open(os.path.join(W, 'omml_display.json'), 'w', encoding='utf-8'), ensure_ascii=False)
print('display equations ->', len(m1))

il = json.load(open(os.path.join(W, 'inline_list.json'), encoding='utf-8'))
assert len(il) == len(INLINE), 'inline_latex.py 条数(%d) 与 inline_list.json(%d) 不一致' % (len(INLINE), len(il))
m2 = to_omml([(il[i]['file'], INLINE[i]) for i in range(len(il))])
json.dump(m2, open(os.path.join(W, 'omml_inline.json'), 'w', encoding='utf-8'), ensure_ascii=False)
print('inline formulas ->', len(m2))
