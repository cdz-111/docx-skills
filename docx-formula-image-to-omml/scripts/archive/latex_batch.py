# -*- coding: utf-8 -*-
"""（历史脚本，已归档；路径参数须由环境变量或显式传参提供）
第 1 章公式：LaTeX → OMML（pandoc 一次批量），产出 omml_928.json。
"""
import subprocess, re, json, os, zipfile

WORK = os.environ.get('DOCXFMT_WORKDIR')
PANDOC = os.environ.get('PANDOC') or 'pandoc'

# ---------- 行间公式（按段落号，文档顺序即编号顺序）----------
DISP = [
    (224,  r'\mathbf{e}_k=\mathcal{G}(\mathbf{h}_k^{\mathrm{ctx}},\mathbf{h}_k^{\mathrm{seq}}),\quad \forall k\in\{q,c\}'),
    (226,  r'P(\hat{y}=1|x)=\frac{P_\theta(v_{\mathrm{yes}}|x)}{P_\theta(v_{\mathrm{yes}}|x)+P_\theta(v_{\mathrm{no}}|x)}'),
    (228,  r'\mathbf{r}_k^{\mathrm{ctx}}=\mathcal{E}_{\mathrm{ctx}}(T_k),\quad \forall k\in\{q,c\}'),
    (229,  r'\mathbf{h}_k^{\mathrm{ctx}}=\mathcal{P}_\phi(\mathbf{r}_k^{\mathrm{ctx}}),\quad \forall k\in\{q,c\}'),
    (231,  r'\mathbf{r}_k^{\mathrm{seq}}=\mathcal{E}_{\mathrm{seq}}(x_k),\quad \forall k\in\{q,c\}'),
    (232,  r'\mathbf{h}_k^{\mathrm{seq}}=\mathcal{P}_\theta(\mathbf{r}_k^{\mathrm{seq}}),\quad \forall k\in\{q,c\}'),
    (242,  r'\Psi_{ljk}=b_{jk}+d_j\delta_{lj}X_{jk}+\varepsilon_{ljk}'),
    (245,  r'P(\delta_{lj}=1)=\frac{1}{1+\exp(-\alpha_j^{\mathrm{T}}\theta_l)}'),
    (248,  r'\Theta^*=\arg\max_{\Theta}\sum_{l\in L_{\mathrm{tr}}}\sum_{j\in J}\log P(R_{lj}=r_{lj})'),
    (878,  r'S_\eta=\frac{1}{|\Omega_\eta|}\sum_{i=1}^{|\Omega_\eta|}\left(5-\alpha_p\times D_i-\alpha_p^{m}\times D_i^{m}+\alpha_r\times L_i\right)'),
    (880,  r'S_\eta^{\mathrm{ALL}}=\beta\times S_\eta^{\mathrm{P}\text{-}\mathrm{R}}+(1-\beta)\times S_\eta^{\mathrm{R}\text{-}\mathrm{A}}'),
    (883,  r'R_e(\tau|G)=\mathrm{score}(\tau,G)'),
    (885,  r'R_a(t_i|P_{\mathrm{sim}})=R_{\mathrm{fmt}}(t_i)+R_{\mathrm{com}}(t_i)+R_{\mathrm{int}}(t_i|P_{\mathrm{sim}})'),
    (888,  r'A(t_i^{(n)})=A^{\mathrm{T}}(\tau_n)+\omega\cdot A^{\mathrm{S}}(t_i^{(n)})'),
    (1712, r'|f(x)-f(x_\epsilon)|\le\epsilon'),
    (1717, r'k=\frac{s_{n-1}-s_{\min}}{s_{\max}-s_{\min}}+0.5'),
]

# ---------- 行内符号（按媒体文件）----------
INLINE = {
    'image4.jpeg':  r'\mathbf{h}_k^{\mathrm{ctx}}',
    'image5.jpeg':  r'\mathbf{h}_k^{\mathrm{seq}}',
    'image6.jpeg':  r'\mathbf{e}_k',
    'image9.jpeg':  r'\mathbf{r}_k^{\mathrm{ctx}}',
    'image12.jpeg': r'\mathbf{r}_k^{\mathrm{seq}}',
    'image16.jpeg': r'l',
    'image17.jpeg': r'j',
    'image18.jpeg': r'k',
    'image20.jpeg': r'b_{jk}',
    'image21.jpeg': r'\varepsilon_{ljk}',
    'image22.jpeg': r'\delta_{lj}',
    'image23.jpeg': r'X_{jk}',
    'image24.jpeg': r'k',
    'image25.jpeg': r'd_j',
    'image26.jpeg': r'j',
    'image27.jpeg': r'\theta_l',
    'image28.jpeg': r'\alpha_j',
    'image30.jpeg': r'\alpha_j^{\mathrm{T}}\theta_l',
    'image32.jpeg': r'\mathbf{D}_j=[\mathbf{d}_j,\mathbf{b}_j,\boldsymbol{\alpha}_j]',
    'image33.jpeg': r'\mathbf{D}_j',
    'image42.jpeg': r'A^{\mathrm{T}}',
    'image43.jpeg': r'A^{\mathrm{S}}',
    'image51.jpeg': r'\mathcal{V}_\epsilon\subseteq\mathbb{R}^{D}',
    'image52.jpeg': r'\mathbf{x}\in\mathbb{R}^{D}',
    'image55.jpeg': r'T=L\left(1+\frac{d^{(t)}-d_1}{d_h-d_1}\right)\frac{\mathrm{budget}}{2\beta}',
    'image56.jpeg': r's_i=-\frac{b_{i+1}-b_i}{d_{i+1}-d_i}',
}

# 统一批量转换
items = [('D%d' % p, tex) for p, tex in DISP] + [('I' + k, tex) for k, tex in INLINE.items()]
md = '\n\n'.join('$' + t + '$' for _, t in items)
mdp = os.path.join(WORK, '_f928.md')
dxp = os.path.join(WORK, '_f928.docx')
open(mdp, 'w', encoding='utf-8').write(md)

r = subprocess.run([PANDOC, '-f', 'markdown', '-t', 'docx', '-o', dxp, mdp],
                   capture_output=True, text=True)
print('pandoc rc=%s  stderr=%s' % (r.returncode, r.stderr.strip()[:300]))

x = zipfile.ZipFile(dxp).read('word/document.xml').decode('utf-8')
x = re.sub(r'</?m:oMathPara[^>]*>', '', x)      # 去掉段落级包裹
found = re.findall(r'<m:oMath>.*?</m:oMath>', x, re.S)
print('输入 %d 条，产出 oMath %d 条' % (len(items), len(found)))
assert len(found) == len(items), '数量不匹配！'

omml = {key: f for (key, _), f in zip(items, found)}
json.dump(omml, open(os.path.join(WORK, 'omml_928.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
print('已写 omml_928.json')

# 抽查 3 条
for k in ('D226', 'D245', 'Iimage32.jpeg', 'Iimage55.jpeg'):
    s = omml[k]
    txt = ''.join(re.findall(r'<m:t(?: [^>]*)?>([^<]*)</m:t>', s))
    print('  %-16s len=%-5d 文本=%r' % (k, len(s), txt))
print('  D880 文本=%r' % ''.join(re.findall(r'<m:t[^>]*>([^<]*)</m:t>', omml['D880'])))
