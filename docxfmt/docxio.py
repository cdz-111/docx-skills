# -*- coding: utf-8 -*-
"""
docxio.py —— 底层 IO 与 OOXML 构造工具。

只做「读进来 / 写出去 / 造元素」三件事，不含任何排版判断。
"""
import copy
import os
import shutil
import zipfile

from lxml import etree

from . import XSPACE, qn

CT_MAIN = ("application/vnd.openxmlformats-officedocument"
           ".wordprocessingml.document.main+xml")
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


# ------------------------------------------------------------------ 打包 / 解包
def unpack(docx, workdir):
    """docx -> 解包目录。已存在则先清空（shutil.rmtree 对大目录比 os.remove 稳）。"""
    if os.path.exists(workdir):
        shutil.rmtree(workdir, ignore_errors=True)
    os.makedirs(workdir, exist_ok=True)
    with zipfile.ZipFile(docx) as z:
        z.extractall(workdir)
    return workdir


def pack(workdir, out):
    """解包目录 -> docx。zipfile 'w' 本就覆盖写，无需先删旧文件。"""
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    n = 0
    try:
        zf = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    except PermissionError:
        # 最常见的原因：文档正被 Word/WPS 打开着
        raise SystemExit(
            "[pack] 写不进去：%s\n"
            "       十有八九是这个文件正被 Word/WPS 打开着。关掉后重试，\n"
            "       或者换一个输出路径。" % out)
    with zf as z:
        # [Content_Types].xml 必须第一个
        ct = os.path.join(workdir, "[Content_Types].xml")
        if os.path.exists(ct):
            z.write(ct, "[Content_Types].xml")
            n += 1
        for root, _dirs, files in os.walk(workdir):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, workdir).replace("\\", "/")
                if rel == "[Content_Types].xml":
                    continue
                z.write(full, rel)
                n += 1
    return n


def read_style_names(docx_or_dir):
    """读 styleId -> w:name 映射。给 zip 路径或解包目录都行。"""
    if os.path.isdir(docx_or_dir):
        p = os.path.join(docx_or_dir, "word", "styles.xml")
        st = etree.parse(p).getroot()
    else:
        with zipfile.ZipFile(docx_or_dir) as z:
            st = etree.fromstring(z.read("word/styles.xml"))
    sm = {}
    for s in st.iter(qn("style")):
        nm = s.find(qn("name"))
        sm[s.get(qn("styleId"))] = (nm.get(qn("val")) if nm is not None else "")
    return sm


# ------------------------------------------------------------------ 文档对象
class Docx:
    """一份解包后的文档。所有修改都作用在内存树上，最后 save() 落盘。"""

    def __init__(self, workdir):
        self.dir = workdir
        self.doc_path = os.path.join(workdir, "word", "document.xml")
        self.sty_path = os.path.join(workdir, "word", "styles.xml")
        self._tree = None
        self._sty = None

    # ---- document.xml
    @property
    def tree(self):
        if self._tree is None:
            self._tree = etree.parse(self.doc_path)
        return self._tree

    @property
    def root(self):
        return self.tree.getroot()

    @property
    def body(self):
        return self.root.find(qn("body"))

    def paras(self):
        """物化段落列表。

        lxml 每次 iter() 返回**新的代理对象**，id() 不稳定，
        所以必须一次性 list()，再用索引做键。
        """
        return list(self.root.iter(qn("p")))

    # ---- styles.xml
    @property
    def sty(self):
        if self._sty is None:
            self._sty = etree.parse(self.sty_path)
        return self._sty

    def save(self):
        """⚠️ 必须在所有改动之后调用，且是最后一步。"""
        self.sty.write(self.sty_path, xml_declaration=True,
                       encoding="UTF-8", standalone=True)
        self.tree.write(self.doc_path, xml_declaration=True,
                        encoding="UTF-8", standalone=True)


# ------------------------------------------------------------------ 元素构造
def el(tag, **attrs):
    e = etree.Element(qn(tag))
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def mk_run(text=None, rpr=None, preserve=None, tag="t"):
    """造一个 w:r。rpr 必须是 w:r 的首个子元素（OOXML 顺序要求）。"""
    r = etree.Element(qn("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if text is not None:
        e = etree.SubElement(r, qn(tag))
        e.text = text
        if preserve is None:
            preserve = (text != text.strip())
        if preserve:
            e.set(XSPACE, "preserve")
    return r


def field_runs(instr, cached=None, rpr=None):
    """造一组构成一个域的 run：begin / instrText / separate / 结果 / end。

    域结构（括号等字面量必须写在域**外**，不要塞进 instrText）：
        fldChar begin | instrText " SEQ 图 \\* ARABIC " | separate | t "1" | end
    """
    out = [mk_run(rpr=rpr), mk_run(rpr=rpr)]
    out[0].append(el("fldChar", fldCharType="begin"))
    it = etree.SubElement(out[1], qn("instrText"))
    it.set(XSPACE, "preserve")
    it.text = instr
    if cached is not None:
        out.append(mk_run(rpr=rpr))
        out[-1].append(el("fldChar", fldCharType="separate"))
        out.append(mk_run(text=cached, rpr=rpr))
    out.append(mk_run(rpr=rpr))
    out[-1].append(el("fldChar", fldCharType="end"))
    return out


def get_ppr(p):
    ppr = p.find(qn("pPr"))
    if ppr is None:
        ppr = etree.Element(qn("pPr"))
        p.insert(0, ppr)
    return ppr


def set_style(p, style_id):
    """给段落套样式。pPr 内 pStyle 必须是第一个子元素。"""
    ppr = get_ppr(p)
    ps = ppr.find(qn("pStyle"))
    if ps is None:
        ps = etree.Element(qn("pStyle"))
        ppr.insert(0, ps)
    ps.set(qn("val"), style_id)


def get_style(p):
    ppr = p.find(qn("pPr"))
    if ppr is None:
        return None
    ps = ppr.find(qn("pStyle"))
    return ps.get(qn("val")) if ps is not None else None


def p_text(p):
    return "".join(t.text or "" for t in p.iter(qn("t")))


def para_of(node):
    """从任意 run/字段节点往上找到它所属的 w:p。"""
    while node is not None and node.tag != qn("p"):
        node = node.getparent()
    return node


def in_table(node):
    n = node.getparent()
    while n is not None:
        if n.tag == qn("tbl"):
            return True
        n = n.getparent()
    return False


def has_drawing(node):
    return node.find(".//" + qn("drawing")) is not None


def first_rpr(p):
    """段落里第一个带 rPr 的 run 的 rPr，拿来当新 run 的格式来源。"""
    for r in p.findall(qn("r")):
        rr = r.find(qn("rPr"))
        if rr is not None:
            return rr
    return None


# ------------------------------------------------------------------ 书签 / 域工具
def max_bookmark_id(root):
    mx = 0
    for e in root.iter(qn("bookmarkStart")):
        try:
            mx = max(mx, int(e.get(qn("id"), "0")))
        except (TypeError, ValueError):
            pass
    return mx


def new_bookmark_id(root):
    return max_bookmark_id(root) + 1


def bookmark(name, bk_id):
    s = el("bookmarkStart", id=bk_id, name=name)
    e = el("bookmarkEnd", id=bk_id)
    return s, e


def iter_fields(p):
    """抽出段落里的域，yield (指令, 缓存结果)。"""
    instr, result, on = [], [], False
    for c in p.iter():
        if c.tag == qn("fldChar"):
            t = c.get(qn("fldCharType"))
            if t == "begin":
                instr, result, on = [], [], False
            elif t == "separate":
                on = True
            elif t == "end":
                if instr:
                    yield ("".join(instr).strip(), "".join(result).strip())
                instr, result, on = [], [], False
        elif c.tag == qn("instrText"):
            instr.append(c.text or "")
        elif c.tag == qn("t") and on:
            result.append(c.text or "")


def has_field(p, kind="REF"):
    return any(kind.upper() in (it.text or "").upper()
               for it in p.iter(qn("instrText")))
