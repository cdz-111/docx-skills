# -*- coding: utf-8 -*-
"""
styles.py —— 样式表管理。

铁律：**一切按 `w:name` 语义定位，绝不按 styleId。**
不同文档的 styleId 是两套编号；WPS 保存还会把语义 styleId 重写成数字。
所以：绝不能直接复制模板的 styles.xml，只能按 `w:name` 移植定义 + 重映射 basedOn。
"""
import copy

from lxml import etree

from . import qn
from .docxio import el

# 角色 -> 模板里可能叫什么名字（不同来源的中文/英文命名差异）
# ⚠️ 必须是「一组别名」：同一角色在不同模板里名字完全不同 ——
#    「标题 1」在中文模板叫 标题 1，在 WPS/英文模板叫 heading 1。
NAME_ALIASES = {
    "正文": ["正文", "Normal", "Body Text", "body", "Normal (Web)"],
    "题注": ["题注", "caption", "Caption"],
    "图片": ["图片", "Image", "image"],
    "表格表头文字": ["表格表头文字", "表格文字", "TableText"],
    "目录标题": ["目录标题", "TocTitle", "TOC Title", "TOC标题"],
    "附录": ["附录", "Appendix"],
    "公式编号": ["公式编号", "FormulaNumber"],
    "页眉": ["页眉", "header", "Header"],
    "页脚": ["页脚", "footer", "Footer"],
}
# 标题 / 目录样式别名动态生成（标题 1 <-> heading 1 <-> Heading1）
for _i in range(1, 7):
    NAME_ALIASES["标题 %d" % _i] = [
        "标题 %d" % _i, "标题%d" % _i, "heading %d" % _i,
        "Heading %d" % _i, "Heading%d" % _i]
    NAME_ALIASES["目录 %d" % _i] = [
        "目录 %d" % _i, "目录%d" % _i, "toc %d" % _i,
        "TOC %d" % _i, "toc%d" % _i]


class StyleTable:
    """按 w:name 索引的样式表。"""

    def __init__(self, sty_root):
        self.root = sty_root
        self.by_name = {}
        self.by_id = {}
        for s in sty_root.iter(qn("style")):
            sid = s.get(qn("styleId"))
            nm = s.find(qn("name"))
            nmv = (nm.get(qn("val")) if nm is not None else "") or ""
            self.by_id[sid] = s
            # 英文模板里 name 可能是 "heading 1"，中文是 "标题 1"，都归一
            self.by_name.setdefault(nmv.strip(), s)
            self.by_name.setdefault(nmv.strip().lower(), s)
        self._next_id = 1000

    # ---- 查
    def sid(self, name):
        """按 w:name 取 styleId。找不到返回 None。"""
        s = self.get(name)
        return s.get(qn("styleId")) if s is not None else None

    def get(self, name):
        if not name:
            return None
        s = self.by_name.get(name)
        if s is None:
            s = self.by_name.get(name.lower())
        return s

    def resolve(self, want, aliases=None):
        """容错解析：先按原名，再按别名表。

        want 可能是 "正文"，模板里叫 "Normal" —— 也能找到。
        """
        s = self.get(want)
        if s is not None:
            return s
        for a in (aliases or []):
            s = self.get(a)
            if s is not None:
                return s
        return None

    def ensure(self, name, based_on=None, next_style=None, kind="paragraph",
               quick=True):
        """确保存在一个叫 name 的样式，没有就建。返回 styleId。"""
        s = self.get(name)
        if s is not None:
            return s.get(qn("styleId"))
        while ("style%s" % self._next_id) in self.by_id:
            self._next_id += 1
        sid = "style%s" % self._next_id
        self._next_id += 1
        st = el("style", type=kind, styleId=sid)
        st.append(el("name", val=name))
        if based_on:
            bsid = self.sid(based_on)
            if bsid:
                st.append(el("basedOn", val=bsid))
        if next_style:
            nsid = self.sid(next_style)
            if nsid:
                st.append(el("next", val=nsid))
        if quick:
            st.append(el("qFormat"))
        self.root.append(st)
        self.by_id[sid] = st
        self.by_name[name] = st
        self.by_name[name.lower()] = st
        return sid

    def new_id(self):
        while ("style%s" % self._next_id) in self.by_id:
            self._next_id += 1
        sid = "style%s" % self._next_id
        self._next_id += 1
        return sid

    # ---- 改
    def child(self, style_el, tag, create=True):
        c = style_el.find(qn(tag))
        if c is None and create:
            c = etree.SubElement(style_el, qn(tag))
        return c

    def set_ppr(self, name, **kw):
        """改段落级几何。

        kw 支持：align(center/left/right/both) / first_line(pt) / left(pt) /
        before / after(pt) / line(倍数或 pt) / keep_next(bool) / page_break(bool)
        """
        st = self.get(name)
        if st is None:
            return False
        ppr = self.child(st, "pPr")
        jc = kw.get("align")
        if jc:
            self.child(ppr, "jc").set(qn("val"), {
                "center": "center", "left": "left",
                "right": "right", "both": "both"}[jc])
        if "first_line" in kw:
            ind = self.child(ppr, "ind")
            v = int(round(kw["first_line"] * 20))       # pt -> twips
            ind.set(qn("firstLineChars"), "0")
            ind.set(qn("firstLine"), str(v))
        if "left" in kw:
            self.child(ppr, "ind").set(qn("left"),
                                       str(int(round(kw["left"] * 20))))
        sp = kw.get("before", kw.get("after"))
        if sp is not None or "line" in kw:
            sppr = self.child(ppr, "spacing")
            if sp is not None:
                sppr.set(qn("beforeLines"), "0")
                sppr.set(qn("afterLines"), "0")
                sppr.set(qn("before"), str(int(round(sp * 20))))
                sppr.set(qn("after"), str(int(round(sp * 20))))
            if "line" in kw:
                v = kw["line"]
                if isinstance(v, str) and v.endswith("x"):
                    sppr.set(qn("lineRule"), "auto")
                    sppr.set(qn("line"), str(int(float(v[:-1]) * 240)))
                else:
                    sppr.set(qn("lineRule"), "exact")
                    sppr.set(qn("line"), str(int(round(v * 20))))
        for flag, tag in (("keep_next", "keepNext"),
                          ("page_break", "pageBreakBefore"),
                          ("keep_lines", "keepLines")):
            if kw.get(flag):
                self.child(ppr, tag)
        return True

    def set_rpr(self, name, size=None, font=None, bold=None, color=None):
        st = self.get(name)
        if st is None:
            return False
        rpr = self.child(st, "rPr")
        if size is not None:
            # sz 单位是**半磅**（size=15pt -> w:sz val="30"）
            self.child(rpr, "sz").set(qn("val"), str(int(round(size * 2))))
            self.child(rpr, "szCs").set(qn("val"), str(int(round(size * 2))))
        if font:
            f = self.child(rpr, "rFonts")
            f.set(qn("ascii"), font)
            f.set(qn("hAnsi"), font)
            f.set(qn("eastAsia"), font)
        if bold is not None:
            for t in ("b", "bCs"):
                e = self.child(rpr, t)
                if not bold:
                    e.set(qn("val"), "0")
        if color:
            self.child(rpr, "color").set(qn("val"), color)
        return True


# ------------------------------------------------------------------ 移植
def _based_on_id(st, table):
    b = st.find(qn("basedOn"))
    if b is None:
        return None
    return b.get(qn("val"))


def _set_child_val(st, tag, val):
    e = st.find(qn(tag))
    if e is None:
        e = etree.SubElement(st, qn(tag))
    e.set(qn("val"), val)
    return e


def transplant(target_root, template_root, style_names):
    """把模板里指定的样式定义移植到目标 styles.xml。

    绝不整体复制 styles.xml —— 两边 styleId 体系不同，会造成语义错位。

    style_names: 要移植的样式 w:name 列表
    返回 (覆盖数, 新增数, 重映射 basedOn 处数, 缺失列表)
    """
    ttab = StyleTable(template_root)
    ytab = StyleTable(target_root)

    # 1) 目标里缺的样式先建壳（保留目标自己的 styleId，语义才不会错位）
    missing = []
    added = 0
    for name in style_names:
        t = ttab.resolve(name, NAME_ALIASES.get(name))
        if t is None:
            missing.append(name)
            continue
        if ytab.get(name) is None:
            tname = t.find(qn("name"))
            tname_v = tname.get(qn("val")) if tname is not None else name
            ytab.ensure(tname_v, kind=t.get(qn("type"), "paragraph"),
                        quick=False)
            added += 1

    # 2) 覆盖同名样式的定义
    over = 0
    id_map = {}
    for name in style_names:
        t = ttab.resolve(name, NAME_ALIASES.get(name))
        if t is None:
            continue
        tname = t.find(qn("name"))
        tname_v = (tname.get(qn("val")) if tname is not None else name)
        y = ytab.get(tname_v)
        if y is None:
            y = ytab.get(name)
        if y is None:
            continue
        # 记住「模板 id -> 目标 id」，下面重映射 basedOn/next/link 用
        id_map[t.get(qn("styleId"))] = y.get(qn("styleId"))
        # 替换除 styleId / name 之外的所有子元素
        keep = {qn("styleId"), qn("name")}
        for c in list(y):
            if c.tag not in keep:
                y.remove(c)
        for c in list(t):
            if c.tag in keep:
                continue
            y.append(copy.deepcopy(c))
        over += 1

    # 3) basedOn / next / link 里的模板 id 要换成目标 id
    remap = 0
    for st in target_root.iter(qn("style")):
        for tag in ("basedOn", "next", "link"):
            e = st.find(qn(tag))
            if e is None:
                continue
            v = e.get(qn("val"))
            if v in id_map and id_map[v] != v:
                e.set(qn("val"), id_map[v])
                remap += 1

    return over, added, remap, missing


def patch_heading_indent(target_root, heading_names=("标题 1", "标题 2", "标题 3",
                                                     "标题 4", "标题 5", "标题 6",
                                                     "heading 1", "heading 2",
                                                     "heading 3", "heading 4")):
    """标题样式兜底 firstLine=0。

    坑：模板里没用过的样式（如 heading 4）pPr 里没有 ind，
    会继承正文的「首行缩进 2 字」，标题就跟着缩进了。
    """
    tab = StyleTable(target_root)
    n = 0
    for nm in heading_names:
        st = tab.get(nm)
        if st is None:
            continue
        ppr = st.find(qn("pPr"))
        if ppr is None:
            ppr = etree.SubElement(st, qn("pPr"))
        ind = ppr.find(qn("ind"))
        if ind is None:
            ind = etree.SubElement(ppr, qn("ind"))
        if ind.get(qn("firstLine")) not in ("0",) or ind.get(qn("firstLineChars")):
            n += 1
        ind.set(qn("firstLine"), "0")
        ind.set(qn("firstLineChars"), "0")
        ind.set(qn("left"), "0")
    return n
