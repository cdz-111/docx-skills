# -*- coding: utf-8 -*-
"""
recognize.py —— 语义识别：把 docx 解析成文档树。

**树是逻辑视图，不重建 OOXML**：每个节点持有原始 lxml 元素引用（Node.ref），
渲染只改这些元素。一旦按树重新生成 XML，域 / 书签 / 批注 / 图片关系全丢。

识别一律按语义（w:name + outlineLvl + 文本模式 + 直接格式），
**绝不按 styleId** —— 不同文档 styleId 是两套编号，WPS 保存还会重写。
"""
import re

from . import qn
from .docxio import has_drawing, in_table, p_text

# ------------------------------------------------------------------ 标题判定
RE_L1_CN = re.compile(r"^\s*([一二三四五六七八九十百]+)\s*、")
# 「1. 概述」这种阿拉伯数字一级标题（**只在已确认是 heading 1 时才认**，
# 否则表格里的「1.xxx」说明会被误判）
RE_L1_ARABIC = re.compile(r"^\s*(\d{1,2})\s*[.．]\s*\S")
RE_L3SEG = re.compile(r"^\s*\d{1,2}\s*[.．]\s*\d{1,2}\s*[.．]\s*\d{1,2}")
RE_L2SEG = re.compile(r"^\s*\d{1,2}\s*[.．]\s*\d{1,2}")
RE_L4_PAREN = re.compile(r"^\s*[（(]\s*\d{1,2}\s*[)）]")
RE_L2 = re.compile(r"^\s*\d{1,2}\s*[.．、]")
RE_APPENDIX = re.compile(r"^\s*(附录|附\s*录)\s*[A-Za-z一二三四五六七八九十\d]*")
RE_CAP = re.compile(r"^\s*(图|表)\s*(?:[xX]\s*|\d+\s*|[\u4e00-\u9fa5])")
RE_CAP_HEAD = re.compile(r"^\s*(图|表)(?:\s*([xX]|\d+)\s*|\s+)")
RE_TOC_STYLE = re.compile(r"^\s*(toc|WPSOffice手动目录|目录)\s*\d*\s*$", re.I)
RE_DRAW_ONLY = re.compile(r"^\s*(图|表)\s*(?:x|）|\)|$)", re.I)
# 整段就是一个公式编号：(1) / （12）
RE_FORMULA_ONLY = re.compile(r"^\s*[(（]\s*\d{1,3}\s*[)）]\s*$")

MAX_HEADING_LEN = 60

# 目录 / 清单标题
TOC_TITLES = {"目录", "目次", "插图清单", "附表清单", "图目录", "表目录",
              "图清单", "表清单", "图表目录"}


def _style_name(p, sid2name):
    ppr = p.find(qn("pPr"))
    if ppr is None:
        return None
    e = ppr.find(qn("pStyle"))
    if e is None:
        return None
    sid = e.get(qn("val"))
    return sid2name.get(sid, sid)


def _outline_lvl(p):
    ppr = p.find(qn("pPr"))
    if ppr is None:
        return None
    e = ppr.find(qn("outlineLvl"))
    if e is not None:
        try:
            return int(e.get(qn("val")))
        except (TypeError, ValueError):
            return None
    return None


def _is_bold_black(p, sid2name):
    """一级标题常用「黑体」。手搓文档里没有样式，只能看直接格式。"""
    nm = _style_name(p, sid2name) or ""
    if "黑体" in nm or (nm or "").lower() == "simhei":
        return True
    for r in p.findall(qn("r")):
        rpr = r.find(qn("rPr"))
        if rpr is None:
            continue
        f = rpr.find(qn("rFonts"))
        if f is not None:
            v = f.get(qn("eastAsia")) or f.get(qn("ascii")) or ""
            if "黑体" in v or "SimHei" in v:
                return True
    return False


def _bold_ratio(p):
    """整段是否「通体加粗」。标题段几乎总是通体加粗，正文加粗只是个别词。"""
    rs = p.findall(qn("r"))
    rs = [r for r in rs if r.find(qn("t")) is not None]
    if not rs:
        return 0.0
    b = 0
    for r in rs:
        rpr = r.find(qn("rPr"))
        if rpr is not None and rpr.find(qn("b")) is not None:
            v = rpr.find(qn("b")).get(qn("val"))
            if v not in ("0", "false"):
                b += 1
    return b / float(len(rs))


def _looks_like_heading(p, txt, sid2name=None):
    """段落看起来像标题吗（用于「1. xxx」这种弱信号）。

    两个判据：通体加粗 / 用了黑体。两者都不占就别当标题 ——
    宁可漏判也不要误判，误判会把正文抽成标题、目录里全是正文。
    """
    if _bold_ratio(p) >= 0.8:
        return True
    if _is_bold_black(p, sid2name or {}):
        return True
    return False


# ------------------------------------------------------------------ 裸标题识别
# 有些标题逐段信号一个都没有：不套 heading 样式、无 outlineLvl、文字里也没编号
# （不是「一、xxx」也不是「1. xxx」）—— 单段判不出来。
#
# 只保留一条排他性强的结构信号：
#    「下一段是 N.M 形式的小节标题」-> 说明本段是这一节的**父级标题**。
# 只有真正的章节起始，后面才会紧跟「1.1 概述」这类小节。
STRUCT_GAP = 3            # 允许隔几个空段
# 小节号：整数.整数，且**整数部分从 1 起**。
# ⛔ 不能放宽成 `\d+.\d+` —— 表格里的实验数据 `0.7084` 会被当成「0.70 小节」。
#    章节号从 1 开始，0.xxxx 是小数，必须排除。
RE_SUB_NUMBERING = re.compile(
    r"^\s*([1-9]\d?)\s*[.．]\s*([1-9]\d?)\s*[.．]?\s*\S")


def _is_sub_numbering(txt):
    """下一段是否是「N.M xxx」形式的小节标题（光学上的父级证据）。"""
    return bool(RE_SUB_NUMBERING.match(txt))


def find_bare_headings(paras, texts):
    """扫出「无样式无编号」的裸标题，返回 {段索引: (级别, 理由)}。

    判据只有一条：**下一段是 N.M 形式的小节标题**（N、M 均从 1 起）。
    这条信号几乎不会误判 —— 只有真正的章节起始，后面才会紧跟「1.1 概述」这类小节。
    另外几条必须加：
      - 本段**不在表格里**（表格数据天生符合所有形状特征）
      - 本段不是题注 / 公式编号 / 纯数字 / 纯符号

    ⛔ 级别不能一律当 1 级：下一段是「1.1」说明本段是**第 1 章**，
    下一段是「2.3.1」说明本段是**第 2 章的三级标题**。
    按小节号的段数推层级，否则会把原来的二三级标题全挤掉。
    """
    n = len(texts)
    out = {}
    for i in range(n):
        t = texts[i]
        if not t or len(t) > 30:
            continue
        if RE_CAP_HEAD.match(t) or RE_FORMULA_ONLY.match(t):
            continue
        if _is_numeric(t) or _is_punct(t):
            continue
        # ⛔ 表格里的段落一律不参与 —— 表格单元格天生短、纯数字、无标点结尾
        if in_table(paras[i]):
            continue
        # 下一段（允许隔空段）必须是 N.M 小节标题
        for k in range(1, STRUCT_GAP + 1):
            j = i + k
            if j >= n:
                break
            if not texts[j]:
                continue                      # 空段，继续往后看
            m = RE_SUB_NUMBERING.match(texts[j])
            if m:
                # 下一段是 N.M.K -> 本段是 N 章的 K 级标题
                # 下一段是 N.M   -> 本段是 N 章的 1 级标题（章名）
                tail = texts[j][m.end() - 1:]
                level = 1
                mm = re.match(r"^\s*(\d{1,2})\s*[.．]", tail)
                if mm:
                    level = min(6, 1 + int(mm.group(1)))
                else:
                    level = 1
                out[i] = (level,
                          "下一段是「%s」（小节标题）" % texts[j][:14])
                break
            break                           # 非空且不是小节标题 -> 放弃
    return out


def _is_numeric(t):
    """纯数字（含小数、百分号、正负号）—— 表格里的实验数据。"""
    return bool(re.match(r"^\s*[-+±]?[\d.,%±]+\s*$", t)) or bool(
        re.match(r"^\s*[-+]?\d+\.?\d*\s*[%±]?\s*$", t))


def _is_punct(t):
    """纯符号/标点。"""
    return bool(re.match(r"^\s*[-—–/\\|+·．.,，。：:；;、()（）\[\]【】<>《》]+\s*$", t))


def find_cover_end(paras, sid2name):
    """封面区结束位置 = **第一个一级标题**的下标。

    ⛔ 不能按「第一个分节符」算 —— 那样会把第一章开头连同正文一起吞进封面。
    目录条目要优先排除（目录里的「一、xx」跟一级标题同形）。

    ⚠️ 一级标题有两种形态：「一、xxx」（中文）和「1. xxx」（阿拉伯数字）。
       只认中文形态的文档（如纯 `1.` / `1.1` 编号的论文），会一个都找不到 →
       find_cover_end 退化成 len(paras) → **整篇都被当封面跳过**。
    """
    for i, p in enumerate(paras):
        nm = _style_name(p, sid2name) or ""
        if RE_TOC_STYLE.match(nm):
            continue
        t = p_text(p).strip()
        if not t:
            continue                       # 空标题段不能当封面终点
        # ① 已套 heading 1 样式 / outlineLvl=0 —— 最强信号，不管有没有编号文字
        #    （自动列表编号的文档里标题文字不带「一、」这类编号）
        if _is_h1(p, nm):
            return i
        # ② 「一、xxx」中文编号
        if RE_L1_CN.match(t) and len(t) <= MAX_HEADING_LEN:
            return i
        if RE_APPENDIX.match(t) and len(t) <= MAX_HEADING_LEN:
            return i
        # ③ 「1. xxx」阿拉伯数字编号（同样要求有 heading 1 信号佐证）
        if RE_L1_ARABIC.match(t) and len(t) <= MAX_HEADING_LEN \
                and _looks_like_heading(p, t, sid2name):
            return i
    return len(paras)


def _is_h1(p, nm=None):
    """段落是不是「一级标题」的最强信号：heading 1 样式 或 outlineLvl=0。"""
    if nm is None:
        nm = _style_name(p, {}) or ""
    if (nm or "").lower() in ("heading 1", "标题 1", "标题1"):
        return True
    return _outline_lvl(p) == 0


def find_toc_block(paras, sid2name):
    """找已有的目录区：返回 (起, 止) 下标区间（含目录标题，不含第一个一级标题）。

    用来「先删旧的再插新的」，避免目录重复。
    """
    start = None
    for i, p in enumerate(paras):
        t = p_text(p).strip()
        nm = _style_name(p, sid2name) or ""
        if start is None:
            if t in ("目录", "目次") and len(t) <= 4:
                start = i
            continue
        # 目录区结束：遇到第一个一级标题，或遇到超过 40 条无目录样式的内容
        if RE_L1_CN.match(t) and len(t) <= MAX_HEADING_LEN:
            return (start, i)
        if RE_APPENDIX.match(t) and len(t) <= MAX_HEADING_LEN:
            return (start, i)
        if nm and not RE_TOC_STYLE.match(nm) and t and not _is_toc_line(t):
            return (start, i)
    return (start, start) if start is not None else (None, None)


def _is_toc_line(t):
    """目录条目典型形态：以 tab + 数字页码结尾。"""
    return bool(re.search(r"\d+\s*$", t)) and ("\t" in t or "…" in t)


# ------------------------------------------------------------------ 节点
class Node:
    __slots__ = ("type", "id", "level", "text", "ref", "meta",
                 "children", "style", "parent", "sec_kind")

    def __init__(self, type_, id_=None, level=0, text="", ref=None,
                 meta=None, sec_kind=None):
        self.type = type_
        self.id = id_
        self.level = level
        self.text = text
        self.ref = ref
        self.meta = meta or {}
        self.sec_kind = sec_kind      # chapter / front / appendix
        self.children = []
        self.style = None
        self.parent = None

    def add(self, c):
        c.parent = self
        self.children.append(c)
        return c

    def walk(self, fn):
        fn(self)
        for c in self.children:
            c.walk(fn)

    def to_dict(self):
        d = {"type": self.type, "id": self.id, "level": self.level}
        if self.text:
            d["text"] = self.text[:120]
        if self.style:
            d["style"] = self.style
        if self.sec_kind:
            d["sec"] = self.sec_kind
        if self.meta:
            d["meta"] = {k: v for k, v in self.meta.items() if k != "ref"}
        if self.children:
            d["children"] = [c.to_dict() for c in self.children]
        return d

    @staticmethod
    def from_dict(d, index=None, parent=None):
        n = Node(d.get("type"), d.get("id"), d.get("level", 0),
                 d.get("text", ""), None, d.get("meta"), d.get("sec"))
        n.style = d.get("style")
        n.parent = parent
        if index and n.id in index:
            n.ref = index[n.id]
        for c in d.get("children", []):
            n.add(Node.from_dict(c, index, n))
        return n

    def __repr__(self):
        return "<%s %s L%s %r>" % (self.type, self.id, self.level,
                                   self.text[:20])


# ------------------------------------------------------------------ 分类
def _has_any_heading_signal(p, txt, sid2name, cover_end, idx):
    """本段是否**已经**有别的途径会被判成标题（有就别用裸标题通道）。

    判据与 classify_one 里的规则一一对应：heading 样式 / outlineLvl /
    文本编号正则 / 加粗黑体佐证。任何一条成立就让既有规则优先。
    """
    nm = (_style_name(p, sid2name) or "").lower()
    for lv in range(1, 7):
        if nm in ("heading %d" % lv, "标题 %d" % lv):
            return True
    if _outline_lvl(p) is not None and 0 <= _outline_lvl(p) <= 5:
        return True
    if len(txt) > MAX_HEADING_LEN:
        return False
    if (RE_L1_CN.match(txt) or RE_L2SEG.match(txt) or RE_L3SEG.match(txt)
            or RE_L4_PAREN.match(txt) or RE_L2.match(txt)
            or RE_APPENDIX.match(txt)):
        return True
    if RE_L1_ARABIC.match(txt) and _looks_like_heading(p, txt, sid2name):
        return True
    return False


def classify_one(p, txt, sid2name, cover_end, idx, state, bare=None):
    """判定单个段落的 (type, level)。state 跨段落传递层级推断。

    bare: {段索引: 判定理由} —— 结构信号认出的裸标题，优先于逐段规则。
    """
    nm = _style_name(p, sid2name) or ""

    if not txt:
        return "empty", 0
    # 目录条目先判：目录里的「一、xx」跟一级标题同形，但它是 toc 样式
    if RE_TOC_STYLE.match(nm):
        m = re.search(r"(\d)\s*$", nm)
        return "toc", int(m.group(1)) if m else 1
    if idx < cover_end:
        return "cover", 0
    # 目录 / 清单标题
    if txt in TOC_TITLES and len(txt) <= 6:
        return "toc_title", 0

    if in_table(p):
        return ("image" if (has_drawing(p) and len(txt) <= 12) else "table_text"), 0

    # 题注判定必须排在绘图判定前面（图片+题注同段的情况在 parse 前已拆开）
    if RE_CAP_HEAD.match(txt) and len(txt) <= 80:
        return "caption", 0

    if has_drawing(p):
        return ("image" if len(txt) <= 12 else "paragraph"), 0

    # 独立成段的公式编号：整段就是 (1) / （12）
    # 必须排在 heading 判定之前 —— 公式编号段虽然短，但不该进目录
    if RE_FORMULA_ONLY.match(txt) and not in_table(p):
        return "formula_number", 0

    # 附录
    if RE_APPENDIX.match(txt) and len(txt) <= MAX_HEADING_LEN:
        state["active3"] = False
        return "heading", 1

    # 裸标题（无样式无编号）—— 结构信号判定。
    # 只在本段**没有**任何既有标题信号时才用（下面的 heading 样式 / outlineLvl
    # / 文本正则都判不到），否则会抢掉已经套了 heading 样式的段。
    if (bare and idx in bare
            and not _has_any_heading_signal(p, txt, sid2name, cover_end, idx)):
        lv, _why = bare[idx]
        state["active3"] = (lv == 2)
        return "heading", lv

    # 已套 heading 样式的沿用原级别
    low = nm.lower()
    for lv in range(1, 7):
        if low in ("heading %d" % lv, "标题 %d" % lv):
            if lv <= 3:
                state["active3"] = False
            return "heading", lv

    # outlineLvl 也认（有些文档没挂 heading 样式但设了大纲级别）
    olv = _outline_lvl(p)
    if olv is not None and 0 <= olv <= 5:
        state["active3"] = (olv == 2)
        return "heading", olv + 1

    long_ = len(txt) > MAX_HEADING_LEN
    if RE_L1_CN.match(txt) and not long_:
        state["active3"] = False
        return "heading", 1
    # 「1. 概述」这种阿拉伯数字一级标题：必须字号/加粗明显不同于正文才认，
    # 否则表格里「1.xxx」的开头说明会被误判
    if RE_L1_ARABIC.match(txt) and not long_ and _looks_like_heading(
            p, txt, sid2name):
        state["active3"] = False
        return "heading", 1
    if RE_L3SEG.match(txt) and not long_:
        state["active3"] = True
        return "heading", 3
    if RE_L2SEG.match(txt) and not long_:
        state["active3"] = False
        return "heading", 3
    if RE_L4_PAREN.match(txt) and not long_:
        return "heading", 4
    if RE_L2.match(txt) and not long_:
        return "heading", (4 if state.get("active3") else 2)
    return "paragraph", 0


def mark_sections(root, skip_keywords):
    """给一级标题打 chapter / front / appendix 标记。

    引言/摘要这类前置部分和附录只算格式需要，不是正文章节，
    编号不应从它们起算。

    ⛔ 判定顺序：**先判附录，再判前置词**。
       `skip_keywords` 里本来就有「附录」两个字（为了让附录不占章号），
       如果先做关键词匹配，附录会被打成 `front` 而不是 `appendix` ——
       后果是 `lint` 的 `exclude: [appendix]` 与事实核查的「附录起自哪段」
       全部失效，附录内容被当成正文扫。
    """
    kws = tuple(skip_keywords or ())
    for n in _iter_all(root):
        if n.type == "heading" and n.level == 1:
            t = re.sub(r"\s+", "", n.text or "")
            if RE_APPENDIX.match(n.text or ""):
                n.sec_kind = "appendix"
            elif any(k in t for k in kws):
                n.sec_kind = "front"
            else:
                n.sec_kind = "chapter"
    return root


def _iter_all(root):
    out = []
    def rec(n):
        out.append(n)
        for c in n.children:
            rec(c)
    rec(root)
    return out


def parse(doc_root, sid2name, cover_end=None, skip_keywords=None,
          detect_bare=True):
    """OOXML -> 文档树。标题是容器节点，其下内容挂为 children。"""
    paras = list(doc_root.iter(qn("p")))
    texts = [p_text(p).strip() for p in paras]
    if cover_end is None:
        cover_end = find_cover_end(paras, sid2name)

    bare = {}
    if detect_bare:
        bare = find_bare_headings(paras, texts)
        # 已经被样式/编号认成标题的、或在封面区的，不算裸标题
        for i in list(bare):
            if i < cover_end:
                bare.pop(i)

    root = Node("document", "root")
    stack = [(0, root)]
    state = {"active3": False}
    index = {}

    for i, p in enumerate(paras):
        txt = texts[i]
        ntype, level = classify_one(p, txt, sid2name, cover_end, i, state,
                                    bare=bare)
        node = Node(ntype, "p%d" % i, level, txt, ref=p,
                    meta={"idx": i, "in_table": in_table(p),
                          "has_drawing": has_drawing(p),
                          "bare_heading": (bare.get(i) or (None, None))[1]})
        index[node.id] = p
        if ntype == "heading":
            while len(stack) > 1 and stack[-1][0] >= level:
                stack.pop()
            stack[-1][1].add(node)
            stack.append((level, node))
        else:
            stack[-1][1].add(node)

    if skip_keywords is not None:
        mark_sections(root, skip_keywords)
    return root, index, cover_end


# ------------------------------------------------------------------ 报告
def tree_report(root, out=None, max_text=44, max_depth=5):
    lines = []

    def rec(n, depth):
        if n.type == "document":
            for c in n.children:
                rec(c, depth)
            return
        if n.type == "empty":
            return
        pad = "  " * depth
        tag = "H%d" % n.level if n.type == "heading" else n.type
        st = (" -> %s" % n.style) if n.style else ""
        sk = (" [%s]" % n.sec_kind) if n.sec_kind else ""
        t = re.sub(r"\s+", " ", n.text)[:max_text]
        lines.append("%s[%s] %s%s%s  %s" % (pad, tag, n.id, st, sk, t))
        if depth < max_depth:
            for c in n.children:
                rec(c, depth + 1)

    rec(root, 0)
    s = "\n".join(lines)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(s)
    return s


def stats(root):
    from collections import Counter
    c = Counter()
    def cnt(n):
        if n.type != "document":
            c[n.type] += 1
            if n.type == "heading":
                c["heading%d" % n.level] += 1
            if n.sec_kind:
                c["sec_" + n.sec_kind] += 1
    root.walk(cnt)
    return c
