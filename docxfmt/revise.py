# -*- coding: utf-8 -*-
"""
revise.py —— 把「已批准」的改动写进文档，**以修订模式呈现**，每条挂批注说明原因。

设计边界（刻意保守）：
    * 只执行 decision ∈ {approved, modified} 的项；其余原样不动。
    * 只认领 executor ∈ {text_replace, run_dedup}；其余 executor
      （manual / pagebreak / table_border / insert_caption）**原样退回**，
      记入回执「转交他处」，不记失败。
    * 一处替换**跨多个 run** 时不做通用的自动拼接 —— 直接进回执「失败」栏，交人工。
      **唯一例外**：改动只是「删空白」时（`after == before 去掉空白`），逐 run 就地
      删空白是机械安全的，自动处理（见 `_try_ws_delete`）。这条例外覆盖了
      WPS 把 `精度 0.8554` 拆成 `<w:t>精度 </w:t><w:t>0.8554</w:t>` 的绝大多数真实形态。
    * 涉及域 / 书签 / 编号的改动不走这里（那些交给编号治理流程，风险等级 high）。

写盘纪律：备份（带时间戳）→ 在临时副本上改 → 验收 → 才落盘。
"""
import copy
import json
import os
import re
import shutil
import time
import zipfile

from lxml import etree

from . import W, XSPACE, qn
from .crs import CRS_VERSION, approved_items, load, summarize
from .docxio import Docx, unpack

DEFAULT_AUTHOR = "docx-format 校对"
#: 本执行器认领的 executor。一份批准清单会被 revise / polish / captions 三个
#: 执行器依次读，不属于自己的条目应**原样退回**（记入回执「转交他处」），
#: 而不是记成"失败"——否则回执里会堆一大片假失败，把真失败淹掉。
SUPPORTED = ("text_replace", "run_dedup")
# 会被 "去空格" 规则盯上的空白字符
_WS = " \u3000\t"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_CT = ("application/vnd.openxmlformats-officedocument"
               ".wordprocessingml.comments+xml")

# w:settings 子元素的 schema 顺序（只列到 trackChanges 附近，够用）
_SETTINGS_ORDER = [
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "removeDateAndTime", "doNotDisplayPageBoundaries",
    "displayBackgroundShape", "printPostScriptOverText",
    "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
    "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState",
    "formsDesign", "attachedTemplate", "linkStyles",
    "stylePaneFormatFilter", "stylePaneSortMethod", "documentType",
    "mailMerge", "revisionView", "trackChanges",
]


def _now_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def _stamp():
    return time.strftime("%Y%m%d_%H%M%S", time.localtime())


# ------------------------------------------------------------------ 主流程
def apply(docx, crs_path, approved_path=None, out=None, author=DEFAULT_AUTHOR,
          dry=False, log=print):
    """返回 receipt dict（含 已改/未勾选未改/失败 三栏）。"""
    crs = load(crs_path)
    todo_all = approved_items(crs, approved_path) if approved_path else [
        i for i in crs["items"] if i["decision"] in ("approved", "modified")]
    pending = [i for i in crs["items"]
               if i["decision"] not in ("approved", "modified", "closed")]
    # 分流：只认领文字类 executor；其余（manual / pagebreak / table_border /
    # insert_caption）原样退回，由 polish / captions / 人工接手。
    todo, others = [], []
    for it in todo_all:
        ex = it.get("executor")
        if ex in SUPPORTED:
            todo.append(it)
        else:
            others.append(it)
    if others:
        log("[scope] %d 项不属于文字类（executor∈%s），原样退回"
            % (len(others), sorted({i.get("executor") for i in others})))

    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_revise")
    unpack(docx, workdir)
    doc = Docx(workdir)

    # ---- 备份
    backup = None
    if not dry:
        backup = os.path.join(
            os.path.dirname(os.path.abspath(docx)),
            "%s_备份_%s.docx" % (os.path.splitext(os.path.basename(docx))[0],
                                 _stamp()))
        shutil.copy2(docx, backup)

    before_counts = _counts(doc.root)
    cm = _Comments(workdir)
    rev = _RevIds(doc.root)
    author_s = author
    date_s = _now_iso()

    ok, fail = [], []
    paras = list(doc.root.iter(qn("p")))
    for it in todo:
        try:
            n = _apply_one(doc.root, paras, it, cm, rev, author_s, date_s)
            if n:
                ok.append(it)
            else:
                fail.append((it, "定位不到 before 串（文档可能已变，需重新出单）"))
        except _Skip as e:
            fail.append((it, str(e)))
        except Exception as e:                      # noqa: BLE001
            fail.append((it, "%s: %s" % (type(e).__name__, e)))

    if cm.dirty:
        cm.write()

    # ---- 落盘
    receipt = {
        "crs_version": CRS_VERSION,
        "target": crs.get("target", {}),
        "summary": {"total": len(crs["items"]), "approved": len(todo_all),
                    "changed": len(ok), "pending": len(pending),
                    "failed": len(fail), "other": len(others)},
        "changed": ok, "pending": pending, "others": others,
        "failed": [{"id": i["id"], "reason": r} for i, r in fail],
        "backup": backup, "backup_md5": _md5(backup) if backup else None,
        "dry": dry,
    }

    if dry:
        log("[dry] 未落盘。将改 %d 项，失败 %d 项" % (len(ok), len(fail)))
        shutil.rmtree(workdir, ignore_errors=True)
        return receipt

    # 验收
    after_counts = _counts(doc.root)
    problems = _validate(before_counts, after_counts, doc.root, cm)
    if problems:
        log("[验收] 未通过，**不落盘**：")
        for x in problems:
            log("   - " + x)
        shutil.rmtree(workdir, ignore_errors=True)
        receipt["failed"].append({"id": "-", "reason": "验收未通过：" +
                                  "；".join(problems)})
        return receipt

    doc.save()
    try:
        _enable_track_changes(workdir)
    except Exception as e:                          # noqa: BLE001
        log("[warn] 未能开启 settings.trackChanges（不影响已有修订标记）：%s" % e)
    out = out or os.path.join(
        os.path.dirname(os.path.abspath(docx)),
        "%s_修订.docx" % os.path.splitext(os.path.basename(docx))[0])
    _pack(workdir, out)
    shutil.rmtree(workdir, ignore_errors=True)
    receipt["output"] = out
    log("[ok] %d 项已改（修订模式），%d 项失败 -> %s"
        % (len(ok), len(fail), out))
    return receipt


class _Skip(Exception):
    pass


def _apply_one(root, paras, it, cm, rev, author, date):
    if it.get("executor") == "manual":
        raise _Skip("该项需人工处理（executor=manual）")
    para_no = (it.get("where") or {}).get("para") or 0
    if not para_no or para_no >= len(paras):
        raise _Skip("段号 %s 越界" % para_no)
    p = paras[para_no]
    if it.get("executor") == "run_dedup":
        return _apply_run_dedup(p, it, cm, rev, author, date)
    before = it.get("before") or ""
    after = it.get("after_override") or it.get("after") or ""
    if not before:
        raise _Skip("缺 before 串")

    loc = _locate(p, before)
    if loc is None:
        # 同一处空白常被两条规则同时盯上（如 `， 为` 与 ` 为候选框坐标，`），
        # 前一条已把空格删了，后一条自然找不到 before —— 这不是缺陷。
        if _ws_only(before, after) and _locate(p, after) is not None:
            raise _Skip("与同段另一条重复，该空白已由前者移除（已生效）")
        return 0
    # 落在域结果区（交叉引用/页码缓存）里的文字，一律不动
    if _runs_in_field_result(p) & {r for (r, _a, _b) in loc}:
        raise _Skip("位于域结果区（交叉引用缓存），不动 —— Word 会按题注重算")
    if len(loc) != 1:
        # 特例：跨 run，但改动**只是删空白** —— 逐 run 就地删，机械安全。
        # （WPS 常把 `精度 0.8554` 拆成 <w:t>精度 </w:t><w:t>0.8554</w:t>）
        if _ws_only(before, after) and _try_ws_delete(
                p, loc, before, after, it, cm, rev, author, date):
            return 1
        raise _Skip("匹配跨 %d 个 run，保守起见交人工" % len(loc))
    r, a, b = loc[0]
    parent = r.getparent()
    rpr = r.find(qn("rPr"))
    old = "".join(t.text or "" for t in r.findall(qn("t")))
    # run 里除 rPr/w:t 外还有别的东西（书签、域、br…）-> 不动
    for c in r:
        if c.tag not in (qn("rPr"), qn("t")):
            raise _Skip("run 内含非文本元素（%s），不动" %
                        etree.QName(c).localname)

    cid = cm.add_comment(author, date,
                         "[%s|%s] %s" % (it["id"], it.get("category", ""),
                                         it.get("basis", "")))
    nodes = []
    if a > 0:
        nodes.append(_run(old[:a], rpr))
    nodes.append(_cmark("commentRangeStart", cid))
    nodes.append(_ins_or_del("del", old[a:b], rpr, rev.next(), author, date))
    if after:
        nodes.append(_ins_or_del("ins", after, rpr, rev.next(), author, date))
    nodes.append(_cmark("commentRangeEnd", cid))
    nodes.append(_comment_ref(cid))
    if b < len(old):
        nodes.append(_run(old[b:], rpr))

    idx = list(parent).index(r)
    parent.remove(r)
    for k, nd in enumerate(nodes):
        parent.insert(idx + k, nd)
    return 1


def _apply_run_dedup(p, it, cm, rev, author, date):
    """把「两个相邻且内容相同的 run」中的第二个以修订模式删掉。"""
    keep = it.get("after_override") or it.get("after") or ""
    if not keep:
        raise _Skip("run_dedup 缺 after（要保留哪一份）")
    rs = [(r, "".join(t.text or "" for t in r.findall(qn("t"))))
          for r in p.iter(qn("r"))]
    for i in range(len(rs) - 1):
        r1, t1 = rs[i]
        r2, t2 = rs[i + 1]
        if t1 == t2 == keep and r1.getparent() is r2.getparent():
            for c in r2:
                if c.tag not in (qn("rPr"), qn("t")):
                    raise _Skip("要删的 run 内含非文本元素，不动")
            rpr = r2.find(qn("rPr"))
            cid = cm.add_comment(author, date, "[%s|%s] %s"
                                 % (it["id"], it.get("category", ""),
                                    it.get("basis", "")))
            nodes = [_cmark("commentRangeStart", cid),
                     _ins_or_del("del", t2, rpr, rev.next(), author, date),
                     _cmark("commentRangeEnd", cid), _comment_ref(cid)]
            parent = r2.getparent()
            idx = list(parent).index(r2)
            parent.remove(r2)
            for k, nd in enumerate(nodes):
                parent.insert(idx + k, nd)
            return 1
    return 0


# ------------------------------------------------------------------ 定位
def _locate(p, needle):
    """在段落的文本 run 序列里定位 needle。返回 [(run, local_a, local_b)] 或 None。"""
    runs = []
    for r in p.iter(qn("r")):
        ts = r.findall(qn("t"))
        if not ts:
            continue
        s = "".join(t.text or "" for t in ts)
        if s:
            runs.append((r, s))
    total = "".join(s for _r, s in runs)
    i = total.find(needle)
    if i < 0:
        return None
    out, pos = [], 0
    for r, s in runs:
        rs, re_ = pos, pos + len(s)
        a, b = max(i, rs), min(i + len(needle), re_)
        if a < b:
            out.append((r, a - rs, b - rs))
        pos = re_
    return out


def _runs_in_field_result(p):
    """返回段内处于「域结果区」（`separate` 之后、`end` 之前）的文本 run 集合。

    域结果区里躺着的是**缓存文本**（如 REF 的 `图 1`）。缓存由 Word 按域语义渲染，
    改它既无意义（F9 就改回来）又有害（缓存与题注不一致 → 编号治理验收直接 FAIL）。
    """
    out = set()
    stack = []
    for el in p.iter():
        tag = el.tag
        if tag == qn("fldChar"):
            t = el.get(qn("fldCharType"))
            if t == "begin":
                stack.append("begin")
            elif t == "separate" and stack:
                stack[-1] = "result"
            elif t == "end" and stack:
                stack.pop()
        elif tag == qn("r"):
            # 承载 fldChar 的那个 run 不算结果区
            if el.find(qn("fldChar")) is None and "result" in stack:
                out.add(el)
    return out


def _ws_only(before, after):
    """before 去掉所有空白后是否恰好等于 after（即这条改动只是「删空白」）。"""
    if not after:
        return False
    return "".join(ch for ch in before if ch not in _WS) == after


def _try_ws_delete(p, loc, before, after, it, cm, rev, author, date):
    """跨 run，但**只删空白**——逐 run 就地删，机械安全，不动书签/域。

    为什么敢自动做：`after == before 去掉空白`，所以**一个字都不用插**，
    只要在各自 run 里把空白标成删除即可。这恰好覆盖 WPS 把 `精度 0.8554`
    拆成 `<w:t>精度 </w:t><w:t>0.8554</w:t>` 这类绝大多数真实形态。

    返回 True=已处理，False=不满足安全条件（交调用方报 `_Skip`）。
    """
    # 1) 每个受影响的 run 只能含 rPr / w:t，否则不动（书签、域、br 都在这里被挡下）
    for (r, _a, _b) in loc:
        for c in r:
            if c.tag not in (qn("rPr"), qn("t")):
                return False
    # 2) 拼接匹配片段必须精确等于 before（防止 _locate 的定位与实际不符）
    joined = "".join(
        ("".join(t.text or "" for t in r.findall(qn("t"))))[a:b]
        for (r, a, b) in loc)
    if joined != before:
        return False

    cid = cm.add_comment(author, date,
                         "[%s|%s] %s" % (it["id"], it.get("category", ""),
                                         it.get("basis", "")))
    last = loc[-1][0]
    for (r, a, b) in loc:
        txt = "".join(t.text or "" for t in r.findall(qn("t")))
        rpr = r.find(qn("rPr"))
        parent = r.getparent()
        nodes = []
        if r is loc[0][0]:
            nodes.append(_cmark("commentRangeStart", cid))
        if a > 0:
            nodes.append(_run(txt[:a], rpr))
        # 把 [a,b) 切成 连续同类 的片段：保留的用普通 run，空白的用 w:del
        seg = txt[a:b]
        frags = []
        for ch in seg:
            kind = "del" if ch in _WS else "keep"
            if frags and frags[-1][0] == kind:
                frags[-1][1] += ch
            else:
                frags.append([kind, ch])
        for kind, s in frags:
            if kind == "keep":
                nodes.append(_run(s, rpr))
            else:
                nodes.append(_ins_or_del("del", s, rpr, rev.next(),
                                         author, date))
        if b < len(txt):
            nodes.append(_run(txt[b:], rpr))
        if r is last:
            nodes.append(_cmark("commentRangeEnd", cid))
            nodes.append(_comment_ref(cid))
        ii = list(parent).index(r)
        parent.remove(r)
        for k, nd in enumerate(nodes):
            parent.insert(ii + k, nd)
    return True


# ------------------------------------------------------------------ XML 构造
def _run(text, rpr):
    r = etree.Element(qn("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    t.set(XSPACE, "preserve")
    return r


def _ins_or_del(kind, text, rpr, rid, author, date):
    e = etree.Element(qn(kind))
    e.set(qn("id"), str(rid))
    e.set(qn("author"), author)
    e.set(qn("date"), date)
    r = etree.SubElement(e, qn("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, qn("delText") if kind == "del" else qn("t"))
    t.text = text
    t.set(XSPACE, "preserve")
    return e


def _cmark(kind, cid):
    e = etree.Element(qn(kind))
    e.set(qn("id"), str(cid))
    return e


def _comment_ref(cid):
    r = etree.Element(qn("r"))
    rpr = etree.SubElement(r, qn("rPr"))
    etree.SubElement(rpr, qn("rStyle")).set(qn("val"), "CommentReference")
    etree.SubElement(rpr, qn("sz")).set(qn("val"), "18")
    ref = etree.SubElement(r, qn("commentReference"))
    ref.set(qn("id"), str(cid))
    return r


class _RevIds:
    def __init__(self, root):
        mx = 0
        for tag in ("ins", "del", "moveFrom", "moveTo"):
            for e in root.iter(qn(tag)):
                try:
                    mx = max(mx, int(e.get(qn("id"), "0")))
                except (TypeError, ValueError):
                    pass
        self._n = mx

    def next(self):
        self._n += 1
        return self._n


# ------------------------------------------------------------------ 批注部件
class _Comments:
    def __init__(self, workdir):
        self.dir = workdir
        self.path = os.path.join(workdir, "word", "comments.xml")
        self.dirty = False
        if os.path.exists(self.path):
            self.root = etree.parse(self.path).getroot()
        else:
            self.root = etree.Element(qn("comments"), nsmap={"w": W})
        self._next = self._max_id()

    def _max_id(self):
        mx = -1
        for c in self.root.iter(qn("comment")):
            try:
                mx = max(mx, int(c.get(qn("id"), "-1")))
            except (TypeError, ValueError):
                pass
        return mx

    def add_comment(self, author, date, text):
        self._next += 1
        cid = self._next
        c = etree.SubElement(self.root, qn("comment"))
        c.set(qn("id"), str(cid))
        c.set(qn("author"), author)
        c.set(qn("date"), date)
        c.set(qn("initials"), "DF")
        p = etree.SubElement(c, qn("p"))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text
        t.set(XSPACE, "preserve")
        self.dirty = True
        return cid

    def write(self):
        uniq = set()
        for c in self.root.iter(qn("comment")):
            i = c.get(qn("id"))
            assert i not in uniq, "批注 id 重复：%s" % i
            uniq.add(i)
        etree.ElementTree(self.root).write(
            self.path, xml_declaration=True, encoding="UTF-8",
            standalone=True)
        _ensure_comments_plumbing(self.dir)


def _ensure_comments_plumbing(workdir):
    # [Content_Types].xml
    ct_path = os.path.join(workdir, "[Content_Types].xml")
    ct = etree.parse(ct_path)
    r = ct.getroot()
    part = "/word/comments.xml"
    if not any(o.get("PartName") == part for o in r.iter("{%s}Override" % CT_NS)):
        o = etree.SubElement(r, "{%s}Override" % CT_NS)
        o.set("PartName", part)
        o.set("ContentType", COMMENTS_CT)
        ct.write(ct_path, xml_declaration=True, encoding="UTF-8",
                 standalone=True)
    # document.xml.rels
    rels_path = os.path.join(workdir, "word", "_rels", "document.xml.rels")
    rels = etree.parse(rels_path)
    r = rels.getroot()
    tgt_type = R_NS + "/comments"
    if not any(x.get("Type") == tgt_type for x in r.iter("{%s}Relationship" % PKG_REL_NS)):
        used = {x.get("Id") for x in r.iter("{%s}Relationship" % PKG_REL_NS)}
        n = 1
        while ("rId%d" % n) in used:
            n += 1
        e = etree.SubElement(r, "{%s}Relationship" % PKG_REL_NS)
        e.set("Id", "rId%d" % n)
        e.set("Type", tgt_type)
        e.set("Target", "comments.xml")
        rels.write(rels_path, xml_declaration=True, encoding="UTF-8",
                   standalone=True)


# ------------------------------------------------------------------ 验收
def _counts(root):
    c = {}
    for tag in ("p", "tbl", "drawing", "bookmarkStart", "bookmarkEnd",
                "fldChar", "instrText"):
        c[tag] = len(list(root.iter(qn(tag))))
    return c


def _validate(before, after, root, cm):
    problems = []
    for k in ("p", "tbl", "drawing", "bookmarkStart", "bookmarkEnd",
              "fldChar", "instrText"):
        if before[k] != after[k]:
            problems.append("%s 计数从 %d 变成 %d（不允许）"
                            % (k, before[k], after[k]))
    try:
        from lxml import etree as _e
        _e.fromstring(etree.tostring(root))
    except Exception as e:                          # noqa: BLE001
        problems.append("XML 不可解析：%s" % e)
    # 域三态配平
    types = [e.get(qn("fldCharType")) for e in root.iter(qn("fldChar"))]
    b_, s_, e_ = (types.count("begin"), types.count("separate"),
                  types.count("end"))
    if not (b_ == s_ == e_):
        problems.append("域三态不平衡 begin=%d separate=%d end=%d" % (b_, s_, e_))
    if cm.dirty:
        ids = [c.get(qn("id")) for c in cm.root.iter(qn("comment"))]
        if len(ids) != len(set(ids)):
            problems.append("批注 id 重复")
    return problems


def _md5(path):
    import hashlib
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def _enable_track_changes(workdir):
    """在 settings.xml 里打开 trackChanges，让后续编辑也继续留痕。"""
    p = os.path.join(workdir, "word", "settings.xml")
    if not os.path.exists(p):
        return False
    t = etree.parse(p)
    r = t.getroot()
    if r.find(qn("trackChanges")) is not None:
        return True
    e = etree.Element(qn("trackChanges"))
    want = _SETTINGS_ORDER.index("trackChanges")
    idx = len(r)
    for i, c in enumerate(r):
        name = etree.QName(c).localname
        if name in _SETTINGS_ORDER and _SETTINGS_ORDER.index(name) > want:
            idx = i
            break
    r.insert(idx, e)
    t.write(p, xml_declaration=True, encoding="UTF-8", standalone=True)
    return True


def _pack(workdir, out):
    from .docxio import pack
    return pack(workdir, out)


# ------------------------------------------------------------------ 回执
def write_receipt(receipt, path):
    s = receipt["summary"]
    L = ["# 改动回执 —— %s" % os.path.basename(
        receipt.get("target", {}).get("path", "") or "文档"),
         "",
         "- 依据：改动事项单（target md5 `%s`）"
         % receipt.get("target", {}).get("md5", "?"),
         "- 备份：%s　md5 `%s`" % (receipt.get("backup"),
                                  receipt.get("backup_md5")),
         "- 结果：已改 %d / 未勾选未改 %d / 失败 %d / 转交他处 %d"
         % (s["changed"], s["pending"], s["failed"], s.get("other", 0)),
         "",
         "## 一、已改（%d）" % s["changed"], "",
         "| CR-ID | 位置 | 改前 | 改后 |", "|---|---|---|---|"]
    for it in receipt["changed"]:
        w = it.get("where") or {}
        L.append("| %s | %s · 段 %s | %s | %s |"
                 % (it["id"], w.get("section", ""), w.get("para", ""),
                    _md(it.get("before")), _md(it.get("after"))))
    L += ["", "## 二、未勾选未改（%d）" % s["pending"], "",
          "| CR-ID | 档 | 位置 | 摘要 |", "|---|---|---|---|"]
    for it in receipt["pending"]:
        w = it.get("where") or {}
        L.append("| %s | %s | %s · 段 %s | %s |"
                 % (it["id"], it.get("grade", ""), w.get("section", ""),
                    w.get("para", ""), _md(it.get("before"))))
    if receipt.get("others"):
        L += ["", "## 三、转交他处（%d）—— 非文字类，不归本执行器" 
              % len(receipt["others"]), "",
              "| CR-ID | executor | 位置 | 摘要 |", "|---|---|---|---|"]
        for it in receipt["others"]:
            w = it.get("where") or {}
            L.append("| %s | %s | %s · 段 %s | %s |"
                     % (it["id"], it.get("executor", ""), w.get("section", ""),
                        w.get("para", ""),
                        _md(it.get("before") or it.get("after"))))
    L += ["", "## 四、失败（%d）" % s["failed"], "",
          "| CR-ID | 原因 |", "|---|---|"]
    for f in receipt["failed"]:
        L.append("| %s | %s |" % (f["id"], f["reason"]))
    L += ["", "## 五、安全", "",
          "- 备份路径与 md5（上）",
          "- 改动以**修订模式**呈现，每条挂批注说明原因",
          "- 域未重算：需在 Word 里 `Ctrl+A` → `F9` 一次"]
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path


def _md(s, n=46):
    s = re.sub(r"\s+", " ", (s or "")).strip()
    return s if len(s) <= n else s[:n] + "…"
