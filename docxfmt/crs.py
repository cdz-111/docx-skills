# -*- coding: utf-8 -*-
"""
crs.py —— 改动事项单（CRS）的构造 / 落盘 / 渲染 / 裁决读取。

协议细则见 `references/change-request-sheet.md`（该文件是唯一权威）。
本模块是**唯一的实现**：lint / polish / propose / 编排层都从这里拿 CRS 读写能力，
不要各写一套。
"""
import hashlib
import json
import os
import re
import time
import zipfile

CRS_VERSION = "1.0"
SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CRS_TEMPLATE = os.path.join(SKILL_ROOT, "templates", "crs.html")

GRADES = ("S", "A", "B", "C")
GRADE_LABEL = {"S": "S 必改", "A": "A 应改", "B": "B 建议", "C": "C 待核"}
DECISIONS = ("pending", "approved", "rejected", "modified", "closed")


# ------------------------------------------------------------------ 指纹
def fingerprint(docx):
    """出单时给被改文档留指纹。执行前必须复核（见协议 §9）。"""
    with open(docx, "rb") as f:
        raw = hashlib.md5(f.read()).hexdigest()
    with zipfile.ZipFile(docx) as z:
        x = z.read("word/document.xml")
    return {
        "path": os.path.abspath(docx),
        "size": os.path.getsize(docx),
        "md5": raw,
        "document_xml_md5": hashlib.md5(x).hexdigest(),
        "paragraphs": x.count(b"<w:p "),
        "captured_at": _now(),
    }


def _now():
    return time.strftime("%Y-%m-%dT%H:%M:%S+08:00", time.localtime())


def resolve_out(base, docx, suffix):
    """<文档名>_<suffix> 落在被改文档同目录。suffix 形如 '改动事项'。"""
    d = os.path.dirname(os.path.abspath(docx))
    stem = os.path.splitext(os.path.basename(docx))[0]
    return os.path.join(d, "%s_%s" % (stem, suffix))


# ------------------------------------------------------------------ 构造
class CrsBuilder:
    """累积 CRS 条目，自动分配稳定 CR-ID。"""

    def __init__(self, target_docx, sources=("lint",), baseline_note=""):
        self.target = fingerprint(target_docx) if target_docx else {}
        self.sources = list(sources)
        self.baseline = {"note": baseline_note or
                         "目标文件指纹若变化，本单全部作废，必须重新出单"}
        self.items = []
        self._n = 0

    def add(self, grade, category, where, before, after, basis,
            risk="low", impact="无", source=None, executor="text_replace",
            notes=""):
        """where 是 dict：{section, para, anchor, count}。"""
        self._n += 1
        it = {
            "id": "CR-%03d" % self._n,
            "grade": grade,
            "source": source or (self.sources[0] if self.sources else "lint"),
            "category": category,
            "where": _norm_where(where),
            "before": before,
            "after": after,
            "basis": basis,
            "risk": risk,
            "impact": impact,
            "executor": executor,
            "decision": "pending",
            "after_override": None,
            "notes": notes,
        }
        self.items.append(it)
        return it["id"]

    def add_raw(self, item):
        """直接塞一条（用于合并外部 CRS 后重新编号）。"""
        self._n += 1
        it = dict(item)
        it["id"] = "CR-%03d" % self._n
        it.setdefault("decision", "pending")
        it.setdefault("after_override", None)
        it.setdefault("notes", "")
        self.items.append(it)
        return it["id"]

    def build(self, gate_state="awaiting_approval"):
        n_ap = sum(1 for i in self.items if i["decision"] == "approved")
        n_rj = sum(1 for i in self.items if i["decision"] == "rejected")
        return {
            "crs_version": CRS_VERSION,
            "target": self.target,
            "baseline": self.baseline,
            "sources": self.sources,
            "gate": {"state": gate_state, "total": len(self.items),
                     "approved": n_ap, "rejected": n_rj},
            "items": self.items,
        }


def _norm_where(where):
    w = dict(where or {})
    w.setdefault("section", "")
    w.setdefault("para", 0)
    w.setdefault("anchor", "")
    w.setdefault("count", 1)
    return w


# ------------------------------------------------------------------ 落盘 / 渲染
def write_json(crs, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(crs, f, ensure_ascii=False, indent=2)
    return path


def render_html(crs, out_path, template=None):
    """把 CRS 注入可勾选页模板。数据里的 `<` 必须转义，否则会截断脚本。"""
    tpl_path = template or CRS_TEMPLATE
    with open(tpl_path, "r", encoding="utf-8") as f:
        tpl = f.read()
    data = json.dumps(crs, ensure_ascii=False, indent=2)
    safe = data.replace("<", "\\u003c").replace("\u2028", "\\u2028") \
                .replace("\u2029", "\\u2029")
    pat = re.compile(
        r'(<script id="crs-data" type="application/json">)(.*?)(</script>)',
        re.S)
    if not pat.search(tpl):
        raise SystemExit("[crs] 模板里找不到 <script id=\"crs-data\"> 数据块：%s"
                         % tpl_path)
    out = pat.sub(lambda m: m.group(1) + "\n" + safe + "\n" + m.group(3),
                  tpl, count=1)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(out)
    return out_path


def load(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def approved_items(crs, decisions_dir=None):
    """取要执行的项：decision=approved / modified。可按 _approved.json 覆盖。"""
    if decisions_dir:
        p = decisions_dir
        if os.path.isdir(p):
            p = os.path.join(p, os.path.basename(p))  # 不猜，要求直接给文件
        if os.path.exists(p):
            dec = load(p)
            by = {d["id"]: d for d in dec.get("items", [])}
            for it in crs["items"]:
                d = by.get(it["id"])
                if d:
                    it["decision"] = d["decision"]
                    if d.get("after_override"):
                        it["after_override"] = d["after_override"]
    out = []
    for it in crs["items"]:
        if it["decision"] in ("approved", "modified"):
            out.append(it)
    return out


# ------------------------------------------------------------------ 统计
def summarize(crs):
    from collections import Counter
    g = Counter(i["grade"] for i in crs["items"])
    d = Counter(i["decision"] for i in crs["items"])
    return {"total": len(crs["items"]),
            "by_grade": dict(g), "by_decision": dict(d)}
