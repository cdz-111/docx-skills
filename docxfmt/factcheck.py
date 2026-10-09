# -*- coding: utf-8 -*-
"""
factcheck.py —— S5b 事实核查的**可执行部分**。

协议见 `docx-unify-pipeline/references/fact-check-protocol.md`（唯一权威）。
本模块只做协议里的「第 2 节 抽取」与「第 5/6 节 落档」——
**检索与判定必须由调用方用 `WebSearch` 真查**，本模块绝不猜、绝不臆造来源。

流水线：
    ① `extract(docx)`            正文 → 原子断言候选清单（json + md）
    ② 调用方逐条联网核 → 产出 verdicts.json
    ③ `to_crs(claims, verdicts)`  → 事实核查 CRS（交由 crs_merge.py 合并）
    ④ `render_md(...)`            → 人读核查表

四态之外还有 `out_of_scope`（小范围事件）/ `unverifiable`（主观表述），
这两类**进表但不进 CRS**（协议 §4）。
"""
import json
import os
import re

from . import qn
from .docxio import Docx, in_table, p_text, read_style_names, unpack
from .crs import CrsBuilder

# ---- 命中模式（顺序即优先级）
# ⚠️ 关键词必须**具体**。踩过的坑：`发布|实施|落地` 这类宽词会把
#    「该方法的落地路径」「实施任务卸载」「获得性能奖励」当成事实断言 ——
#    实测一篇技术报告能报出 8 条全是误报。所以：
#      * award 必须落到**具体奖名**（"奖"后面不能跟"励"，否则"奖励"会中招）
#      * news / award 还必须再有**年份或机构/载体锚点**才算（见 _anchored）
KIND_PATTERNS = [
    ("award", re.compile(
        r"(荣获|获得|摘得|斩获|评为|被评为|授予|入选|获)"
        r"[^，。；、]{0,20}?"
        r"(一等奖|二等奖|三等奖|特等奖|金奖|银奖|铜奖|"
        r"优秀奖|提名奖|科技进步奖|自然科学奖|技术发明奖|科学技术奖|"
        r"发明奖|创新奖|优秀博士(?:学位)?论文|优秀论文|"
        r"荣誉称号|先进个人|先进集体|教学名师|领军人才)")),
    ("official", re.compile(
        r"(国务院|教育部|工业和信息化部|工信部|科学技术部|科技部|"
        r"国家发展改革委|发改委|财政部|人力资源社会保障部|知识产权局|"
        r"国家标准|行业标准|地方标准|国标委|"
        r"〔\s*\d{4}\s*〕|GB/?T\s*\d+|ZL\s*\d{5,}|CN\s*\d{6,})")),
    ("news", re.compile(
        r"(正式发布|重磅发布|对外发布|获批|批准|认定为|印发|出台|"
        r"发布|公布|施行|上线|开源|奠基|揭牌)")),
]
# 事实锚点：年份 或 机构/载体。没有锚点的一律当技术描述，不是可核查的外部事实。
RE_ENTITY = re.compile(
    r"(大学|学院|研究院|研究所|学会|协会|委员会|"
    r"教育部|工业和信息化部|工信部|科学技术部|科技部|国务院|"
    r"国家发展改革委|发改委|财政部|人力资源社会保障部|"
    r"国家标准|行业标准|地方标准|国标委|知识产权局|专利局|"
    r"[^\s，。；、]{2,12}(会议|大会|峰会|论坛|期刊|学报|出版社|基金会))")
# 小范围事件（协议 §1「不做」）—— 命中即 out_of_scope
RE_LOCAL = re.compile(
    r"(课题组内部|团队内部|内部测试|内部评审|内部讨论|我校内部|"
    r"本课题组|本团队|本实验室|本项目组|内部验收|组会|例会)")
# 主观表述（不可证伪）
RE_SUBJECTIVE = re.compile(r"(效果显著|意义重大|具有重要意义|表现优异|性能优越|"
                           r"前景广阔|国内领先|国际先进)")
RE_YEAR = re.compile(r"(19|20)\d{2}\s*年")
RE_SENT = re.compile(r"[^。；！？]+[。；！？]?")
CJK = r"\u4e00-\u9fff"


def _sentences(text):
    out = []
    for m in RE_SENT.finditer(text):
        s = m.group(0).strip()
        if s:
            out.append(s)
    return out


def _keywords(s, n=6):
    """粗抽检索词：实体串（≥2 汉字）、数字、英文词。"""
    toks = re.findall(r"[%s]{2,12}|[A-Za-z][A-Za-z0-9.\-]{2,}|"
                      r"20\d{2}\s*年(?:\s*\d{1,2}\s*月)?|第\s*\d+\s*号" % CJK, s)
    seen, out = set(), []
    for t in toks:
        t = t.strip()
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out[:n]


def _kind_of(s):
    for kind, rx in KIND_PATTERNS:
        if rx.search(s):
            return kind, rx.search(s).group(0)
    return None, None


def extract(docx, spec=None, include_local=False, log=print):
    """正文 → 原子断言候选。返回 {claims, stats, docx}。

    ⛔ 范围判定**必须走文档树**，不能靠"扫到『附录』就停"：
       目录区里也有 `附录A：发表论文列表158` 这种条目，
       一旦命中就会把后面 4000 多段整片跳过（实测只剩 6 段）。
       → 只处理树里 type == 'paragraph' 的节点，目录/表格/题注/公式/
         封面天然被排除；附录起点取**真正的 appendix 级标题**。
    """
    from .recognize import parse

    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_facts")
    unpack(docx, workdir)
    doc = Docx(workdir)
    root = doc.root
    sid2name = read_style_names(workdir)
    skip_kw = ((spec or {}).get("numbering") or {}).get("skip_keywords")
    tree, _index, _cover = parse(root, sid2name, skip_keywords=skip_kw)
    nodes = _all(tree)

    ap_start = None
    for n in nodes:
        if n.type == "heading" and n.sec_kind == "appendix" \
                and n.meta.get("idx") is not None:
            if ap_start is None or n.meta["idx"] < ap_start:
                ap_start = n.meta["idx"]

    claims = []
    stats = {"scanned_paras": 0, "sentences": 0, "hit": 0,
             "out_of_scope": 0, "subjective": 0, "no_anchor": 0,
             "appendix_from": ap_start}

    for n in nodes:
        if n.type != "paragraph":
            continue                       # 目录/表格/题注/公式/封面全排除
        i = n.meta.get("idx")
        if i is None:
            continue
        if ap_start is not None and i >= ap_start:
            continue
        t = (n.text or "").strip()
        if len(t) < 8:
            continue
        stats["scanned_paras"] += 1
        sec = _section_of(n)
        for s in _sentences(t):
            stats["sentences"] += 1
            if not (8 <= len(s) <= 200):
                continue
            kind, hit = _kind_of(s)
            if not kind:
                continue
            # award / news 必须有「年份」或「机构/载体」锚点；
            # 没有锚点的只是技术描述，不是可核查的外部事实。
            if kind in ("award", "news") and not (RE_YEAR.search(s)
                                                  or RE_ENTITY.search(s)):
                stats["no_anchor"] += 1
                continue
            if RE_LOCAL.search(s):
                stats["out_of_scope"] += 1
                claims.append(_claim(len(claims) + 1, s, "local", i, sec, hit))
                continue
            if RE_SUBJECTIVE.search(s) and not RE_YEAR.search(s):
                stats["subjective"] += 1
                claims.append(_claim(len(claims) + 1, s, "subjective", i, sec,
                                     hit))
                continue
            claims.append(_claim(len(claims) + 1, s, kind, i, sec, hit))
            stats["hit"] += 1

    stats["total"] = len(claims)
    import shutil
    shutil.rmtree(workdir, ignore_errors=True)
    log("[facts] 扫 %d 段 / %d 句 → 断言候选 %d 条"
        "（事实 %d / 小范围 %d / 主观 %d / 无锚点丢弃 %d；附录起自段 %s）"
        % (stats["scanned_paras"], stats["sentences"], stats["total"],
           stats["hit"], stats["out_of_scope"], stats["subjective"],
           stats["no_anchor"], ap_start if ap_start is not None else "无"))
    return {"docx": os.path.abspath(docx), "claims": claims, "stats": stats}


def _all(tree):
    """把文档树摊平成节点列表（前序）。"""
    out = []

    def rec(n):
        out.append(n)
        for c in n.children:
            rec(c)
    rec(tree)
    return out


def _section_of(node):
    """沿 parent 链上溯最近的 heading 文本作为章节名。"""
    p = node
    sec = None
    while p is not None:
        if p.type == "heading":
            sec = p.text
        p = getattr(p, "parent", None)
    return (re.sub(r"\s+", " ", sec or "")[:40] or "（无章节）")


def _claim(n, text, kind, para, sec, hit):
    return {"id": "F-%03d" % n, "kind": kind, "claim": text,
            "hit": hit, "keywords": _keywords(text),
            "where": {"section": sec, "para": para, "anchor": text[:40],
                      "count": 1},
            "status": "pending",       # pending | checked | out_of_scope
            "verdict": None,           # supported / partially_supported / ...
            "sources": [], "after": None, "note": ""}


def _sid(p, sid2name):
    ppr = p.find(qn("pPr"))
    if ppr is None:
        return None
    e = ppr.find(qn("pStyle"))
    return sid2name.get(e.get(qn("val"))) if e is not None else None


# ------------------------------------------------------------------ 回填判定
def apply_verdicts(claims, verdicts):
    """verdicts: {"F-001": {"verdict": "...", "sources": [...], "after": "...",
                            "note": "..."}}"""
    by = verdicts.get("items", verdicts) if isinstance(verdicts, dict) else {}
    n = 0
    for c in claims:
        v = by.get(c["id"])
        if not v:
            continue
        c["verdict"] = v.get("verdict")
        c["sources"] = v.get("sources") or []
        c["after"] = v.get("after")
        c["note"] = v.get("note", "")
        c["status"] = ("out_of_scope" if c["verdict"] in
                       ("out_of_scope", "unverifiable") else "checked")
        n += 1
    return n


GRADE_OF = {"partially_supported": ("S", "mid"), "contradicted": ("S", "mid"),
            "insufficient": ("C", "low")}


def to_crs(docx, claims, spec=None, sources_note=""):
    """把已判定的断言落成事实核查 CRS（只有 S/C 两类进单）。"""
    b = CrsBuilder(docx, sources=("fact-check",),
                   baseline_note="事实核查单；目标指纹变化即作废")
    n = 0
    for c in claims:
        v = c.get("verdict")
        if v not in GRADE_OF:
            continue
        grade, risk = GRADE_OF[v]
        src = c.get("sources") or []
        basis = "；".join(src) if src else "未找到公开来源（协议 §4：不是错）"
        if c.get("note"):
            basis += "｜" + c["note"]
        b.add(grade, "事实", c["where"], c["claim"],
              c.get("after") or "",
              basis, risk=risk,
              executor="text_replace" if v in ("partially_supported",
                                               "contradicted") else "manual",
              notes="四格核对（主体/时间/事件/载体）")
        n += 1
    return b.build(), n


# ------------------------------------------------------------------ 表格（人读）
def render_md(claims, path, doc_name="文档", date_str=""):
    buckets = {"supported": [], "partially_supported": [], "contradicted": [],
               "insufficient": [], "out_of_scope": [], "unverifiable": [],
               "pending": [], None: []}
    for c in claims:
        buckets.setdefault(c.get("verdict"), []).append(c)
    need = buckets.get("partially_supported", []) + \
        buckets.get("contradicted", [])
    L = ["# 事实核查表 —— %s%s" % (doc_name,
                                  ("（访问日期 %s）" % date_str) if date_str
                                  else ""), "",
         "共 %d 条：需改 %d / 未找到公开来源 %d / 已证实 %d / 超范围 %d / 未核 %d"
         % (len(claims), len(need), len(buckets.get("insufficient", [])),
            len(buckets.get("supported", [])),
            len(buckets.get("out_of_scope", [])) +
            len(buckets.get("unverifiable", [])),
            len(buckets.get("pending", [])) + len(buckets.get(None, []))), ""]

    L += ["## 一、需要改（%d）" % len(need), "",
          "| 编号 | 位置 | 原文断言 | 判定 | 权威来源 | 建议改为 |",
          "|---|---|---|---|---|---|"]
    for c in need:
        L.append("| %s | %s · 段 %s | %s | %s | %s | %s |"
                 % (c["id"], c["where"]["section"], c["where"]["para"],
                    _md(c["claim"], 60), c["verdict"],
                    "；".join(c.get("sources") or []) or "—",
                    _md(c.get("after") or "—", 40)))
    L += ["", "## 二、未找到公开来源（%d）—— 请老板确认"
          % len(buckets.get("insufficient", [])), "",
          "| 编号 | 位置 | 原文断言 | 说明 |", "|---|---|---|---|"]
    for c in buckets.get("insufficient", []):
        L.append("| %s | %s · 段 %s | %s | %s |"
                 % (c["id"], c["where"]["section"], c["where"]["para"],
                    _md(c["claim"], 60), c.get("note") or "官网未公开到该年度"))
    L += ["", "## 三、已证实（%d）—— 无需改动，留档备查"
          % len(buckets.get("supported", [])), "",
          "| 编号 | 断言 | 来源 |", "|---|---|---|"]
    for c in buckets.get("supported", []):
        L.append("| %s | %s | %s |" % (c["id"], _md(c["claim"], 70),
                                       "；".join(c.get("sources") or [])))
    L += ["", "## 四、超范围 / 不可证伪（%d）"
          % (len(buckets.get("out_of_scope", [])) +
             len(buckets.get("unverifiable", []))), "",
          "| 编号 | 断言 | 原因 |", "|---|---|---|"]
    for c in buckets.get("out_of_scope", []) + buckets.get("unverifiable", []):
        L.append("| %s | %s | %s |" % (c["id"], _md(c["claim"], 60),
                                       "小范围事件，无公开来源"
                                       if c["kind"] == "local" else "主观表述"))
    pend = buckets.get("pending", []) + buckets.get(None, [])
    if pend:
        L += ["", "## 五、尚未核查（%d）" % len(pend), "",
              "| 编号 | 类型 | 断言 | 检索词 |", "|---|---|---|---|"]
        for c in pend:
            L.append("| %s | %s | %s | %s |"
                     % (c["id"], c["kind"], _md(c["claim"], 70),
                        " ".join(c.get("keywords") or [])[:60]))
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path


def _md(s, n):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[:n] + "…"


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(obj, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    return path
