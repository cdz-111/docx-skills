# -*- coding: utf-8 -*-
"""
propose.py —— 把「需要人来（或 AI）拍板的地方」挑出来，带上下文。

分工
    工具负责**机械判定**（样式/outlineLvl/编号正则），这部分 100% 可靠；
    但总有工具判不出来、或者判得没把握的地方。
    这个模块把那些地方挑出来，附上前后文，供人/AI 做最终判断，
    判断结果写成一份「覆盖表」喂回 pipeline（--tree-in）。

为什么要这一层
    规则再完善也有边界：比如「总结」「后记」后面不跟小节号，结构上没信号；
    又比如自动列表编号的文档，标题正文是「引言」这种光秃秃的文字。
    这些地方工具**不该硬猜**，而是把上下文摊开让人拍板。
    拍板结果落成 JSON，下次可复用 —— 这就是「AI 参与判断」的落点。
"""
import json
import re

from . import qn
from .docxio import has_drawing, in_table, p_text
from .recognize import (_outline_lvl, _style_name, classify_one,
                        find_cover_end, parse)


def build_proposal(doc_root, sid2name, cover_end=None,
                   context=2, max_len=60, split=True):
    """产出「待拍板清单」。

    ⚠️ 必须与 pipeline 走**完全相同的前置处理**（拆段），
       否则两边的段落序号会错位，清单里报的段号对不上渲染时的段号。
    """
    n_split = 0
    if split:
        from .captions import split_caption_paras
        n_split = split_caption_paras(doc_root)

    paras = list(doc_root.iter(qn("p")))
    texts = [p_text(p).strip() for p in paras]
    if cover_end is None:
        cover_end = find_cover_end(paras, sid2name)

    # 先跑一遍现有规则，拿到工具的判断
    from .recognize import mark_sections
    tmp_root, index, _ = parse(doc_root, sid2name, cover_end=cover_end)

    tool = {}
    for n in _iter(tmp_root):
        if n.id:
            tool[n.id] = (n.type, n.level)

    items = []
    for i, p in enumerate(paras):
        t = texts[i]
        if not t or i < cover_end:
            continue
        ntype, lvl = tool.get("p%d" % i, ("", 0))
        nm = _style_name(p, sid2name)
        olv = _outline_lvl(p)
        in_tbl = in_table(p)

        reason = None
        if ntype == "heading":
            # 工具判成标题：标出「凭什么」
            if nm and ("heading" in nm.lower() or "标题" in nm):
                conf = "样式"
            elif olv is not None:
                conf = "outlineLvl=%s" % olv
            else:
                conf = "仅文本规则"          # 弱信号，值得复核
                reason = "工具凭编号正则判定，无样式佐证"
            # 连「有样式佐证」的标题也进清单：工具判成 heading 的段，
            # 该不该进目录、该是几级，仍需调用方确认。
            if reason is None:
                reason = "已判为标题，请复核层级与是否该进目录"
        elif ntype == "paragraph" and not in_tbl and len(t) <= 40:
            # 短段 + 无句末标点 + 不在表格里 -> 可能是漏掉的标题
            if not re.search(r"[。；：，、）】]$", t) and not t.startswith("图"):
                conf = "疑似"
                reason = "短段且无句末标点，可能是漏判的标题"
        if reason is None:
            continue

        items.append({
            "id": "p%d" % i,
            "idx": i,
            "text": t[:max_len],
            "len": len(t),
            "tool_type": ntype,
            "tool_level": lvl,
            "confidence": conf,
            "reason": reason,
            "style": nm,
            "outline_lvl": olv,
            "in_table": in_tbl,
            "has_drawing": has_drawing(p),
            "before": [{"id": "p%d" % j, "text": texts[j][:max_len]}
                       for j in range(max(0, i - context), i) if texts[j]],
            "after": [{"id": "p%d" % j, "text": texts[j][:max_len]}
                      for j in range(i + 1,
                                     min(len(paras), i + 1 + context))
                      if texts[j]],
        })
    return {"cover_end": cover_end, "total_paras": len(paras),
            "split_caption_paras": n_split, "items": items}


def _iter(root):
    out = []
    def rec(n):
        out.append(n)
        for c in n.children:
            rec(c)
    rec(root)
    return out


def apply_overrides(root, overrides):
    """把「人工/AI 拍板结果」应用到树上。

    overrides: {段落id: {"type": "heading", "level": 2}} 或
               {段落id: "heading2"}（简写）
    返回实际改动的条数。
    """
    byid = {}
    for n in _iter(root):
        if n.id:
            byid[n.id] = n
    n = 0
    for pid, spec in (overrides or {}).items():
        node = byid.get(pid)
        if node is None:
            continue
        if isinstance(spec, str):
            m = re.match(r"^heading\s*(\d)$", spec)
            if m:
                spec = {"type": "heading", "level": int(m.group(1))}
            else:
                spec = {"type": spec}
        t = spec.get("type")
        if t:
            node.type = t
        if "level" in spec:
            node.level = int(spec["level"])
        # 改完要重算 sec_kind
        if node.type == "heading" and node.level == 1:
            node.meta["manual"] = True
        n += 1
    return n


def to_report(proposal, out_path=None, limit=200):
    """人/AI 可读的报告。"""
    L = []
    L.append("=" * 74)
    L.append("待拍板清单 —— %d 项（全文 %d 段，封面区 [0,%d)）"
             % (len(proposal["items"]), proposal["total_paras"],
                proposal["cover_end"]))
    L.append("=" * 74)
    for it in proposal["items"][:limit]:
        L.append("")
        L.append("[%s] %s" % (it["id"], it["text"]))
        L.append("   工具判: %s%s   置信: %s"
                 % (it["tool_type"],
                    (" L%d" % it["tool_level"]) if it["tool_type"] == "heading"
                    else "", it["confidence"]))
        L.append("   原因: %s" % it["reason"])
        L.append("   样式: %s   outlineLvl: %s   表格内: %s"
                 % (it["style"] or "无", it["outline_lvl"], it["in_table"]))
        for b in it["before"]:
            L.append("   ↑ %s" % b["text"])
        for a in it["after"]:
            L.append("   ↓ %s" % a["text"])
    s = "\n".join(L)
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(s)
    return s
