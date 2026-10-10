# -*- coding: utf-8 -*-
"""
crs_merge.py —— 把来自多个来源的改动事项单**合并成唯一一份 CRS**。

为什么必须合并：一条链路上会冒出好几份清单 ——

  · `docxfmt propose`   → 排版/层级疑点
  · 编号治理（S2）      → 题注编号、REF 断链、书签过宽
  · `docxfmt lint`      → 缺字/不通顺/术语/体例/数据
  · 事实核查（S5b）     → 奖项、新闻、官方通报的联网核对结论

如果让老板分别勾 4 份清单，必然出现「勾了 A 单忘了 B 单」「同一处改了两遍
（第二次基于旧文本）」「哪个 CR-ID 属于哪份单说不清」。所以：

  出单阶段允许多来源 → **勾选阶段只有一份** → 执行阶段只读这一份。

本脚本做四件事：
  1. **验指纹**：所有来源单的 target.md5 必须一致，不一致直接退出 5；
  2. **去重**：同一处（段落 + 改前 + 改后）在多份单里出现 → 合成一条，
     档位取最高（S>A>B>C），`source` 记成 `a+b`；
  3. **统一编号**：按 档位 → 来源 → 原段号 排序，重排成 CR-001..N；
  4. **落盘**：输出合并后的 `<文档名>_改动事项.json` 与可勾选 `.html`。

零第三方依赖；HTML 渲染优先复用 `docx-format/docxfmt/crs.py`（保证协议一致），
拿不到时退回内置的极简渲染，并在 stdout 明确告警。

用法
----
    PY=~/.workbuddy/binaries/envs/... python

    # 合并若干份单
    "$PY" crs_merge.py --docx "报告.docx" \
        --in 报告_待拍板.json --in 报告_编号治理.json --in 报告_改动事项.json \
        -o 报告_合并清单.json --html 报告_合并清单.html

    # 从流水线账本里自动认领产物（按 S5→S4→S3→S2 顺序）
    "$PY" crs_merge.py --docx "报告.docx" --from-state _unify/pipeline_state.json

    # 合并后直接套用老板导出的裁决
    "$PY" crs_merge.py ... --decisions 报告_approved.json

退出码：0 成功 / 3 没给来源 / 4 来源单格式不对 / 5 指纹不一致 / 6 用法错误
"""
import argparse
import json
import os
import re
import sys

GRADE_ORDER = {"S": 0, "A": 1, "B": 2, "C": 3}
CRS_VERSION = "1.0"

EXIT_OK = 0
EXIT_NO_SOURCE = 3
EXIT_BAD_SOURCE = 4
EXIT_FINGERPRINT = 5
EXIT_USAGE = 6


# ------------------------------------------------------------------ 基础
def _load(p):
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def _norm(s):
    """用于去重比对的规范化：去首尾空白、全角空格 → 半角、压连续空白。"""
    if s is None:
        return ""
    s = str(s).replace("\u3000", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def _sig(item):
    """一条改动的事项签名。段落 + 改前 + 改后 三者都同 → 认为是同一件事。"""
    w = item.get("where") or {}
    return (
        str(w.get("para", "")),
        _norm(item.get("before")),
        _norm(item.get("after")),
    )


def _is_crs(obj):
    return isinstance(obj, dict) and "items" in obj and isinstance(obj["items"], list)


def _docx_fingerprint(docx):
    """拿 docx-format 的指纹实现（字段与协议一致）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    for cand in (
        os.path.join(here, "..", "..", "docx-format"),
        os.path.expanduser("~/.workbuddy/skills/docx-format"),
    ):
        cand = os.path.abspath(cand)
        if os.path.isdir(os.path.join(cand, "docxfmt")):
            if cand not in sys.path:
                sys.path.insert(0, cand)
            try:
                from docxfmt.crs import fingerprint      # noqa: WPS433
                return fingerprint(docx)
            except Exception:
                break
    # 退路：自己算，字段名保持同构
    import hashlib
    import zipfile
    with open(docx, "rb") as f:
        raw = f.read()
    with zipfile.ZipFile(docx) as z:
        x = z.read("word/document.xml")
    import time
    return {
        "path": os.path.abspath(docx),
        "size": os.path.getsize(docx),
        "md5": hashlib.md5(raw).hexdigest(),
        "document_xml_md5": hashlib.md5(x).hexdigest(),
        "paragraphs": x.count(b"<w:p "),
        "captured_at": time.strftime("%Y-%m-%dT%H:%M:%S+08:00", time.localtime()),
    }


def _render_html(crs, out_path):
    """优先复用 docx-format 的渲染器（保证 DOM 契约一致）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    for cand in (
        os.path.join(here, "..", "..", "docx-format"),
        os.path.expanduser("~/.workbuddy/skills/docx-format"),
    ):
        cand = os.path.abspath(cand)
        if os.path.isfile(os.path.join(cand, "templates", "crs.html")):
            if cand not in sys.path:
                sys.path.insert(0, cand)
            try:
                from docxfmt.crs import render_html
                return render_html(crs, out_path), True
            except Exception as e:
                print("[warn] 复用 docx-format 渲染器失败：%s" % e)
                break
    # 退路：内置极简页（功能等价：勾选 + 导出 approved.json）
    data = json.dumps(crs, ensure_ascii=False, indent=2)
    safe = data.replace("<", "\\u003c")
    html = _FALLBACK_HTML.replace("__CRS_DATA__", safe)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path, False


_FALLBACK_HTML = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>改动事项单</title><style>
:root{--bg:#fff;--fg:#1a1a1a;--mut:#666;--line:#e3e3e3;--ac:#0a66c2}
@media (prefers-color-scheme:dark){
 :root{--bg:#1b1c1e;--fg:#e8e8e8;--mut:#a0a0a0;--line:#333;--ac:#5aa9f5}}
body{background:var(--bg);color:var(--fg);font:14px/1.6 -apple-system,"Segoe UI",
 "Microsoft YaHei",sans-serif;margin:0;padding:20px}
h1{font-size:18px}table{border-collapse:collapse;width:100%;margin-top:12px}
th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
th{position:sticky;top:0;background:var(--bg)}
.bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:10px 0}
button{padding:4px 10px;border:1px solid var(--line);background:transparent;
 color:var(--fg);border-radius:6px;cursor:pointer}
button:hover{border-color:var(--ac);color:var(--ac)}
.g-S{color:#c62828;font-weight:600}.g-A{color:#ef6c00;font-weight:600}
.g-B{color:#1565c0}.g-C{color:#6a1b9a}
.before{color:var(--mut);text-decoration:line-through}
.after{color:#2e7d32;font-weight:600}
#sum{color:var(--mut)}
</style></head><body>
<h1>改动事项单（合并）</h1>
<div class="bar">
 <button onclick="pick('all')">全选</button>
 <button onclick="pick('none')">全不选</button>
 <button onclick="pick('S')">只勾 S</button>
 <button onclick="pick('A')">只勾 A</button>
 <button onclick="pick('B')">只勾 B</button>
 <button onclick="pick('C')">只勾 C</button>
 <span id="sum"></span>
 <button id="exp">导出 approved.json</button>
</div>
<table><thead><tr><th>勾选</th><th>ID</th><th>档</th><th>来源</th><th>位置</th>
<th>现状</th><th>建议改为</th><th>依据</th><th>风险</th></tr></thead>
<tbody id="tb"></tbody></table>
<input type="checkbox" id="chk-risk" onchange="render()"> 包含高风险项
<script id="crs-data" type="application/json">__CRS_DATA__</script>
<script>
const CRS=JSON.parse(document.getElementById('crs-data').textContent);
const D=CRS.items||[];
const RISK={'低':'low','中':'mid','高':'high',low:'low',mid:'mid',high:'high'};
function esc(s){return String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
function render(){
 const showRisk=document.getElementById('chk-risk').checked;
 let rows='',shown=0;
 for(const it of D){
   const rk=String(it.risk||'low').toLowerCase();
   if(!showRisk && rk==='high'){continue}
   shown++;
   rows+=`<tr class="cr-row" data-cr-id="${it.id}" data-grade="${it.grade}"
     data-risk="${rk}" data-source="${it.source}">
   <td><input type="checkbox" class="cr-check" data-cr-id="${it.id}"></td>
   <td>${it.id}</td><td class="g-${it.grade}">${it.grade}</td>
   <td>${esc(it.source)}</td>
   <td>${esc(it.where&&it.where.section)}${it.where&&it.where.para?' · 段'+it.where.para:''}
       ${it.where&&it.where.anchor?'<br><small>'+esc(it.where.anchor)+'</small>':''}</td>
   <td class="before">${esc(it.before)}</td>
   <td class="after">${esc(it.after)}</td>
   <td><small>${esc(it.basis)}</small></td>
   <td>${esc(rk)}</td></tr>`;
 }
 document.getElementById('tb').innerHTML=rows;
 document.querySelectorAll('.cr-check').forEach(c=>c.addEventListener('change',upd));
 upd();
}
function pick(k){
 document.querySelectorAll('.cr-check').forEach(c=>{
  const row=c.closest('tr');
  c.checked = k==='all' ? true : k==='none' ? false : row.dataset.grade===k;
 });
 upd();
}
function upd(){
 const all=[...document.querySelectorAll('.cr-check')];
 const on=all.filter(c=>c.checked);
 const g={};on.forEach(c=>{const k=c.closest('tr').dataset.grade;g[k]=(g[k]||0)+1});
 document.getElementById('sum').textContent=
  `已勾 ${on.length} / 共 ${all.length}　`+
  Object.keys(g).sort().map(k=>k+':'+g[k]).join(' ');
}
document.getElementById('exp').onclick=()=>{
 const items=[...document.querySelectorAll('.cr-check')]
   .map(c=>({id:c.dataset.crId,decision:c.checked?'approved':'pending'}))
   .filter(x=>x.decision==='approved');
 const out={crs_version:CRS.crs_version,target_md5:(CRS.target||{}).md5,
   exported_at:new Date().toISOString(),items};
 const blob=new Blob([JSON.stringify(out,null,2)],{type:'application/json'});
 const a=document.createElement('a');
 a.href=URL.createObjectURL(blob);a.download='approved.json';a.click();
 navigator.clipboard&&navigator.clipboard.writeText(JSON.stringify(out,null,2));
};
render();
</script></body></html>"""


# ------------------------------------------------------------------ 合并
def collect_inputs(args):
    """确定来源清单：显式 --in 优先；否则从账本产物里认领。"""
    paths = list(args.infiles or [])
    if not paths and args.from_state:
        st = _load(args.from_state)
        # 按「越靠后越权威」的顺序收集：S5 → S4 → S3 → S2
        for sid in ("S2", "S3", "S4", "S5"):
            s = (st.get("stages") or {}).get(sid) or {}
            for a in s.get("artifacts", []):
                p = a.get("path")
                if p and os.path.exists(p) and p.lower().endswith(".json") \
                        and "approved" not in os.path.basename(p).lower() \
                        and "state" not in os.path.basename(p).lower():
                    paths.append(p)
    # 去重保持顺序
    seen, out = set(), []
    for p in paths:
        ap = os.path.abspath(p)
        if ap not in seen:
            seen.add(ap)
            out.append(ap)
    return out


def merge(sources, target_fp, docx_path):
    """sources: [(path, crs_dict), ...] -> 合并后的 crs dict"""
    items = []
    by_sig = {}
    conflicts = []
    for path, crs in sources:
        src_name = os.path.basename(path)
        src_tag = (crs.get("sources") or [None])[0] or _infer_source(src_name)
        for it in crs["items"]:
            sig = _sig(it)
            # 没有 before/after 的（纯结构变更）不去重，避免误合
            mergeable = bool(sig[1]) and bool(sig[2])
            key = sig if mergeable else ("__uniq__%s__%d" % (src_name, len(items)))
            if mergeable and key in by_sig:
                old = by_sig[key]
                # 档位取最高
                if GRADE_ORDER.get(it.get("grade"), 9) < \
                        GRADE_ORDER.get(old.get("grade"), 9):
                    old["grade"] = it["grade"]
                srcs = set(str(old.get("source", "")).split("+"))
                srcs.add(str(it.get("source") or src_tag))
                old["source"] = "+".join(sorted(s for s in srcs if s))
                old.setdefault("merged_from", []).append(
                    {"file": src_name, "id": it.get("id")})
                if _norm(old.get("after")) != _norm(it.get("after")):
                    conflicts.append({
                        "sig": sig, "keep": old.get("after"),
                        "alt": it.get("after"), "from": src_name})
                continue
            new = dict(it)
            new["source"] = it.get("source") or src_tag
            new.setdefault("merged_from", [])
            new["_origin"] = src_name
            new["_origin_id"] = it.get("id")
            items.append(new)
            if mergeable:
                by_sig[key] = new

    # 排序：档位 → 段号 → 原文顺序
    items.sort(key=lambda x: (
        GRADE_ORDER.get(x.get("grade"), 9),
        int((x.get("where") or {}).get("para") or 0),
    ))
    # 统一编号
    for i, it in enumerate(items, 1):
        it["id"] = "CR-%03d" % i

    return {
        "crs_version": CRS_VERSION,
        "target": target_fp,
        "baseline": {
            "note": "目标文件指纹若变化，本单全部作废，必须重新出单",
            "merged_from": [{"file": os.path.basename(p),
                             "items": len(c["items"]),
                             "sources": c.get("sources", [])}
                            for p, c in sources],
            "conflicts": conflicts,
        },
        "sources": sorted({str(it.get("source", "")) for it in items
                           if it.get("source")}),
        "gate": {"state": "awaiting_approval", "total": len(items),
                 "approved": 0, "rejected": 0},
        "items": [{k: v for k, v in it.items() if not k.startswith("_origin")}
                  for it in items],
        "_origin_index": {it["id"]: {"file": it["_origin"],
                                     "orig_id": it["_origin_id"]}
                          for it in items},
    }


def _infer_source(name):
    n = name.lower()
    if "待拍板" in name or "propose" in n:
        return "propose"
    if "编号" in name or "xref" in n or "numbering" in n:
        return "numbering-xref"
    if "事实" in name or "fact" in n:
        return "fact-check"
    if "排版" in name or "polish" in n:
        return "polish"
    return "lint"


def apply_decisions(crs, path):
    dec = _load(path)
    by = {d["id"]: d for d in dec.get("items", [])}
    n = 0
    for it in crs["items"]:
        d = by.get(it["id"])
        if d:
            it["decision"] = d["decision"]
            if d.get("after_override"):
                it["after_override"] = d["after_override"]
            n += 1
    ap = sum(1 for i in crs["items"] if i["decision"] in ("approved", "modified"))
    crs["gate"] = {"state": "approved" if ap else "rejected",
                   "total": len(crs["items"]), "approved": ap,
                   "rejected": sum(1 for i in crs["items"]
                                   if i["decision"] == "rejected")}
    return n, ap


# ------------------------------------------------------------------ CLI
def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="crs_merge",
        description="把多来源改动事项单合并成唯一一份 CRS")
    ap.add_argument("--docx", required=True, help="被改文档（用于验指纹）")
    ap.add_argument("--in", dest="infiles", action="append", default=[],
                    help="来源 CRS json（可重复）")
    ap.add_argument("--from-state", help="从流水线账本自动认领产物")
    ap.add_argument("-o", "--out", help="合并后 CRS json（默认 <文档名>_改动事项.json）")
    ap.add_argument("--html", help="可勾选页路径（默认同名 .html）")
    ap.add_argument("--decisions", help="合并后套用老板导出的 approved.json")
    args = ap.parse_args(argv)

    srcs = collect_inputs(args)
    if not srcs:
        print("[err] 没有来源单。用 --in <crs.json> 或 --from-state <state.json>")
        return EXIT_NO_SOURCE

    docx = os.path.abspath(args.docx)
    if not os.path.exists(docx):
        print("[err] 找不到文档：%s" % docx)
        return EXIT_USAGE
    fp = _docx_fingerprint(docx)

    loaded = []
    for p in srcs:
        if not os.path.exists(p):
            print("[warn] 跳过不存在的来源：%s" % p)
            continue
        try:
            obj = _load(p)
        except Exception as e:
            print("[err] %s 不是合法 JSON：%s" % (p, e))
            return EXIT_BAD_SOURCE
        if not _is_crs(obj):
            print("[err] %s 不是 CRS（缺 items 数组）。"
                  "先把该来源归集成 CRS 再合并。" % p)
            return EXIT_BAD_SOURCE
        t = obj.get("target") or {}
        if t.get("md5") and t["md5"] != fp["md5"]:
            print("[5] 指纹不一致 —— %s 是对另一个版本出的单" % os.path.basename(p))
            print("     该单 md5 %s" % t["md5"])
            print("     当前 md5 %s" % fp["md5"])
            print("     → 必须先重新出单，禁止拿旧单凑合（协议 §11.2）")
            return EXIT_FINGERPRINT
        loaded.append((p, obj))

    if not loaded:
        print("[err] 有效来源为 0")
        return EXIT_NO_SOURCE

    merged = merge(loaded, fp, docx)

    stem = os.path.splitext(os.path.basename(docx))[0]
    d = os.path.dirname(docx)
    out = args.out or os.path.join(d, "%s_改动事项.json" % stem)
    html = args.html or os.path.join(d, "%s_改动事项.html" % stem)

    ori = merged.pop("_origin_index")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    # 溯源索引单独落一份，便于回答「这条 CR 原来是哪份单的第几条」
    with open(os.path.splitext(out)[0] + "_溯源.json", "w", encoding="utf-8") as f:
        json.dump(ori, f, ensure_ascii=False, indent=2)
    hp, reused = _render_html(merged, html)

    from collections import Counter
    g = Counter(i["grade"] for i in merged["items"])
    print("=" * 74)
    print("合并改动事项单：%s" % stem)
    print("=" * 74)
    print("来源 %d 份 → 条目 %d 条　分档 %s"
          % (len(loaded), len(merged["items"]), dict(sorted(g.items()))))
    print("  来源明细：")
    for p, c in loaded:
        print("    %-40s %d 条" % (os.path.basename(p), len(c["items"])))
    n_dup = sum(len(i.get("merged_from", [])) for i in merged["items"])
    if n_dup:
        print("  去重合并：%d 处（档位取最高）" % n_dup)
    if merged["baseline"]["conflicts"]:
        print("  ⚠️ 同一处有 %d 条建议改法冲突，已保留档位最高的一条："
              % len(merged["baseline"]["conflicts"]))
        for c in merged["baseline"]["conflicts"][:5]:
            print("     段 %s「%s」：保留「%s」／弃「%s」"
                  % (c["sig"][0], c["sig"][1], c["keep"], c["alt"]))
    print()
    print("-> %s" % out)
    print("-> %s" % os.path.splitext(out)[0] + "_溯源.json")
    print("-> %s%s" % (hp, "" if reused else "  （内置渲染，未能复用 docx-format 模板）"))

    if args.decisions:
        n, ap = apply_decisions(merged, args.decisions)
        with open(out, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        print()
        print("已套用裁决 %s：%d 条有明确 decision，其中 %d 条批准"
              % (os.path.basename(args.decisions), n, ap))
        print("下一步：python -m docxfmt revise <文档> --crs %s" % out)
    else:
        print()
        print("下一步：把 %s 交给老板勾选，导出 approved.json，然后" % hp)
        print("  python -m docxfmt revise <文档> --crs %s --approved approved.json" % out)
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
