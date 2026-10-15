# -*- coding: utf-8 -*-
"""
confirm.py —— 样式确认单（S1）。

用户给了模板 → 逐部分列出「当前样式 vs 目标样式」；
没给模板 → 用默认模板（spec/default.json）列出同一张表。
**只出单，不套用。** 用户确认/手改后，把导出的 `_样式确认单.json` 当 `--spec` 喂回 run。

产物：
    <文档名>_样式确认单.html   人看 + 手改 + 导出
    <文档名>_样式确认单.json   可直接 `--spec` 喂回（source.style_from = "inline"）
"""
import json
import os
import re
import zipfile
from collections import Counter

from lxml import etree

from . import qn
from .docxio import Docx, read_style_names, unpack
from .recognize import parse
from .spec import Spec
from .styles import NAME_ALIASES

# 部分定义：(键, 显示名, 对应 spec 角色, 级别)
PARTS = [
    ("cover",        "封面",           None,        0),
    ("toc_title",    "目录/清单标题",   "toc_title", 0),
    ("toc",          "目录条目",        "toc",       1),
    ("heading1",     "一级标题",        "heading",   1),
    ("heading2",     "二级标题",        "heading",   2),
    ("heading3",     "三级标题",        "heading",   3),
    ("heading4",     "四级及以下标题",   "heading",   4),
    ("paragraph",    "正文",           "paragraph", 0),
    ("caption",      "题注",           "caption",   0),
    ("image",        "图片",           "image",     0),
    ("table_text",   "表格表头",        "table_text", 0),
    ("table_body",   "表格正文",        "table_body", 0),
    ("formula",      "公式",           "formula",   0),
    ("formula_number", "公式编号",     "formula_number", 0),
    ("appendix",     "附录",           "appendix",  0),
    ("header",       "页眉",           "header",    0),
    ("footer",       "页脚",           "footer",    0),
    ("comment",      "批注",           None,        0),
]

GEOM_KEYS = [
    ("size", "字号 pt"), ("font", "字体"), ("bold", "加粗"),
    ("align", "对齐"), ("first_line", "首行缩进"),
    ("before", "段前"), ("after", "段后"), ("line", "行距"),
]


def _sty_name(p, sid2name):
    ppr = p.find(qn("pPr")) if p is not None else None
    if ppr is None:
        return None
    e = ppr.find(qn("pStyle"))
    if e is None:
        return None
    sid = e.get(qn("val"))
    return sid2name.get(sid, sid)


def _count_part(part_root, sid2name, kind, level=0):
    """统计某一部分的段落数 + 当前样式分布。"""
    cnt = Counter()
    n = 0
    for p in part_root.iter(qn("p")):
        if level and p is not None:
            pass
        n += 1
        cnt[_sty_name(p, sid2name) or "（无样式）"] += 1
    return n, cnt


def collect(docx, spec):
    """返回 (inventory, rows, resolved_spec)。"""
    workdir = os.path.join(os.path.dirname(os.path.abspath(docx)), "_confirm")
    unpack(docx, workdir)
    doc = Docx(workdir)
    sid2name = read_style_names(workdir)
    body = doc.body

    root, _index, cover_end = parse(doc.root, sid2name,
                                    skip_keywords=spec["numbering"]
                                    .get("skip_keywords"))
    # 按节点类型聚合段落
    by_type = Counter()
    style_by_type = {}
    for p in doc.root.iter(qn("p")):
        nm = _sty_name(p, sid2name) or "（无样式）"
        by_type["_all"] += 1

    # 用树判定类型（更准）
    def walk(n):
        if n.type != "document" and n.type != "empty":
            key = n.type if n.type != "heading" else "heading%d" % min(n.level, 4)
            by_type[key] += 1
            style_by_type.setdefault(key, Counter())[
                _sty_name(n.ref, sid2name) or "（无样式）"] += 1
        for c in n.children:
            walk(c)
    walk(root)

    # 页眉页脚 / 批注
    inv = _inventory(docx, workdir, doc, sid2name, by_type)

    rows = []
    smap = spec["styles"]["map"]
    geom = spec["styles"].get("geometry") or {}
    newst = spec["styles"].get("new_styles") or {}
    table = _existing_styles(workdir)

    for key, label, role, level in PARTS:
        if key in ("header", "footer", "comment"):
            n = inv.get(key + "_count", 0)
            styles = inv.get(key + "_styles", Counter())
            target = _target_style(smap, role, level)
            rows.append(_row(key, label, n, styles, target, None, table,
                             geom, newst))
            continue
        if key.startswith("heading"):
            n = by_type.get(key, 0)
            styles = style_by_type.get(key, Counter())
        else:
            n = by_type.get(key, 0)
            styles = style_by_type.get(key, Counter())
        target = _target_style(smap, role, level)
        rows.append(_row(key, label, n, styles, target, None, table,
                         geom, newst))

    resolved = _resolved_spec(spec)
    return inv, rows, resolved


def _target_style(smap, role, level):
    if role is None:
        return None
    v = smap.get(role)
    if isinstance(v, dict):
        return v.get(str(level)) or v.get(level)
    return v


def _row(key, label, n, styles, target, _unused, table, geom, newst):
    top = ", ".join("%s×%d" % (k, v) for k, v in styles.most_common(3)) or "—"
    g = {}
    if target:
        g = geom.get(target) or {}
    # 存在性判定要按别名（「正文」在英文模板里叫 Normal）
    found = None
    if target:
        cand = [target] + list(NAME_ALIASES.get(target, []))
        for c in cand:
            if c in table:
                found = c
                break
            if c.lower() in table:
                found = c
                break
    missing = bool(target) and found is None
    return {
        "key": key, "label": label, "count": n,
        "current": top,
        "target": target or "（不动）",
        "resolved_as": found,
        "geometry": {k: g.get(k) for k, _lbl in GEOM_KEYS if k in g},
        "missing": missing,
        "will_create": bool(target) and target in newst,
    }


def _existing_styles(workdir):
    p = os.path.join(workdir, "word", "styles.xml")
    t = etree.parse(p).getroot()
    out = set()
    for s in t.iter(qn("style")):
        nm = s.find(qn("name"))
        if nm is not None and nm.get(qn("val")):
            out.add(nm.get(qn("val")))
            out.add(nm.get(qn("val")).lower())
    return out


def _inventory(docx, workdir, doc, sid2name, by_type):
    inv = {}
    # 页眉页脚
    for kind, pref in (("header", "word/header"), ("footer", "word/footer")):
        cnt = 0
        styles = Counter()
        for fn in sorted(os.listdir(os.path.join(workdir, "word"))):
            if fn.startswith(kind) and fn.endswith(".xml"):
                cnt += 1
                r = etree.parse(os.path.join(workdir, "word", fn)).getroot()
                for p in r.iter(qn("p")):
                    styles[_sty_name(p, sid2name) or "（无样式）"] += 1
        inv[kind + "_count"] = cnt
        inv[kind + "_styles"] = styles
    # 批注
    cp = os.path.join(workdir, "word", "comments.xml")
    if os.path.exists(cp):
        r = etree.parse(cp).getroot()
        inv["comment_count"] = len(list(r.iter(qn("comment"))))
    else:
        inv["comment_count"] = 0
    inv["comment_styles"] = Counter()
    # 域 / 书签 / 节
    x = doc.root
    inv.update({
        "paragraphs": len(list(x.iter(qn("p")))),
        "tables": len(list(x.iter(qn("tbl")))),
        "drawings": len(list(x.iter(qn("drawing")))),
        "math": len(list(x.iter(qn("oMath")))) if False else
                len(list(x.iter("{%s}oMath" % _M))),
        "fields": len(list(x.iter(qn("fldChar")))),
        "bookmarks": len(list(x.iter(qn("bookmarkStart")))),
        "sections": len(list(x.iter(qn("sectPr")))),
        "headings": sum(v for k, v in by_type.items() if k.startswith("heading")),
        "captions": by_type.get("caption", 0),
    })
    inv["style_names"] = sorted({v for v in sid2name.values() if v})
    return inv


_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _resolved_spec(spec):
    """把生效的 spec 落成可直接 `--spec` 喂回的文件内容（inline）。"""
    out = json.loads(json.dumps(spec, ensure_ascii=False, default=str))
    out["source"] = {"style_from": "inline", "template": None,
                     "requirement": None}
    return out


# ------------------------------------------------------------------ HTML
def render_html(inv, rows, resolved, docx, out_path):
    stem = os.path.splitext(os.path.basename(docx))[0]
    data = json.dumps({"title": stem, "inventory": inv, "rows": rows,
                       "spec": resolved}, ensure_ascii=False)
    safe = data.replace("<", "\\u003c")
    html = (_HTML
            .replace("__TITLE__", stem)
            .replace("__DATA__", safe))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>样式确认单 —— __TITLE__</title>
<style>
:root{--bg:#fff;--fg:#1a1a1a;--mut:#666;--line:rgba(0,0,0,.12);--card:#f7f7f5;--acc:#185FA5;--warn:#a32d2d}
@media (prefers-color-scheme:dark){:root{--bg:#1c1c1e;--fg:#eee;--mut:#9a9a9a;--line:rgba(255,255,255,.16);--card:#26262a;--acc:#85b7eb;--warn:#f09595}}
*{box-sizing:border-box}
body{margin:0;padding:20px;background:var(--bg);color:var(--fg);
font:13px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}
h1{font-size:16px;font-weight:500;margin:0 0 4px}
.sub{color:var(--mut);font-size:12px;margin-bottom:14px}
.grid{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:16px}
.kv{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px 10px;font-size:12px}
.kv b{font-weight:500;color:var(--acc)}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th,td{border:1px solid var(--line);padding:5px 7px;text-align:left;vertical-align:top}
th{background:var(--card);font-weight:500;position:sticky;top:0}
td.miss{color:var(--warn)}
.geo{color:var(--mut);font-size:11.5px}
.note{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px;margin-top:16px;font-size:12px}
button{font:inherit;padding:6px 12px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--fg);cursor:pointer}
button:hover{border-color:var(--acc)}
.row{display:flex;gap:10px;align-items:center;margin-top:10px;flex-wrap:wrap}
input,select{font:inherit;padding:3px 6px;border-radius:6px;border:1px solid var(--line);background:var(--bg);color:var(--fg);width:80px}
label{font-size:12px;color:var(--mut)}
</style></head><body>
<h1>样式确认单</h1>
<div class="sub" id="sub"></div>
<div class="grid" id="inv"></div>
<table><thead><tr>
<th style="width:120px">部分</th><th style="width:60px">段落/数量</th>
<th>当前样式（Top3）</th><th style="width:130px">目标样式</th><th>目标几何</th>
</tr></thead><tbody id="tb"></tbody></table>
<div class="note">
<b>编号与开关</b>（改完点导出，把 <code>_样式确认单.json</code> 交给执行侧）
<div class="row">
<label>题注编号</label>
<select id="cap"><option value="global">连续 1（图1）</option><option value="chapter">分章 1-1（图1-1）</option></select>
<label>分隔符</label><input id="sep" style="width:40px">
<label>目录</label><select id="toc"><option value="1">要</option><option value="0">不要</option></select>
<label>插图清单</label><select id="fig"><option value="1">要</option><option value="0">不要</option></select>
<label>附表清单</label><select id="tab"><option value="1">要</option><option value="0">不要</option></select>
</div>
<div class="row">
<label>分页层级</label><input id="pb" style="width:60px" placeholder="如 1">
<label>表格边框</label><select id="tb2"><option value="auto">自动判定</option><option value="three_line">全部三线表</option><option value="full_grid">全部全框</option></select>
<label>中文↔数字/英文去空格</label><select id="sp"><option value="1">开</option><option value="0">关</option></select>
</div>
<div class="row">
<button id="exp">导出 _样式确认单.json</button>
<button id="copy">复制到剪贴板</button>
<span id="msg" class="sub" style="margin:0"></span>
</div>
</div>
<div class="note">⚠️ 这份单子只是<b>目标态声明</b>，还没有碰文档。确认/修改后把导出的 json 当 <code>--spec</code> 喂回去，才会开始统一格式（仍会先出改动事项单等你勾）。</div>
<script id="sg-data" type="application/json">__DATA__</script>
<script>
const D=JSON.parse(document.getElementById('sg-data').textContent);
document.getElementById('sub').textContent='目标文档：'+D.title;
const invNames={paragraphs:'段落',tables:'表格',drawings:'图片',math:'公式',fields:'域(fldChar)',bookmarks:'书签',sections:'节',headings:'标题',captions:'题注',comment_count:'批注',header_count:'页眉部件',footer_count:'页脚部件'};
document.getElementById('inv').innerHTML=Object.entries(invNames).map(([k,v])=>
  `<div class="kv">${v} <b>${D.inventory[k]??0}</b></div>`).join('');
const gk={size:'字号',font:'字体',bold:'加粗',align:'对齐',first_line:'缩进',before:'段前',after:'段后',line:'行距'};
document.getElementById('tb').innerHTML=D.rows.map(r=>{
 const geo=Object.entries(r.geometry||{}).map(([k,v])=>`${gk[k]||k}=${v}`).join(' · ')||'—';
 const tgt=r.missing?`${r.target} <span class="sub">(文档里没有${r.will_create?'，将新建':'，需先确认'})</span>`
   :(r.resolved_as&&r.resolved_as!==r.target?`${r.target} <span class="sub">→ 实际用「${r.resolved_as}」</span>`:r.target);
 return `<tr><td>${r.label}</td><td>${r.count}</td><td>${r.current}</td>
 <td class="${r.missing?'miss':''}">${tgt}</td><td class="geo">${geo}</td></tr>`}).join('');
const S=D.spec;
document.getElementById('cap').value=(S.numbering&&S.numbering.caption)||'global';
document.getElementById('sep').value=(S.numbering&&S.numbering.sep)||'-';
document.getElementById('toc').value=(S.toc&&S.toc.enabled)?'1':'0';
document.getElementById('fig').value=(S.figure_list&&S.figure_list.enabled)?'1':'0';
document.getElementById('tab').value=(S.table_list&&S.table_list.enabled)?'1':'0';
document.getElementById('pb').value=((S.pagebreak&&S.pagebreak.before_heading)||[1]).join(',');
document.getElementById('tb2').value=(S.tables&&S.tables.border_style)||'auto';
document.getElementById('sp').value=(S.spacing&&S.spacing.normalize_cjk_latin)?'1':'0';
function out(){const o=JSON.parse(JSON.stringify(S));
 o.numbering=o.numbering||{};o.numbering.caption=document.getElementById('cap').value;
 o.numbering.sep=document.getElementById('sep').value;
 o.toc=Object.assign(o.toc||{},{enabled:document.getElementById('toc').value==='1'});
 o.figure_list=Object.assign(o.figure_list||{},{enabled:document.getElementById('fig').value==='1'});
 o.table_list=Object.assign(o.table_list||{},{enabled:document.getElementById('tab').value==='1'});
 o.pagebreak=Object.assign(o.pagebreak||{},{before_heading:document.getElementById('pb').value.split(',').map(x=>parseInt(x)).filter(x=>!isNaN(x))});
 o.tables=Object.assign(o.tables||{},{border_style:document.getElementById('tb2').value});
 o.spacing=Object.assign(o.spacing||{},{normalize_cjk_latin:document.getElementById('sp').value==='1'});
 return o;}
function dl(name,text){const b=new Blob([text],{type:'application/json'});const a=document.createElement('a');
 a.href=URL.createObjectURL(b);a.download=name;a.click();URL.revokeObjectURL(a.href);}
document.getElementById('exp').onclick=()=>{dl('__TITLE___样式确认单.json',JSON.stringify(out(),null,2));msg.textContent='已导出';};
document.getElementById('copy').onclick=async()=>{try{await navigator.clipboard.writeText(JSON.stringify(out(),null,2));msg.textContent='已复制';}catch(e){msg.textContent='复制失败，请用导出';}};
</script></body></html>
"""


def _theme_replace(s):
    return s
