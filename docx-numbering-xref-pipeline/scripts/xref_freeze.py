# -*- coding: utf-8 -*-
"""
xref_freeze.py —— 定稿冻结：把 REF 交叉引用域整体替换为纯文本。

为什么需要：终态渲染管线（转 HTML / 转 PDF / 第三方解析器 / 在线预览系统）
**不认 REF 域**——要么读到域指令本身，要么读到过期缓存，渲染就会出错。
所以版本定稿后，交叉引用应一律冻结成正文文本。
编辑期仍应保留真 REF 域（可 F9 自动更新）——冻结是**单向、终态**动作，
冻结之后不许再 F9、不许再跑 fields。

用法：
    审计（只读，出冻结事项单）：
    python xref_freeze.py audit <in.docx> -o 冻结事项.json [--all]

    执行（按事项单冻结，先备份！）：
    python xref_freeze.py apply <in.docx> <out.docx> --plan 冻结事项.json
        --style dead|indicator|mixed [--ids XF-001,XF-005]
        [--spec spec.json] [--receipt 回执.json]

体例三档（apply --style）：
    dead       死格式：如「图 5-5」「（2-7）」——所见即所得，最稳（推荐）
    indicator  指示词：如「上图」「下表」——按书签与引用的相对位置判定
    mixed      每个标签首次出现用死格式，后续用指示词

体例取值优先级：命令行 --style > spec 的 xref_freeze.style > dead。
（把 --spec 指到 docx-format/spec/rd-program.json 即读其 xref_freeze.style）

取值来源：语义重算（书签包裹的编号区文本）优先，域缓存兜底；
两者不一致会在审计单里标 agree=false——那说明缓存过期，
**建议先刷新域（fields）再冻结**，让缓存与语义一致。

审计范围：默认只收「题注类」REF（语义/缓存以 图/表/式/公式/附图/附表 开头）；
--all 收全部 REF 域。

不变量（任一破坏 → 不落盘，exit 2）：
    w:p / w:tbl / w:drawing / bookmarkStart / bookmarkEnd 计数不变；
    fldChar 计数恰好下降 3×N；REF 指令数恰好下降 N；
    可见文本只允许按事项单预期变化（全文 w:t 逐段拼接比对）。

退出码：0 成功；2 断言失败；4 无可执行项；5 md5 不符；6 用法错误。
"""
import hashlib
import io
import json
import re
import sys
import zipfile

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                  errors="replace")
except Exception:                                       # noqa: BLE001
    pass

FC = re.compile(r'<w:fldChar[^>]*w:fldCharType="(begin|separate|end)"[^>]*/>')
WT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')
INSTR = re.compile(r'<w:instrText[^>]*>([^<]*)</w:instrText>')
TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>|</w:p>')
BMS = re.compile(r'<w:bookmarkStart\b[^>]*/>')
BME = re.compile(r'<w:bookmarkEnd\b[^>]*/>')
ATTR = re.compile(r'w:(id|name)="([^"]*)"')
RE_REF = re.compile(r"(?:^|\s)REF\s+([^\s\\]+)")

# span 内出现这些元素 → 该项跳过（宁可不改也不改坏）
BLACKLIST = ("<w:drawing", "<w:pict", "<w:tab", "<w:br",
             "<w:footnoteReference", "<w:commentReference", "<w:object")

ERR_TEXTS = ("错误！未定义书签", "错误!未定义书签",
             "Error! Reference source not found")


def die(msg, code=2):
    print("[FAIL] " + msg)
    sys.exit(code)


def escapes(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


STYLES = ("dead", "indicator", "mixed")


def style_from_spec(path):
    """从 spec json 读 xref_freeze.style（默认体例）。读不到返回 None。"""
    if not path:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception as e:                               # noqa: BLE001
        print("[warn] 读 spec 失败（%r），忽略" % (repr(e)[:80],))
        return None
    blk = d.get("xref_freeze") or {}
    if isinstance(blk, str):
        st = blk
    else:
        st = blk.get("style")
    if st in STYLES:
        return st
    if st:
        print("[warn] spec 里 xref_freeze.style=%r 不是 %s 之一，忽略"
              % (st, "/".join(STYLES)))
    return None


def run_open_at(x, pos):
    s = x.rfind("<w:r", 0, pos)
    while s != -1 and x[s + 4:s + 5] not in (">", " "):
        s = x.rfind("<w:r", 0, s)
    return s


def run_close_at(x, pos):
    e = x.find("</w:r>", pos)
    return e + 6 if e != -1 else -1


def paras(s):
    r, d, st = [], 0, None
    for m in TAGP.finditer(s):
        t = m.group(0)
        if t.endswith("/>"):
            if d == 0:
                r.append((m.start(), m.end()))
        elif t == "</w:p>":
            d -= 1
            if d <= 0:
                if st is not None:
                    r.append((st, m.end()))
                st = None
                d = 0
        else:
            if d == 0:
                st = m.start()
            d += 1
    return r


def para_index(P, pos):
    for i, (a, b) in enumerate(P):
        if a <= pos < b:
            return i
    return -1


def fields_of(x):
    """域栈扫描 → [(begin_start, sep_start|None, end_end)]，按位置排序。"""
    stack, out = [], []
    for m in FC.finditer(x):
        k = m.group(1)
        if k == "begin":
            stack.append([m.start(), None])
        elif k == "separate":
            if stack and stack[-1][1] is None:
                stack[-1][1] = m.start()
        else:
            if stack:
                b, s = stack.pop()
                out.append((b, s, m.end()))
    out.sort()
    return out


def bookmark_map(x):
    """→ (name -> (pos, id), id -> end_pos)。"""
    by_name, end_by_id = {}, {}
    for m in BMS.finditer(x):
        attrs = dict(ATTR.findall(m.group(0)))
        if "name" in attrs:
            by_name[attrs["name"]] = (m.start(), attrs.get("id"))
    for m in BME.finditer(x):
        attrs = dict(ATTR.findall(m.group(0)))
        if "id" in attrs:
            end_by_id[attrs["id"]] = m.start()
    return by_name, end_by_id


def text_between(x, a, b):
    return "".join(m.group(1) for m in WT.finditer(x, a, b))


def kind_of(label):
    m = re.match(r"^\s*(附?)\s*(图|表|式|公式)", label or "")
    if not m:
        return None
    return "式" if m.group(2) == "公式" else m.group(2)


# ------------------------------------------------------------------ audit
def audit(x, want_all):
    P = paras(x)
    bms, bme = bookmark_map(x)
    items, skipped = [], []
    n_ref = 0
    for (b, s, e) in fields_of(x):
        joined = "".join(INSTR.findall(x[b:(s if s is not None else e)]))
        m = RE_REF.search(joined)
        if not m:
            continue
        n_ref += 1
        bmname = m.group(1)
        reason = None
        span_start = run_open_at(x, b)
        span_end = run_close_at(x, e)
        if span_start == -1 or span_end == -1:
            reason = "无法定位 run 边界"
        elif any(t in x[span_start:span_end] for t in BLACKLIST):
            reason = "span 含 drawing/tab/br 等非域内容"
        elif len(FC.findall(x[span_start:span_end])) != 3:
            reason = "span 内 fldChar ≠ 3（疑似嵌套域）"
        elif s is None:
            reason = "无 separate（无结果区）"
        elif any(mm.start() >= s for mm in INSTR.finditer(x, s, e)):
            reason = "结果区含 instrText"
        if reason:
            skipped.append({"instr": joined.strip(), "para": para_index(P, b),
                            "reason": reason})
            continue
        cached = text_between(x, s, e)
        bm = bms.get(bmname)
        semantic = ""
        if bm:
            bm_pos, bm_id = bm
            bm_end = bme.get(bm_id)
            if bm_end is not None and bm_end > bm_pos:
                semantic = text_between(x, bm_pos, bm_end)
        label = semantic or cached
        kind = kind_of(semantic) or kind_of(cached)
        if not want_all and kind is None:
            continue
        items.append({
            "id": "XF-%03d" % (len(items) + 1),
            "para": para_index(P, b),
            "instr": joined.strip(),
            "bookmark": bmname,
            "bookmark_found": bm is not None,
            "kind": kind,
            "cached": cached,
            "semantic": semantic,
            "agree": (not semantic or not cached
                      or semantic.replace(" ", "") == cached.replace(" ", "")),
            "above": bool(bm and bm[0] < span_start),
            "dead": label,
            "indicator": ("上" if (bm and bm[0] < span_start) else "下")
                         + kind if (kind and bm) else None,
            "span": [span_start, span_end],
            "sep": s, "end": e,
            "include": True,
        })
    return items, skipped, n_ref


# ------------------------------------------------------------------ apply
def build_edits(x, items, style):
    """→ (edits, applied, fallbacks)。edits = [(start, end, xml_or_text)]。"""
    seen = set()
    edits, applied, fallbacks = [], [], []
    for it in items:
        if not it.get("include", True):
            continue
        mode = style
        if style == "mixed":
            mode = "dead" if it["dead"] not in seen else "indicator"
            seen.add(it["dead"])
        if mode == "indicator" and not it.get("indicator"):
            fallbacks.append((it["id"], "indicator 不可用（kind/书签缺失）→ dead"))
            mode = "dead"
        text = it["dead"] if mode == "dead" else it["indicator"]
        span_s, span_e = it["span"]

        # rPr：优先抄结果区第一个 run；无结果 run 则抄 begin run
        cached_pieces = [[mm.start(1), mm.end(1)]
                         for mm in WT.finditer(x, it["sep"], it["end"])]
        ref_run = None
        for mm in WT.finditer(x, it["sep"], it["end"]):
            ref_run = run_open_at(x, mm.start())
            break
        if ref_run is None:
            ref_run = span_s
        run_xml = x[ref_run:run_close_at(x, ref_run)]
        mr = re.match(r"^<w:r(?: [^>]*)?>(<w:rPr>.*?</w:rPr>)?", run_xml, re.S)
        rpr = mr.group(1) if (mr and mr.group(1)) else ""

        rep = '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (
            rpr, escapes(text))
        edits.append((span_s, span_e, rep))
        rec = {"id": it["id"], "para": it["para"],
               "cached": it["cached"], "replacement": text, "mode": mode,
               "span": [span_s, span_e], "cached_pieces": cached_pieces,
               "cleanups": []}
        applied.append(rec)

        # 冻结后文本与前后字之间的多余空格清理（两种体例通用）
        # ① 句读前不留空格：「…图 5-5 。」→「…图 5-5。」「…下表 。」→「…下表。」
        nm = WT.search(x, span_e)
        if nm:
            t = nm.group(1)
            nt = t
            if mode == "indicator":
                # ② 指示词与「所示」之间不留空格：「上图 所示」→「上图所示」
                nt = re.sub(r"^[ \t\u3000]+(?=所示)", "", nt)
            nt = re.sub(r"^[ \t\u3000]+(?=[。，、；：！？》」』）])", "", nt)
            if nt != t:
                edits.append((nm.start(1), nm.end(1), nt))
                rec["cleanups"].append([nm.start(1), nm.end(1), nt])
        if mode == "indicator":
            # ③ 指示词前不留空格：「如 上图」→「如上图」「见 下表」→「见下表」
            pm = None
            for mm in WT.finditer(x, 0, span_s):
                pm = mm
            if pm and re.search(r"(如|见|参见)[ \t\u3000]+$", pm.group(1)):
                new = re.sub(r"((?:如|见|参见))[ \t\u3000]+$", r"\1",
                             pm.group(1))
                edits.append((pm.start(1), pm.end(1), new))
                rec["cleanups"].append([pm.start(1), pm.end(1), new])
    return edits, applied, fallbacks


def apply_edits(x, edits):
    edits = sorted(edits, key=lambda e: (e[0], e[1]))
    for k in range(1, len(edits)):
        if edits[k][0] < edits[k - 1][1]:
            die("编辑区间重叠: %s / %s" % (edits[k - 1][:2], edits[k][:2]))
    for (s, e, rep) in reversed(edits):
        x = x[:s] + rep + x[e:]
    return x


def visible(x):
    return "".join(m.group(1) for m in WT.finditer(x))


def expected_from_edits(x, edits):
    """按编辑清单推演冻结后的可见文本（与 apply_edits 同一坐标系）。"""
    pieces = [(m.start(1), m.end(1), m.group(1)) for m in WT.finditer(x)]
    cleanup = {}                    # (ps, pe) -> 新文本（text 级编辑）
    spans = []                      # [s, e, 新文本, 已发射]
    for (s, e, rep) in edits:
        if rep.startswith("<w:r"):
            spans.append([s, e, "".join(WT.findall(rep)), False])
        else:
            cleanup[(s, e)] = rep
    spans.sort()
    out = []
    si = 0
    for (ps, pe, pt) in pieces:
        if (ps, pe) in cleanup:
            pt = cleanup[(ps, pe)]
        while si < len(spans) and spans[si][1] <= ps:
            if not spans[si][3]:
                out.append(spans[si][2])
            si += 1
        if si < len(spans) and spans[si][0] <= ps < spans[si][1]:
            if not spans[si][3]:
                out.append(spans[si][2])
                spans[si][3] = True
            continue
        out.append(pt)
    while si < len(spans):
        if not spans[si][3]:
            out.append(spans[si][2])
        si += 1
    return "".join(out)


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 6
    cmd = argv[1]

    if cmd == "audit":
        if len(argv) < 3:
            print(__doc__)
            return 6
        inp = argv[2]
        out = None
        want_all = "--all" in argv
        for i, a in enumerate(argv):
            if a == "-o" and i + 1 < len(argv):
                out = argv[i + 1]
        if not out:
            out = inp.rsplit(".", 1)[0] + "_冻结事项.json"
        raw = open(inp, "rb").read()
        with zipfile.ZipFile(inp) as z:
            x = z.read("word/document.xml").decode("utf-8")
        items, skipped, n_ref = audit(x, want_all)
        plan = {
            "schema": 1, "tool": "xref_freeze", "mode": "audit",
            "docx": inp, "check_md5": hashlib.md5(raw).hexdigest(),
            "want_all": want_all,
            "total_ref_fields": n_ref, "caption_like": len(items),
            "items": items, "skipped": skipped,
            "note": "逐条把 include 改为 false 可剔除；apply 需 --style。",
        }
        with open(out, "w", encoding="utf-8") as f:
            json.dump(plan, f, ensure_ascii=False, indent=1)
        n_dis = sum(1 for it in items if not it["agree"])
        print("REF 域共 %d 个；可冻结 %d 个；跳过 %d 个"
              % (n_ref, len(items), len(skipped)))
        if n_dis:
            print("⚠ %d 个缓存与语义不一致（建议先刷新域再冻结）：" % n_dis)
            for it in items:
                if not it["agree"]:
                    print("   %s 段p%d 缓存=%r 语义=%r"
                          % (it["id"], it["para"], it["cached"],
                             it["semantic"]))
        print("-> %s" % out)
        print("下一步：python xref_freeze.py apply <in.docx> <out.docx> "
              "--plan %s --style dead|indicator|mixed（或 --spec <spec.json>）"
              % out)
        return 0

    if cmd == "apply":
        if len(argv) < 4:
            print(__doc__)
            return 6
        inp, outp = argv[2], argv[3]
        opt = {"--plan": None, "--style": None, "--ids": None,
               "--receipt": None, "--spec": None}
        for i, a in enumerate(argv):
            if a in opt and i + 1 < len(argv):
                opt[a] = argv[i + 1]
        if not opt["--plan"]:
            print("apply 需要 --plan")
            return 6
        # 体例取值优先级：命令行 --style > spec 的 xref_freeze.style > dead
        if opt["--style"] is None:
            opt["--style"] = style_from_spec(opt["--spec"]) or "dead"
            if opt["--spec"] and opt["--style"] != "dead":
                print("[style] 取自 spec：%s" % opt["--style"])
        if opt["--style"] not in STYLES:
            print("--style 只能是 dead / indicator / mixed")
            return 6
        plan = json.load(open(opt["--plan"], encoding="utf-8"))
        got = md5(inp)
        if got != plan["check_md5"]:
            die("源文件 md5 不符：得到 %s，期望 %s（文档在出单后被改过，"
                "重新 audit）" % (got, plan["check_md5"]), 5)
        with zipfile.ZipFile(inp) as z:
            x = z.read("word/document.xml").decode("utf-8")

        items = plan["items"]
        if opt["--ids"]:
            ids = set(opt["--ids"].split(","))
            items = [it for it in items if it["id"] in ids]

        # 出单后文档未变（md5 已验）→ 重算一次校验 plan 仍然成立
        re_items, _sk, _tot = audit(x, plan.get("want_all"))
        if len(re_items) < len(items):
            die("重算可冻结 REF 只有 %d 个 < 计划 %d 个，中止"
                % (len(re_items), len(items)))

        style = opt["--style"]
        edits, applied, fallbacks = build_edits(x, items, style)
        if not applied:
            print("无可执行项（include 全为 false 或被 --ids 过滤空）")
            return 4
        expect = expected_from_edits(x, edits)
        y = apply_edits(x, edits)

        # ---------------- 不变量 ----------------
        for t in ("<w:tbl>", "<w:drawing>", "<w:bookmarkStart",
                  "<w:bookmarkEnd"):
            if x.count(t) != y.count(t):
                die("不变量破坏: %s %d -> %d" % (t, x.count(t), y.count(t)))
        if len(TAGP.findall(x)) != len(TAGP.findall(y)):
            die("段落数变化")
        n_fld0, n_fld1 = x.count("<w:fldChar"), y.count("<w:fldChar")
        if n_fld0 - n_fld1 != 3 * len(applied):
            die("fldChar 下降 %d ≠ 3×%d" % (n_fld0 - n_fld1, len(applied)))
        n_ref0 = sum(1 for m in INSTR.finditer(x) if RE_REF.search(m.group(1)))
        n_ref1 = sum(1 for m in INSTR.finditer(y) if RE_REF.search(m.group(1)))
        if n_ref0 - n_ref1 != len(applied):
            die("REF 指令下降 %d ≠ %d" % (n_ref0 - n_ref1, len(applied)))
        vy = visible(y)
        if vy != expect:
            i = next((k for k, (a, b) in enumerate(zip(vy, expect))
                      if a != b), min(len(vy), len(expect)))
            die("可见文本与预期不符（首差位于 %d：%r vs %r）"
                % (i, vy[i:i + 20], expect[i:i + 20]))
        for t in ERR_TEXTS:
            if t in vy:
                die("冻结后出现 %r" % t)

        # ---------------- 打包（其余 part 字节透传） ----------------
        with zipfile.ZipFile(inp) as zin:
            with zipfile.ZipFile(outp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        data = y.encode("utf-8")
                    zi = zipfile.ZipInfo(item.filename,
                                         date_time=item.date_time)
                    zi.compress_type = item.compress_type
                    zi.external_attr = item.external_attr
                    zout.writestr(zi, data)

        receipt = {
            "schema": 1, "tool": "xref_freeze", "in": inp, "out": outp,
            "style": style, "frozen": len(applied),
            "remaining_ref": n_ref1, "fallbacks": fallbacks,
            "items": applied,
            "invariants": {
                "fldChar": [n_fld0, n_fld1],
                "ref_instr": [n_ref0, n_ref1],
                "w:p": [len(TAGP.findall(x)), len(TAGP.findall(y))],
                "bookmark": [x.count("<w:bookmarkStart"),
                             y.count("<w:bookmarkStart")],
            },
        }
        rp = opt["--receipt"] or outp.rsplit(".", 1)[0] + "_冻结回执.json"
        with open(rp, "w", encoding="utf-8") as f:
            json.dump(receipt, f, ensure_ascii=False, indent=1)
        print("冻结 %d 处（style=%s）；REF 指令 %d -> %d"
              % (len(applied), style, n_ref0, n_ref1))
        for fb in fallbacks:
            print("  [fallback] %s %s" % fb)
        print("-> %s" % outp)
        print("-> %s" % rp)
        print("OK")
        return 0

    print(__doc__)
    return 6


if __name__ == "__main__":
    sys.exit(main(sys.argv))
