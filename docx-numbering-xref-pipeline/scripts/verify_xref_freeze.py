# -*- coding: utf-8 -*-
"""
verify_xref_freeze.py —— xref_freeze 冻结产物的独立验收。

用法：
    python verify_xref_freeze.py <frozen.docx>
        [--base <原稿.docx>] [--receipt <冻结回执.json>]
        [--allow-remaining N]     # 允许残留 N 个 REF（默认 0）

检查项（全 PASS 才算过；任一 FAIL → exit 2）：
    1. fldChar 三态配平（begin == separate == end）
    2. REF 指令残留 == 允许值（默认 0）
    3. bookmarkStart / bookmarkEnd 配平，且每个书签 start 有同 id 的 end
    4. 可见文本无「错误！未定义书签」/「Error! Reference source not found」
    5. [--base] w:p / w:tbl / w:drawing / 书签计数与原稿一致
    6. [--base + --receipt] 可见文本 == 原稿可见文本按回执逐条替换的期望值
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
TAGP = re.compile(r'<w:p(?: [^>]*)?/>|<w:p [^>]*>|<w:p>')
BMS = re.compile(r'<w:bookmarkStart\b[^>]*/>')
BME = re.compile(r'<w:bookmarkEnd\b[^>]*/>')
ATTR = re.compile(r'w:(id|name)="([^"]*)"')
RE_REF = re.compile(r"(?:^|\s)REF\s+([^\s\\]+)")
ERR_TEXTS = ("错误！未定义书签", "错误!未定义书签",
             "Error! Reference source not found")


def read_doc_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def visible(x):
    return "".join(m.group(1) for m in WT.finditer(x))


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 6
    frozen = argv[1]
    base = receipt_path = None
    allow = 0
    for i, a in enumerate(argv):
        if a == "--base" and i + 1 < len(argv):
            base = argv[i + 1]
        elif a == "--receipt" and i + 1 < len(argv):
            receipt_path = argv[i + 1]
        elif a == "--allow-remaining" and i + 1 < len(argv):
            allow = int(argv[i + 1])

    y = read_doc_xml(frozen)
    fails = []

    def chk(name, cond, detail=""):
        print(("PASS  " if cond else "FAIL  ") + name +
              ("   " + detail if detail else ""))
        if not cond:
            fails.append(name)

    # 1. fldChar 三态配平
    kinds = {"begin": 0, "separate": 0, "end": 0}
    for m in FC.finditer(y):
        kinds[m.group(1)] += 1
    chk("1. fldChar 三态配平 (%d/%d/%d)"
        % (kinds["begin"], kinds["separate"], kinds["end"]),
        kinds["begin"] == kinds["separate"] == kinds["end"])

    # 2. REF 残留
    n_ref = sum(1 for m in INSTR.finditer(y) if RE_REF.search(m.group(1)))
    chk("2. REF 指令残留 %d（允许 %d）" % (n_ref, allow), n_ref == allow)

    # 3. 书签配平
    starts, ends = {}, set()
    for m in BMS.finditer(y):
        attrs = dict(ATTR.findall(m.group(0)))
        if "id" in attrs:
            starts[attrs["id"]] = attrs.get("name", "")
    for m in BME.finditer(y):
        attrs = dict(ATTR.findall(m.group(0)))
        if "id" in attrs:
            ends.add(attrs["id"])
    orphan = [starts[i] for i in starts if i not in ends]
    chk("3. 书签 start/end 配平（%d 对，孤儿 %d）"
        % (len(starts), len(orphan)), not orphan,
        "孤儿: %s" % orphan[:5] if orphan else "")

    # 4. 错误文本
    vy = visible(y)
    hit = [t for t in ERR_TEXTS if t in vy]
    chk("4. 无未定义书签错误文本", not hit, str(hit))

    # 5. 与基线的结构比对
    if base:
        x = read_doc_xml(base)
        for t in ("<w:tbl>", "<w:drawing>", "<w:bookmarkStart",
                  "<w:bookmarkEnd"):
            chk("5. %s 计数一致（%d）" % (t, x.count(t)),
                x.count(t) == y.count(t))
        chk("5. w:p 计数一致（%d）" % len(TAGP.findall(x)),
            len(TAGP.findall(x)) == len(TAGP.findall(y)))

        # 6. 可见文本 == 回执期望（按回执记录的**坐标**重建，不按文本匹配——
        #    题注里的「图 5-5」和域缓存同文，按文本会误伤）
        if receipt_path:
            r = json.load(open(receipt_path, encoding="utf-8"))
            pieces = [(m.start(1), m.end(1), m.group(1))
                      for m in WT.finditer(x)]
            drop = set()
            insert = {}          # pos -> 新文本（首个缓存 piece 处）
            override = {}        # (s, e) -> 清理后文本
            for it in r.get("items", []):
                cps = it.get("cached_pieces") or []
                for (s, e) in cps:
                    drop.add((s, e))
                if cps:
                    insert[cps[0][0]] = it["replacement"]
                for (s, e, new) in it.get("cleanups", []):
                    override[(s, e)] = new
            out = []
            for (ps, pe, pt) in pieces:
                if (ps, pe) in override:
                    pt = override[(ps, pe)]
                if ps in insert:
                    out.append(insert[ps])
                if (ps, pe) in drop:
                    continue
                out.append(pt)
            expect = "".join(out)
            ok = (vy == expect)
            detail = ""
            if not ok:
                i = next((k for k, (a, b) in enumerate(zip(vy, expect))
                          if a != b), min(len(vy), len(expect)))
                detail = "首差@%d：%r vs %r" % (i, vy[i:i + 20],
                                                expect[i:i + 20])
            chk("6. 可见文本 == 回执期望（%d 项替换）" % len(r.get("items", [])),
                ok, detail)
        else:
            print("WARN  6. 未给 --receipt，跳过可见文本比对")
    else:
        print("WARN  5/6. 未给 --base，跳过基线比对")

    print()
    if fails:
        print("结果：FAIL x%d -> %s" % (len(fails), fails))
        return 2
    print("结果：全绿")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
