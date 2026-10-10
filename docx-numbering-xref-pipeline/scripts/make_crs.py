# -*- coding: utf-8 -*-
"""
make_crs.py —— 把「编号与引用治理」的改动项清单，转成符合 CRS 协议的清单。

为什么需要它：本 skill 的审计脚本（`audit_captions.py` / `recon_hardref.py` …）
输出的是**人读报告**，不是 CRS。而编排链路要求 S2 的产出必须是
`<文档名>_编号治理.json`（见 SKILL.md §9.2 / §9.3）。
本脚本就是那一步的适配器 —— **格式与落盘都复用 `docx-format` 的 `crs` 模块**，
不另起一套（§不要做 最后一条）。

入参 `--items` 是一个**简单清单**（人手填或由审计脚本产出）：

    [
      {"grade": "S", "category": "引用",
       "where": {"section": "3.4.1", "para": 214,
                 "anchor": "如图 45 所示", "count": 1},
       "before": "如图 45 所示", "after": "如图 44 所示",
       "basis": "图实际编号为 44；书签 _Ref00A3 缓存 SEQ=44 / REF=45",
       "risk": "high", "executor": "fix_ref", "notes": ""}
    ]

也可以给 `{"items": [...]}`；`where` 里的 `section/para/anchor` 是**必填**，
段号必须能被独立复核（SKILL.md §9.3）。

用法：
    python make_crs.py --docx 报告.docx --items 改动项.json \
        -o 报告_编号治理.json --html 报告_编号治理.html
"""
import argparse
import io
import json
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                              errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
DFORMAT = os.path.abspath(os.path.join(HERE, "..", "..", "docx-format"))
if DFORMAT not in sys.path:
    sys.path.insert(0, DFORMAT)

try:
    from docxfmt import crs as CRS          # noqa: E402
except Exception as e:                       # noqa: BLE001
    print("[err] 载入 docx-format 的 crs 模块失败：%r" % (e,))
    print("      期望位置：%s" % DFORMAT)
    sys.exit(1)

REQUIRED = ("grade", "category", "where")
GRADES = ("S", "A", "B", "C")


def load_items(path):
    with open(path, "r", encoding="utf-8") as f:
        d = json.load(f)
    if isinstance(d, dict):
        d = d.get("items") or d.get("list") or []
    if not isinstance(d, list):
        raise SystemExit("[err] --items 必须是数组，或 {'items': [...]}")
    return d


def validate(items):
    bad = []
    for k, it in enumerate(items, 1):
        for f in REQUIRED:
            if not it.get(f):
                bad.append("第 %d 条缺字段 `%s`" % (k, f))
        w = it.get("where") or {}
        if not (w.get("section") and w.get("anchor")):
            bad.append("第 %d 条的 where 必须含 section 与 anchor（§9.3）" % k)
        if it.get("grade") and it["grade"] not in GRADES:
            bad.append("第 %d 条 grade 非法：%s" % (k, it["grade"]))
    return bad


def main(argv=None):
    ap = argparse.ArgumentParser(prog="make_crs",
                                 description="编号治理改动项 -> CRS")
    ap.add_argument("--docx", required=True, help="被改文档（用来取指纹）")
    ap.add_argument("--items", required=True, help="改动项 json")
    ap.add_argument("-o", "--out", required=True, help="输出的 CRS json")
    ap.add_argument("--html", help="输出的可勾选 HTML")
    ap.add_argument("--source", default="numbering-xref",
                    help="source 字段（默认 numbering-xref）")
    ap.add_argument("--baseline", default="", help="baseline.note 补充说明")
    a = ap.parse_args(argv)

    if not os.path.exists(a.docx):
        print("[err] 找不到文档：%s" % a.docx)
        return 1
    items = load_items(a.items)
    bad = validate(items)
    if bad:
        print("[err] 清单不合规，未产出：")
        for b in bad:
            print("   - " + b)
        return 2
    if not items:
        print("[warn] 清单为空 —— 没有需要治理的项。仍然产出一份空单留档。")

    b = CRS.CrsBuilder(
        a.docx, sources=(a.source,),
        baseline_note=(a.baseline or "编号与引用治理单；目标指纹变化即作废"))
    for it in items:
        b.add(it.get("grade", "B"), it.get("category", "结构"), it["where"],
              it.get("before", ""), it.get("after", ""),
              it.get("basis", ""), risk=it.get("risk", "high"),
              impact=it.get("impact", "无"),
              executor=it.get("executor", "manual"),
              notes=it.get("notes", ""), source=a.source)
    c = b.build()
    CRS.write_json(c, a.out)
    print("-> %s" % a.out)
    s = CRS.summarize(c)
    print("   %d 项，分档 %s" % (s["total"], s["by_grade"]))
    if a.html:
        CRS.render_html(c, a.html)
        print("-> %s" % a.html)
    print()
    print("下一步：与 lint / polish / 事实核查 的清单一起合并：")
    print("  python ../../docx-unify-pipeline/scripts/crs_merge.py "
          "--docx %s --in %s --in <其他单>.json -o <统一>.json"
          % (os.path.basename(a.docx), os.path.basename(a.out)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
