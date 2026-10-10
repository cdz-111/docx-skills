# -*- coding: utf-8 -*-
"""
pipeline_state.py —— 七阶段流水线的**状态账本 + 门禁机**。

为什么要有它：这条链路有 7 个阶段、跨 3 个 skill、每个阶段之间都是一道「等老板批准」
的闸门。没有账本就会出现三个典型事故：

  1. 老板勾的是 S5 的清单，执行的是 S3 之前就排好的旧单 → 改错地方；
  2. 中途 Word 重存过一次，段落号全变，后续阶段还在用旧坐标 → 全篇错位；
  3. 上一轮已经改好的项，这一轮又被当成本轮待办报一遍 → 清单失去信任。

本脚本把「阶段进度 + 文件指纹 + 每阶段的产物与闸门结论」落成一份 JSON，
任何一步写盘前先 `check`，指纹不符就拒绝继续。

零第三方依赖。只用标准库。

用法
----
    PY=~/.workbuddy/binaries/python/envs/default/Scripts/python.exe

    # 1) 开局建账本（默认落在 <文档目录>/_unify/pipeline_state.json）
    "$PY" pipeline_state.py init "课题3-中期科技报告.docx"

    # 2) 看进度（哪一步该做、哪一步在等批准）
    "$PY" pipeline_state.py status

    # 3) 标记阶段进度
    "$PY" pipeline_state.py set S1 done
    "$PY" pipeline_state.py set S1 blocked --note "等老板确认样式单"

    # 4) 登记产物与闸门结论
    "$PY" pipeline_state.py artifact S1 "课题3-中期科技报告_样式确认单.json"
    "$PY" pipeline_state.py gate S1 approved --by "老板" --note "全部同意"

    # 5) 写盘前自查（指纹 + 闸门），不通过给非零退出码
    "$PY" pipeline_state.py check --stage S6

退出码：0 通过 / 3 无账本 / 4 闸门未满足 / 5 指纹不符 / 6 用法错误
"""
import argparse
import hashlib
import json
import os
import sys
import time
import zipfile

VERSION = "1.0"

# 阶段定义。gate 取值：
#   "boss"  = 本阶段结束必须等老板明确批准才能进下一阶段
#   "self"  = 本阶段自身有机器可验的断言（verify/verify_all 全绿）即可推进
#   "auto"  = 机械步骤，无需批准
STAGES = [
    ("S0", "侦察与备份", "auto"),
    ("S1", "样式基准确认", "boss"),
    ("S2", "编号与交叉引用治理", "self"),
    ("S3", "题注补插与引用落地", "boss"),
    ("S4", "排版（分页/空白/孤字/同页/表格/空格）", "boss"),
    ("S5", "内容校验与事实核查", "boss"),
    ("S6", "修订落盘与交付", "boss"),
]

STAGE_IDS = [s[0] for s in STAGES]
STAGE_TITLE = {s[0]: s[1] for s in STAGES}
STAGE_GATE_KIND = {s[0]: s[2] for s in STAGES}

STATUSES = ("pending", "running", "done", "blocked", "skipped", "failed")
GATE_STATES = ("n/a", "awaiting_approval", "approved", "rejected")

EXIT_OK = 0
EXIT_NO_STATE = 3
EXIT_GATE = 4
EXIT_FINGERPRINT = 5
EXIT_USAGE = 6


# ------------------------------------------------------------------ 基础
def _now():
    return time.strftime("%Y-%m-%dT%H:%M:%S+08:00", time.localtime())


def fingerprint(docx):
    """与 docx-format/docxfmt/crs.py 的 fingerprint 保持同构（字段名一致）。"""
    docx = os.path.abspath(docx)
    with open(docx, "rb") as f:
        raw = f.read()
    with zipfile.ZipFile(docx) as z:
        x = z.read("word/document.xml")
    return {
        "path": docx,
        "size": os.path.getsize(docx),
        "md5": hashlib.md5(raw).hexdigest(),
        "document_xml_md5": hashlib.md5(x).hexdigest(),
        "paragraphs": x.count(b"<w:p "),
        "captured_at": _now(),
    }


def default_workdir(docx):
    return os.path.join(os.path.dirname(os.path.abspath(docx)), "_unify")


def state_path(workdir):
    return os.path.join(workdir, "pipeline_state.json")


def find_state(start=None):
    """从 start（默认 cwd）向上找 pipeline_state.json，最多 3 层。"""
    cur = os.path.abspath(start or os.getcwd())
    for _ in range(3):
        p = state_path(cur)
        if os.path.exists(p):
            return p
        p2 = os.path.join(cur, "_unify", "pipeline_state.json")
        if os.path.exists(p2):
            return p2
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return None


def load(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save(state, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    return path


def _log(ledger, action, **kw):
    """ledger 是账本 dict。第一参数不能叫 state —— 调用处会传 state= 关键字。"""
    rec = {"at": _now(), "action": action}
    rec.update(kw)
    ledger.setdefault("history", []).append(rec)


# ------------------------------------------------------------------ 命令
def cmd_init(args):
    docx = os.path.abspath(args.docx)
    if not os.path.exists(docx):
        print("[err] 找不到文档：%s" % docx)
        return EXIT_USAGE
    wd = os.path.abspath(args.workdir or default_workdir(docx))
    sp = state_path(wd)
    if os.path.exists(sp) and not args.force:
        print("[err] 账本已存在：%s" % sp)
        print("      要重开请加 --force（会丢历史，但指纹会重记）")
        return EXIT_USAGE
    state = {
        "version": VERSION,
        "doc": fingerprint(docx),
        "spec": os.path.abspath(args.spec) if args.spec else "",
        "workdir": wd,
        "created_at": _now(),
        "updated_at": _now(),
        "stages": {},
        "history": [],
    }
    for sid, title, gk in STAGES:
        state["stages"][sid] = {
            "id": sid, "title": title, "gate_kind": gk,
            "status": "pending", "gate": "n/a", "gate_by": "",
            "gate_note": "", "artifacts": [], "notes": [],
            "started_at": None, "ended_at": None,
        }
    save(state, sp)
    print("账本已建立：%s" % sp)
    print("  目标文档 md5 %s（%d 字节 / %d 段）"
          % (state["doc"]["md5"], state["doc"]["size"], state["doc"]["paragraphs"]))
    print()
    print("下一步：%s %s —— %s" % (STAGE_IDS[0], STAGE_TITLE[STAGE_IDS[0]],
                                   "先备份，再 inspect"))
    return EXIT_OK


def _resolve(args):
    p = args.state or find_state()
    if not p or not os.path.exists(p):
        print("[err] 找不到流水线账本。先跑：pipeline_state.py init <文档.docx>")
        return None
    return p


def cmd_status(args):
    p = _resolve(args)
    if not p:
        return EXIT_NO_STATE
    st = load(p)
    doc = st["doc"]
    print("=" * 74)
    print("流水线账本　%s" % p)
    print("=" * 74)
    print("目标文档　%s" % doc["path"])
    print("  指纹　md5 %s ／ %d 段（%s 记录）"
          % (doc["md5"][:12], doc["paragraphs"], doc["captured_at"]))

    # 指纹是否还成立
    stale = None
    if os.path.exists(doc["path"]):
        now_fp = fingerprint(doc["path"])
        if now_fp["md5"] != doc["md5"]:
            stale = now_fp
    print()
    if stale:
        print("⚠️  文档已被改动：md5 %s → %s（%+d 段）"
              % (doc["md5"][:12], stale["md5"][:12],
                 now_fp["paragraphs"] - doc["paragraphs"]))
        print("    → 后续阶段的段落坐标全部失效，必须重新出单（协议 §11.2）")
    else:
        print("✅ 指纹一致，账本有效")
    print()
    print("%-5s %-8s %-18s %s" % ("阶段", "状态", "闸门", "标题"))
    print("-" * 74)
    nxt = None
    for sid, title, gk in STAGES:
        s = st["stages"][sid]
        flag = {"pending": "·", "running": "▶", "done": "✓",
                "blocked": "⏸", "skipped": "−", "failed": "✗"}[s["status"]]
        gate = s["gate"] if s["gate"] != "n/a" else ("待批准" if s["gate_kind"] == "boss" else "自检")
        print("%-5s %s %-6s %-18s %s" % (sid, flag, s["status"], gate, title))
        if nxt is None and s["status"] in ("pending", "running", "blocked"):
            nxt = s
    print()
    if nxt:
        print("👉 当前该做：%s %s —— %s" % (nxt["id"], nxt["title"],
              "等老板批准" if nxt["status"] == "blocked" else "见 references/pipeline-stages.md"))
        arts = [a for a in st["stages"][nxt["id"]]["artifacts"]]
        if arts:
            print("   已有产物：%s" % "、".join(arts))
    else:
        print("👉 七阶段全部走完")
    if args.json:
        print()
        print(json.dumps(st, ensure_ascii=False, indent=2))
    return EXIT_OK


def cmd_set(args):
    p = _resolve(args)
    if not p:
        return EXIT_NO_STATE
    sid = args.stage.upper()
    if sid not in STAGE_IDS:
        print("[err] 未知阶段：%s（可用：%s）" % (sid, " ".join(STAGE_IDS)))
        return EXIT_USAGE
    if args.status not in STATUSES:
        print("[err] 未知状态：%s（可用：%s）" % (args.status, " ".join(STATUSES)))
        return EXIT_USAGE
    st = load(p)
    s = st["stages"][sid]
    s["status"] = args.status
    if args.status == "running" and not s["started_at"]:
        s["started_at"] = _now()
    if args.status in ("done", "skipped"):
        s["ended_at"] = _now()
    if args.note:
        s["notes"].append({"at": _now(), "text": args.note})
    st["updated_at"] = _now()
    _log(st, "set", stage=sid, status=args.status, note=args.note or "")
    save(st, p)
    print("%s → %s%s" % (sid, args.status, ("　「%s」" % args.note) if args.note else ""))
    if args.status == "done" and s["gate_kind"] == "boss" and s["gate"] != "approved":
        print("  ⚠️  %s 是需要老板批准的阶段，但闸门还是 %s。" % (sid, s["gate"]))
        print("      这不是错误的记录，但 check 会拦；请补："
              "gate %s approved --by <谁>" % sid)
    return EXIT_OK


def cmd_artifact(args):
    p = _resolve(args)
    if not p:
        return EXIT_NO_STATE
    sid = args.stage.upper()
    if sid not in STAGE_IDS:
        print("[err] 未知阶段：%s" % sid)
        return EXIT_USAGE
    st = load(p)
    ap = os.path.abspath(args.path)
    st["stages"][sid]["artifacts"].append({
        "path": ap,
        "exists": os.path.exists(ap),
        "at": _now(),
    })
    st["updated_at"] = _now()
    _log(st, "artifact", stage=sid, path=ap)
    save(st, p)
    mark = "✅" if os.path.exists(ap) else "⚠️ 文件不存在"
    print("%s 产物登记：%s　%s" % (sid, ap, mark))
    return EXIT_OK


def cmd_gate(args):
    p = _resolve(args)
    if not p:
        return EXIT_NO_STATE
    sid = args.stage.upper()
    if sid not in STAGE_IDS:
        print("[err] 未知阶段：%s" % sid)
        return EXIT_USAGE
    state = getattr(args, "gate_state", None)
    if state not in GATE_STATES:
        print("[err] 未知闸门状态：%s（可用：%s）" % (state, " ".join(GATE_STATES)))
        return EXIT_USAGE
    st = load(p)
    s = st["stages"][sid]
    s["gate"] = state
    s["gate_by"] = args.by or ""
    s["gate_note"] = args.note or ""
    st["updated_at"] = _now()
    _log(st, "gate", stage=sid, state=state, by=args.by or "",
         note=args.note or "")
    save(st, p)
    print("%s 闸门 → %s%s" % (sid, state, ("（%s）" % args.by) if args.by else ""))
    if state == "approved" and s["status"] == "blocked":
        print("  提示：批准拿到了，把阶段标成 running 继续："
              "pipeline_state.py set %s running" % sid)
    return EXIT_OK


def cmd_next(args):
    """打印下一个该做的阶段 + 该跑的原始命令（供编排层直接执行）。"""
    p = _resolve(args)
    if not p:
        return EXIT_NO_STATE
    st = load(p)
    for sid, title, gk in STAGES:
        s = st["stages"][sid]
        if s["status"] in ("done", "skipped"):
            continue
        print("%s\t%s\t%s\t%s" % (sid, s["status"], title, s["gate"]))
        if gk == "boss" and s["gate"] != "approved":
            print("GATE\t必须拿到老板批准后：pipeline_state.py gate %s approved --by <谁>" % sid)
        return EXIT_OK
    print("DONE\t七阶段全部走完")
    return EXIT_OK


def cmd_check(args):
    """写盘前自查。逐项检查，任一不过给非零退出码。"""
    p = _resolve(args)
    if not p:
        return EXIT_NO_STATE
    st = load(p)
    doc = st["doc"]
    ok = True

    # ① 目标文档还在吗 / 指纹还成立吗
    if not os.path.exists(doc["path"]):
        print("[5] 目标文档不存在：%s" % doc["path"])
        return EXIT_FINGERPRINT
    now_fp = fingerprint(doc["path"])
    if now_fp["md5"] != doc["md5"]:
        print("[5] 指纹不符 —— 文档在出单之后被改过")
        print("     出单时 md5 %s（%d 段）" % (doc["md5"], doc["paragraphs"]))
        print("     现在   md5 %s（%d 段）" % (now_fp["md5"], now_fp["paragraphs"]))
        print("     → 本账本作废，重新 init（协议 §11.2：不要打补丁，重新出单）")
        print()
        print("❌ 自查未通过，禁止写盘。")
        return EXIT_FINGERPRINT

    # ② 要进到哪个阶段，前面所有 boss 闸门必须 approved
    target = (args.stage or "").upper() or None
    if target and target not in STAGE_IDS:
        print("[6] 未知阶段：%s" % target)
        return EXIT_USAGE
    upto = STAGE_IDS.index(target) if target else len(STAGE_IDS)
    for sid in STAGE_IDS[:upto]:
        s = st["stages"][sid]
        if s["status"] in ("done", "skipped"):
            if s["gate_kind"] == "boss" and s["gate"] != "approved":
                print("[4] %s「%s」标着 done，但闸门是 %s —— 没有老板批准就推进了"
                      % (sid, s["title"], s["gate"]))
                print("     → 停下来补批准：gate %s approved --by <谁>" % sid)
                ok = False
            continue
        print("[4] %s「%s」还没完成（%s）—— 不能跳到 %s"
              % (sid, s["title"], s["status"], target or "后续阶段"))
        ok = False
        break

    # ③ 已登记产物是否还在
    for sid in STAGE_IDS:
        for a in st["stages"][sid]["artifacts"]:
            if not os.path.exists(a["path"]):
                print("[4] %s 的产物不见了：%s" % (sid, a["path"]))
                ok = False

    if ok:
        print("✅ 自查通过：指纹一致、前置闸门齐、产物在位")
        print("   可以安全进入 %s" % (target or "下一步"))
        return EXIT_OK
    print()
    print("❌ 自查未通过，禁止写盘。")
    return EXIT_GATE


# ------------------------------------------------------------------ CLI
def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="pipeline_state",
        description="docx 七阶段流水线的状态账本与门禁机")
    ap.add_argument("--state", help="账本 json 路径（默认自动向上找）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    i = sub.add_parser("init", help="建账本")
    i.add_argument("docx")
    i.add_argument("--workdir", help="账本目录（默认 <文档目录>/_unify）")
    i.add_argument("--spec", help="生效的 spec json（样式确认单）")
    i.add_argument("--force", action="store_true")
    i.set_defaults(func=cmd_init)

    s = sub.add_parser("status", help="看进度")
    s.add_argument("--json", action="store_true", help="额外打印完整 json")
    s.set_defaults(func=cmd_status)

    t = sub.add_parser("set", help="标记阶段状态")
    t.add_argument("stage")
    t.add_argument("status", choices=list(STATUSES))
    t.add_argument("--note")
    t.set_defaults(func=cmd_set)

    a = sub.add_parser("artifact", help="登记产物")
    a.add_argument("stage")
    a.add_argument("path")
    a.set_defaults(func=cmd_artifact)

    g = sub.add_parser("gate", help="记闸门结论")
    g.add_argument("stage")
    # 注意：dest 必须避开 --state（父级全局参数），否则会被位置参数覆盖
    g.add_argument("gate_state", metavar="STATE", choices=list(GATE_STATES))
    g.add_argument("--by", help="谁批的")
    g.add_argument("--note")
    g.set_defaults(func=cmd_gate)

    n = sub.add_parser("next", help="打印下一个该做的阶段")
    n.set_defaults(func=cmd_next)

    c = sub.add_parser("check", help="写盘前自查（指纹 + 闸门 + 产物）")
    c.add_argument("--stage", help="准备进入的阶段，如 S6")
    c.set_defaults(func=cmd_check)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
