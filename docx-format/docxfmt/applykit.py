# -*- coding: utf-8 -*-
"""
applykit.py —— **结构类改动**的通用执行框架。

和 `revise.py` 的分工：
    * `revise.py`  管 **文字级**改动：`w:ins` / `w:del` + 逐条批注，走修订模式。
      它的不变量是「计数一个都不许变」。
    * `applykit.py`（本模块）管 **结构级**改动：插分页符、改表格边框、插题注段。
      这类改动**天然会改变结构计数**（多一个 `w:p`、多一组 `w:tblBorders`），
      所以不能塞进 revise 的校验里；而且 Word 的修订模式**本来就不记录**
      分页符 / 边框 / 图片尺寸的变化 —— 硬套修订模式只会让文档显示错乱。

统一的门禁仍然一样（见 references/change-request-sheet.md）：
    出单 → 老板勾选 → 执行 → 三栏回执。
本模块只负责「执行 + 回执」这一段，**不负责出单**。
"""
import hashlib
import os
import shutil
import time

from lxml import etree

from . import qn
from .crs import approved_items, load
from .docxio import Docx, pack, unpack


def stamp():
    return time.strftime("%Y%m%d_%H%M%S", time.localtime())


def md5(path):
    if not path or not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def backup(docx):
    """带秒级时间戳的备份。绝不覆盖原文件。"""
    p = os.path.join(
        os.path.dirname(os.path.abspath(docx)),
        "%s_备份_%s.docx" % (os.path.splitext(os.path.basename(docx))[0],
                             stamp()))
    shutil.copy2(docx, p)
    return p


def load_todo(crs_path, approved_path=None):
    """读 CRS，返回 (全部条目, 待执行条目, 未勾选条目)。"""
    crs = load(crs_path)
    if approved_path:
        todo = approved_items(crs, approved_path)
    else:
        todo = [i for i in crs["items"]
                if i["decision"] in ("approved", "modified")]
    pending = [i for i in crs["items"]
               if i["decision"] not in ("approved", "modified", "closed")]
    return crs, todo, pending


class Applier:
    """结构类改动的执行器基类。

    子类实现 `apply_one(doc, paras, item) -> int`（1 = 改了，0 = 跳过）。
    本类负责：备份 → 解包 → 逐条执行 → 计数核对 → 落盘 → 回执。
    """

    #: 允许变化的计数键 -> 允许的增量（None = 任意）
    allowed_delta = {}
    #: 本执行器认领的 executor 集合。None = 全认。
    #: 一份批准清单会被多个执行器读（revise 管文字、applykit 子类管结构），
    #: 不属于自己的条目应**原样退回**，而不是记成"失败"——
    #: 否则回执里会堆一大片假失败，掩盖真正的失败。
    supported = None

    def __init__(self, docx, author="docx-format 排版"):
        self.src = docx
        self.author = author

    # ---- 子类实现
    def apply_one(self, doc, paras, item):
        raise NotImplementedError

    # ---- 计数
    def counts(self, root):
        out = {}
        for tag in ("p", "tbl", "drawing", "bookmarkStart", "bookmarkEnd",
                    "fldChar", "instrText", "br"):
            out[tag] = len(list(root.iter(qn(tag))))
        return out

    def check_counts(self, before, after):
        problems = []
        for k, b in before.items():
            a = after.get(k, 0)
            if a == b:
                continue
            lim = self.allowed_delta.get(k)
            if lim is None:
                problems.append("%s 计数从 %d 变成 %d（本类改动不允许变）"
                                % (k, b, a))
            elif lim != "*" and (a - b) != lim:
                problems.append("%s 增量应为 %s，实际 %+d" % (k, lim, a - b))
        return problems

    # ---- 主流程
    def run(self, crs_path, approved_path=None, out=None, dry=False,
            log=print):
        crs, todo, pending = load_todo(crs_path, approved_path)
        mine, others = [], []
        for it in todo:
            if self.supported is not None and \
                    it.get("executor") not in self.supported:
                others.append(it)
            else:
                mine.append(it)
        if others:
            log("[scope] %d 项不属于本执行器（executor∈%s），原样退回"
                % (len(others), sorted({i.get("executor") for i in others})))
        workdir = os.path.join(os.path.dirname(os.path.abspath(self.src)),
                               "_apply_" + type(self).__name__.lower())
        unpack(self.src, workdir)
        doc = Docx(workdir)
        done, fail = [], []

        paras = doc.paras()
        for it in mine:
            try:
                n = self.apply_one(doc, paras, it)
                (done if n else fail).append(
                    it if n else (it, "定位不到目标（文档可能已变，需重新出单）"))
            except _Skip as e:
                fail.append((it, str(e)))
            except Exception as e:                  # noqa: BLE001
                fail.append((it, "%s: %s" % (type(e).__name__, e)))

        receipt = {
            "target": crs.get("target", {}),
            "summary": {"total": len(crs["items"]), "approved": len(todo),
                        "changed": len(done), "pending": len(pending),
                        "failed": len(fail), "other": len(others)},
            "changed": done, "pending": pending, "others": others,
            "failed": [{"id": i["id"], "reason": r} for i, r in fail],
            "backup": None, "backup_md5": None, "dry": dry,
            "author": self.author,
        }
        if dry:
            log("[dry] 未落盘。将改 %d 项，失败 %d 项，退回 %d 项"
                % (len(done), len(fail), len(others)))
            shutil.rmtree(workdir, ignore_errors=True)
            return receipt

        before = self._before_counts
        after = self.counts(doc.root)
        problems = self.check_counts(before, after)
        if problems:
            log("[验收] 未通过，**不落盘**：")
            for x in problems:
                log("   - " + x)
            shutil.rmtree(workdir, ignore_errors=True)
            receipt["failed"].append({"id": "-", "reason": "验收未通过：" +
                                      "；".join(problems)})
            return receipt

        bk = backup(self.src)
        receipt["backup"] = bk
        receipt["backup_md5"] = md5(bk)
        doc.save()
        out = out or os.path.join(
            os.path.dirname(os.path.abspath(self.src)),
            "%s_排版.docx" % os.path.splitext(os.path.basename(self.src))[0])
        pack(workdir, out)
        shutil.rmtree(workdir, ignore_errors=True)
        receipt["output"] = out
        log("[ok] %d 项已改，%d 项失败 -> %s" % (len(done), len(fail), out))
        return receipt

    def run_prepared(self, crs_path, approved_path=None, out=None, dry=False,
                     log=print):
        """同 run()，但先解包取 before 计数。供需要「先看现状」的子类调用。"""
        workdir = os.path.join(os.path.dirname(os.path.abspath(self.src)),
                               "_apply_" + type(self).__name__.lower())
        unpack(self.src, workdir)
        self._before_counts = self.counts(
            etree.parse(os.path.join(workdir, "word", "document.xml")).getroot())
        shutil.rmtree(workdir, ignore_errors=True)
        return self.run(crs_path, approved_path, out, dry, log)


class _Skip(Exception):
    pass


# ------------------------------------------------------------------ 回执
def write_receipt(receipt, path, title="排版改动回执", extra=None):
    s = receipt["summary"]
    L = ["# %s —— %s" % (title, os.path.basename(
        receipt.get("target", {}).get("path", "") or "文档")),
        "",
        "- 依据：改动事项单（target md5 `%s`）"
        % receipt.get("target", {}).get("md5", "?"),
        "- 备份：%s　md5 `%s`" % (receipt.get("backup"),
                                  receipt.get("backup_md5")),
        "- 执行者：%s" % receipt.get("author", ""),
        "- 结果：已改 %d / 未勾选未改 %d / 失败 %d / 转交他处 %d"
        % (s["changed"], s["pending"], s["failed"], s.get("other", 0)),
        ""]
    if extra:
        for line in extra:
            L.append("- " + line)
        L.append("")
    L += ["## 一、已改（%d）" % s["changed"], "",
          "| CR-ID | 位置 | 动作 |", "|---|---|---|"]
    for it in receipt["changed"]:
        w = it.get("where") or {}
        pos = "%s · 段 %s" % (w.get("section", ""), w.get("para", ""))
        if w.get("wps_para"):
            pos += "（WPS 段 %s / 第 %s 页）" % (w.get("wps_para"),
                                              w.get("page", "?"))
        L.append("| %s | %s | %s |" % (it["id"], pos, it.get("after", "")))
    L += ["", "## 二、未勾选未改（%d）" % s["pending"], "",
          "| CR-ID | 档 | 位置 | 摘要 |", "|---|---|---|---|"]
    for it in receipt["pending"]:
        w = it.get("where") or {}
        L.append("| %s | %s | %s · 段 %s | %s |"
                 % (it["id"], it.get("grade", ""), w.get("section", ""),
                    w.get("para", ""), _md(it.get("before") or it.get("after"))))
    if receipt.get("others"):
        L += ["", "## 三、转交他处（%d）—— 不在本执行器职责内，未改"
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
          "- 备份路径与 md5（上）—— 回滚就靠这一行",
          "- 本类改动（分页符 / 表格边框 / 插入段）**不被 Word 修订模式记录**，"
          "所以不进修订通道",
          "- 域未重算：需在 Word 里 `Ctrl+A` → `F9` 一次"]
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path


def _md(s, n=46):
    import re
    s = re.sub(r"\s+", " ", (s or "")).strip()
    return s if len(s) <= n else s[:n] + "…"
