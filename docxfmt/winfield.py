# -*- coding: utf-8 -*-
"""
winfield.py —— 用 Word/WPS COM 重算全部域（独立进程）。

为什么要独立进程：win32com 只装了 anaconda3 那个 python，
而主流程跑在 managed python 里。拆成子进程，两边各用各的解释器。

⚠️ 本机 Word COM 被 WPS 完全劫持（连 CLSID 都被接管，Name 伪装成
   "Microsoft Word"，只有 Path 暴露），只能走 WPS。
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                              errors="replace")
import win32com.client as win32

WORD_CLSID = "{000209FF-0000-0000-C000-000000000046}"


def open_word():
    fallback = None
    for i, how in enumerate((WORD_CLSID, "Word.Application")):
        try:
            app = win32.DispatchEx(how)
        except Exception as e:
            print("  [warn] %s 启动失败 %s" % (how, repr(e)[:100]))
            continue
        path = str(getattr(app, "Path", "") or "")
        is_wps = "wps" in path.lower() or "kingsoft" in path.lower()
        if not is_wps and "Microsoft Word" in str(app.Name):
            return app, False
        print("  [warn] %s 实际是 WPS（Path=%s）" % (how, path))
        if i == 0:
            fallback = app
            continue
        try:
            app.Quit()
        except Exception:
            pass
    if fallback is None:
        raise RuntimeError("起不了任何 Word/WPS 实例")
    print("  [warn] 本机 Word COM 已被 WPS 劫持，用 WPS 执行。")
    return fallback, True


def main():
    src = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else None
    app, is_wps = open_word()
    app.Visible = False
    app.DisplayAlerts = 0
    print("已启动:", app.Name, app.Version)
    try:
        doc = app.Documents.Open(src, ReadOnly=False, AddToRecentFiles=False)
        print("打开:", os.path.basename(src), "| 域总数", doc.Fields.Count)
        n = 0
        for i in range(1, 12):                 # 正文/页眉页脚/脚注尾注
            try:
                doc.StoryRanges(i).Fields.Update()
                n += 1
            except Exception:
                pass
        doc.Fields.Update()
        # 多遍：TOC 的页码依赖分页，分页又依赖 SEQ 结果
        for _ in range(2):
            try:
                for t in doc.TablesOfContents:
                    t.Update()
            except Exception:
                pass
            try:
                for t in doc.TablesOfFigures:
                    t.Update()
            except Exception:
                pass
            doc.Repaginate()
        nt = 0
        try:
            for t in doc.TablesOfContents:
                t.Update()
                nt += 1
        except Exception:
            pass
        nf = 0
        try:
            for t in doc.TablesOfFigures:
                t.Update()
                nf += 1
        except Exception:
            pass
        print("更新完成：%d 个 story 的域；目录 %d 个；图表目录 %d 个"
              % (n, nt, nf))
        if out:
            try:
                doc.SaveAs2(out, FileFormat=12)
            except Exception as e:
                print("  SaveAs2 失败（%s），改用 Save" % repr(e)[:60])
                doc.Save()
            print("另存 ->", out, os.path.getsize(out), "bytes")
        else:
            doc.Save()
        doc.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
