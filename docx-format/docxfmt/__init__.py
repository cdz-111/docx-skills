# -*- coding: utf-8 -*-
"""
docxfmt —— docx 排版工具箱。

解决的问题
    批量把 docx 排成规范版式：套样式、插目录/图目录/表目录、插分页符与分节、
    生成题注、建立交叉引用。这些活手工做又慢又容易漏，脚本又散落各处。

设计原则
    1. **直接操作 OOXML**，不走 python-docx / HTML 中转。
       HTML 中转会丢域、书签、批注，无法保真。
    2. **一切按 `w:name` 语义定位**，绝不按 styleId。
       不同文档的 styleId 是两套编号；WPS 保存还会把语义 styleId 重写成数字。
    3. **树是逻辑视图，不重建 OOXML**。节点持有原 lxml 元素引用，只改不改写。
    4. **写对了 ≠ 渲染对**，改完必须真机打开验证。
"""
__version__ = "1.0.0"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(t, ns=W):
    return "{%s}%s" % (ns, t)


__all__ = ["qn", "W", "M", "R", "XSPACE", "__version__"]
