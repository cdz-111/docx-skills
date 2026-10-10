# -*- coding: utf-8 -*-
"""
spec.py —— 排版规范：工具的输入契约。

一份 spec 描述「要排成什么样」，工具只管照做。
换模板 = 换一份 spec，不用改代码。

字段速查（完整示例见 spec/full.json）：

    source.style_from        样式来源："template"（照模板排）
                                    "requirement"（照文字规范排）
                                    "keep"（只做内容层，不动样式）
    source.template          模板 docx 路径（style_from=template 时必填）
    source.requirement       文字规范 json（style_from=requirement 时必填）

    styles.map               角色 -> 样式 w:name 语义名
                            paragraph / heading{1..6} / caption / image /
                            table_text / toc_title / toc{1..3} /
                            appendix / formula_number / header / footer
    styles.geometry          要强制覆盖模板的几何（不填就以模板为准）
                            { 样式名: {size, font, first_line, align, ...} }

    numbering.caption        题注编号：global(图1) / chapter(图2-1)
    numbering.formula        公式编号：none / global / chapter
    numbering.sep            分章分隔符，默认 "-"
    numbering.skip_keywords  这些词开头的章节不参与章号（引言/附录…）

    toc.enabled              插正文目录
    toc.levels               "1-3"
    toc.title                目录标题文字；null = 不加标题
    toc.title_style          目录标题用的样式（默认 toc_title）
    toc.after                目录放在哪："first_heading"（第一个一级标题前）
                            "explicit"（用 toc.anchor 指定的段落之后）
    toc.anchor               toc.after=explicit 时，锚段落以什么文字开头

    figure_list.enabled      插图目录（TOC \\c "图"）
    table_list.enabled       插表目录（TOC \\c "表"）
    figure_list.title        默认 "插图清单" / table_list.title "附表清单"

    pagebreak.before_heading 几级标题前插分页符，如 [1] 表示一级标题另起一页
    pagebreak.keep_with_next 标题段设 keepNext

    section.insert            要插分节符的位置
                              "h1"    每个一级标题前都插一节
                              "h2"    每个二级标题前插
                              []      不插
    section.orientation      "keep"（保留原方向）/ "all-portrait"
    section.landscape_nodes  哪些节点要放横向节（正则匹配标题文字）
    section.first_page_diff  首页不同（页眉页脚）
    section.odd_even         奇偶页不同

    tables.border_style      "three_line"(数据表三线) / "full_grid"(长文本表全框)
                             / "auto"(按内容自动判定，默认)
    spacing.normalize_cjk_latin  中文↔数字/英文/公式 之间的空格要不要归一（去空格）
    spacing.keep_caption_space   题注与域缓存保持「图 1」形态，不参与去空格
    lint.scope               正文校验的范围："body" = 只扫正文
    lint.exclude             正文校验要排除的部分：formula(公式段) / appendix(附录)
"""
import json
import os

# 默认模板（用户没给模板时用它）。相对 skill 根目录定位。
DEFAULT_REQUIREMENT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "spec", "default.json")

DEFAULTS = {
    "source": {
        "style_from": "requirement",
        "template": None,
        "requirement": None,          # None = 用 spec/default.json 这份默认模板
    },
    "styles": {
        "map": {
            "paragraph": "正文",
            "heading": {"1": "标题 1", "2": "标题 2", "3": "标题 3",
                        "4": "标题 4", "5": "标题 5", "6": "标题 6"},
            "caption": "题注",
            "image": "图片",
            "table_text": "表格表头文字",
            "toc_title": "目录标题",
            "toc": {"1": "目录 1", "2": "目录 2", "3": "目录 3"},
            "appendix": "附录",
            "formula_number": "公式编号",
            "header": "页眉",
            "footer": "页脚",
        },
        "geometry": {},
        "new_styles": {},
    },    "numbering": {
        "caption": "global",
        "formula": "none",
        "sep": "-",
        "skip_keywords": ["引言", "前言", "绪论", "序言", "摘要", "ABSTRACT",
                          "目录", "插图清单", "附表清单", "图目录", "表目录",
                          "参考文献", "致谢", "附录"],
    },
    "toc": {
        "enabled": True,
        "levels": "1-3",
        "title": "目录",
        "title_style": "目录标题",
        "after": "first_heading",
        "anchor": None,
    },
    "figure_list": {
        "enabled": False,
        "title": "插图清单",
        "title_style": "目录标题",
    },
    "table_list": {
        "enabled": False,
        "title": "附表清单",
        "title_style": "目录标题",
    },
    "captions": {
        "enabled": True,
        "kinds": ["图", "表"],
        # S3 续表：一张逻辑表被拆成多个 w:tbl 时，续段（紧贴前一张表、前面没段落的表）
        # 不该有独立题注。skip_continuation=True 时维持现状（只跳过、不出题注）。
        # continuation_label 默认 None=关闭；设为模板（如 "续表 {n}"）后，
        # 给每个续段额外出一条 CRS 项（executor=insert_continuation_label），
        # {n} = 被续的那张逻辑表的编号（沿相邻表链回溯到第一张表的题注号）。
        # ⚠️ 标注段**不加 SEQ 域、不进图表清单**，只是套「题注」样式的普通段落。
        # ⚠️ 模板渲染结果应**以「续」起头**（如 "续表 {n}"）或**带「（续）」后缀**
        #    （如 "表 {n}（续）"），才能被 `verify_all` 第 10 项（正文残留硬写引用）
        #    识别为标注并豁免；若自定义成别的形状，需同步调整 verify_all 的
        #    豁免规则（`RE_CONT_A` / `RE_CONT_B`），否则工具会把自己插入的标注误判成硬写引用。
        #
        # ⚠️ 默认值真源就在**这里的 DEFAULTS**：这两个开关**刻意不放进 `spec/default.json`**。
        #    因为 default.json 会以 `requirement` 身份被 `_load_requirement` 反向合并，
        #    **覆盖**用户 `--spec` 里的显式值——放进去就等于把用户的手按住，
        #    导致 `--spec {"captions":{"skip_continuation":false}}` 关不掉、失效。
        "skip_continuation": True,
        "continuation_label": None,
    },
    "xref": {
        "enabled": True,
    },
    "pagebreak": {
        "before_heading": [1],
        "keep_with_next": True,
    },
    "section": {
        "insert": [],
        "orientation": "keep",
        "landscape_nodes": [],
        "first_page_diff": False,
        "odd_even": False,
    },
    "page": {
        "copy_from_template": True,
    },
    # ---- 表格美化（polish 读；见 references/table-style.md）
    "tables": {
        "border_style": "auto",          # auto | three_line | full_grid
        "three_line_data_table": True,   # 数据表 -> 三线表
        "full_grid_text_table": True,    # 长文本表 -> 全框
        "text_table_min_chars": 60,      # 单格文本超过这么多字 -> 判为长文本表
        "max_width_percent": 100,
        "repeat_header_row": True,
    },
    # ---- 空格规则（polish 读；见 references/space-rules.md）
    "spacing": {
        "normalize_cjk_latin": True,     # 去掉 中文↔数字/英文/公式 之间的空格
        "keep_caption_space": True,      # 题注/域缓存保持「图 1」不参与
    },
    # ---- 正文校验范围（lint 读）
    "lint": {
        "scope": "body",                 # body = 只扫正文
        "exclude": ["formula", "appendix"],
        "punctuation_fullwidth": True,
        "no_space_after_punct": True,
    },
}


def _merge(base, over):
    """深合并：over 覆盖 base，dict 递归。"""
    out = dict(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = v
    return out


class Spec(dict):
    """规范对象。用法：Spec.load(path) / Spec(dict) / Spec() 取默认值。"""

    @classmethod
    def load(cls, path):
        if not path:
            return cls(DEFAULTS)
        with open(path, "r", encoding="utf-8") as f:
            return cls(_merge(DEFAULTS, json.load(f)))

    def __init__(self, d=None):
        super().__init__(_merge(DEFAULTS, d or {}))

    # ---- 便捷读取
    def style_of(self, role, level=None):
        """角色 -> 样式 w:name 语义名。"""
        m = self["styles"]["map"]
        v = m.get(role)
        if isinstance(v, dict):
            return v.get(str(level) if level is not None else None)
        return v

    def geom(self, style_name):
        return (self["styles"].get("geometry") or {}).get(style_name)

    @property
    def sep(self):
        return self["numbering"].get("sep", "-")

    def enabled(self, section):
        return bool(self.get(section, {}).get("enabled"))


def write_sample(path, mode="template"):
    """写一份带注释的 spec 样例，让用户照着改。"""
    if mode == "default":
        # 把内置默认模板抄一份出去，供用户手改后再喂回来
        with open(DEFAULT_REQUIREMENT, "r", encoding="utf-8") as f:
            src = f.read()
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(src)
        return path
    if mode == "template":
        s = {
            "source": {"style_from": "template",
                       "template": r"C:\path\参考模板.docx"},
            "numbering": {"caption": "global", "sep": "-"},
            "toc": {"enabled": True, "levels": "1-3", "title": "目录"},
            "figure_list": {"enabled": True, "title": "插图清单"},
            "table_list": {"enabled": True, "title": "附表清单"},
            "pagebreak": {"before_heading": [1]},
        }
    else:
        s = {
            "source": {"style_from": "requirement",
                       "requirement": r"C:\path\格式要求.json"},
            "styles": {
                "map": {
                    "paragraph": "Normal",
                    "heading": {"1": "heading 1", "2": "heading 2"},
                    "caption": "caption",
                    "image": "图片",
                    "table_text": "表格表头文字",
                    "toc_title": "TocTitle",
                    "toc": {"1": "toc 1", "2": "toc 2", "3": "toc 3"},
                },
                "geometry": {
                    "题注": {"size": 15, "font": "仿宋", "align": "center"},
                    "图片": {"align": "center", "first_line": 0},
                },
            },
            "numbering": {"caption": "chapter", "formula": "chapter", "sep": "-"},
            "toc": {"enabled": True, "levels": "1-3", "title": "目录"},
            "figure_list": {"enabled": True},
            "table_list": {"enabled": True},
        }
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    return path
