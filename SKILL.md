---
name: docx-format
description: >
  批量把 Word 文档排成规范版式。直接操作 OOXML，不走 python-docx/HTML 中转，
  所以域、书签、批注、图片关系零损坏。能力：照模板或照文字规范移植样式、
  套正文/标题/题注/图片/表格文字样式、插目录、图目录、表目录（TOC 域）、
  分页符、分节符、页面设置、页眉页脚、自动生成题注编号（SEQ）、
  建立交叉引用（REF）、公式自动编号、真机重算域并校验。
  核心是 propose → 智能体定层级 → run --overrides 的回路：
  工具做确定性判断并挑出疑点，层级语义由调用方亲自拍板。
  触发词：排版 docx、Word 文档格式、按模板排版、统一文档格式、套样式、
  插目录、加图目录、生成题注、交叉引用、批量修改 Word、docx 格式化、
  文档层级不对、标题没样式、目录没页码。
---

# docx 排版工具

## 这是什么

**规范驱动**的 docx 排版工具。一份 JSON spec 描述「要排成什么样」，工具照做。
换模板 = 换一份 spec，**代码一行不改**。

**关键前提：层级判断由调用方（你/另一个智能体）亲自做，工具只做机械执行。**
工具负责确定性判断 + 挑出疑点；语义判断（这个段落到底是几级标题、该不该进目录）
规则永远补不完，必须有人（或 AI）读上下文拍板。详见下面的「核心回路」。

## 什么时候用

- 「把这批 docx 按这个模板的格式统一一下」
- 「给这篇论文排版：正文宋体小四、标题黑体、图要有题注、目录自动生成」
- 「文档里的图号表号是手打的，改一处要手工改十处」→ 题注+交叉引用自动化
- 「插图目录 / 表目录 / 让目录能自动更新」

---

## ⭐ 核心回路：propose → 判断 → run

**这是本 skill 与普通脚本最重要的区别。** 不要试图用规则解决所有层级问题。

```bash
cd <skill 目录>

# ① 先侦察（只读）
python -m docxfmt inspect 输入.docx

# ② 挑出「拿不准的」，带上下文
python -m docxfmt propose 输入.docx -o 待拍板.json --report 待拍板.md

# ③ ⭐ 这里停下来，读待拍板.json 逐条判断层级，写进 overrides
#    格式：{"p1642": {"type":"heading","level":2}, "p1644": {"type":"paragraph"}}
#    简写：{"p1642": "heading2"}

# ④ 把判断喂回去渲染
python -m docxfmt run 输入.docx -o 输出.docx --spec my.json --overrides 我的判断.json

# ⑤ ⚠️ 必须重算域
python -m docxfmt fields 输出.docx -o 输出_域重算.docx

# ⑥ 校验
python -m docxfmt verify 输出_域重算.docx --template 模板.docx
```

**⑤ 不能省。** 目录、图表目录、题注编号、交叉引用全是域，不重算还是旧值。

---

## ③ 怎么判断（这一步是核心，别偷懒）

`propose` 挑出三类，每类的判断依据不同：

| 类别 | 含义 | 怎么判 |
|---|---|---|
| `样式` / `outlineLvl` | 已有样式依据 | **别改**。有明确依据的判断是可靠的 |
| `仅文本规则` | 只靠编号正则判定 | 看编号形态对不对：`(1)`→L4、`1.2.1`→L3、`2.`→L2。形态不对才改 |
| `疑似` | 判成了正文，但短、无句末标点 | **重点看这里**。读 `before` / `after` 上下文判断 |

**判断的四个维度**（按重要性排序）：

1. **看它在同类里是什么** —— 找到上下文里已经判对的兄弟段，看它是几级你就该是几级。
   这是最可靠的依据。
2. **看它的编号形态** —— `（一）`和`（1）`通常差一级；`1.` `1.1` `1.1.1` 依次递进。
3. **看它后面跟什么** —— 后面是长正文 → 它是标题；后面是小节号 → 它是父级标题（= L1）。
4. **看它在不在表格里** —— 表格内的短文本（`0.7561`、`项目目标`、`通过`）**几乎都不是标题**。

**几个反复踩到的判断**：

- **表格里的填写项不是标题**：`所属专项：重点专项`、`日期：年月日`、`项目编号` ——
  长得像短标题，其实是表格里等人填的空格。判成正文。
- **注释不是标题**：`注：总结凝练1~3项典型成果案例` —— 同上。
- **附表类不要当 L1**：附表挂在「八、其他材料」之下，判 L2 才不会把目录撑爆，
  也不会让附表占掉正文章节号（正文编号会被打乱）。
- **表格数据永远不是标题**：`0.7084`、`87.59`、`(42)` —— 再短也不是。

**判断要克制**：拿不准的保持工具的判断（`样式`/`outlineLvl` 依据的），
只改「看起来明显不对」的。乱改比不改更糟。

## 快速开始（不含判断的纯机械模式）

只想套样式、不想动层级时用：

```bash
# 照模板排 + 全文连续编号（图 1、图 2）
python -m docxfmt run in.docx -o out.docx --template tpl.docx \
    --caption-scheme global

# 分章编号（图 2-1、图 2-2）
python -m docxfmt run in.docx -o out.docx --template tpl.docx \
    --caption-scheme chapter --number-sep -

# 插齐三种目录
python -m docxfmt run in.docx -o out.docx --template tpl.docx \
    --toc --figure-list --table-list

# 只要样式，不要动内容层
python -m docxfmt run in.docx -o out.docx --template tpl.docx \
    --no-toc --no-figure-list --no-table-list --no-captions --no-xref
```

**注意**：不传 `--overrides` 时，工具只用规则判断。**质量明显低于走完整回路。**

## spec 是什么

一份 JSON，描述「要排成什么样」。这是工具与脚本的区别：
**换模板 = 换一份 spec，代码一行不改。**

```jsonc
{
  "source": {
    "style_from": "template",          // template=照模板排 | requirement=照文字规范排
    "template": "路径/参考模板.docx"
  },
  "styles": {
    "map": {                            // 角色 -> 样式 w:name（不是 styleId！）
      "paragraph": "正文",
      "heading": {"1": "标题 1", "2": "标题 2", "3": "标题 3"},
      "caption": "题注",
      "image": "图片",                  // 图片独立样式：无首行缩进 + 居中
      "table_text": "表格表头文字",
      "toc_title": "目录标题",          // 目录标题独立样式：居中 + 无缩进
      "toc": {"1": "目录 1", "2": "目录 2", "3": "目录 3"}
    },
    "new_styles": {                     // 模板里没有的，新建
      "目录标题": {"align":"center","first_line":0,"based_on":"正文",
                   "size":16,"font":"黑体","bold":true},
      "图片":     {"align":"center","first_line":0,"based_on":"正文"}
    },
    "geometry": {                       // 强制覆盖模板几何
      "题注": {"size":15,"font":"仿宋","align":"center"}
    }
  },
  "numbering": {
    "caption": "global",                // global(图1) | chapter(图2-1) | none
    "formula": "none",                  // none | global | chapter
    "sep": "-",
    "skip_keywords": ["引言","摘要","目录","参考文献","致谢","附录"]
  },
  "toc":        {"enabled":true,"levels":"1-3","title":"目录",
                 "after":"first_heading"},
  "figure_list":{"enabled":true,"title":"插图清单"},
  "table_list": {"enabled":true,"title":"附表清单"},
  "captions":   {"enabled":true,"kinds":["图","表"]},
  "xref":       {"enabled":true},
  "pagebreak":  {"before_heading":[1],"keep_with_next":true},
  "section":    {"insert":[1],              // 几级标题前插分节符
                 "landscape_nodes":[]},     // 匹配这些正则的标题单独成横向节
  "page":       {"margin":null}             // null=用模板边距
}
```

`styles.map` 里写的是 **`w:name` 语义名**，不是 styleId。工具会自动做别名解析，
所以写「标题 1」在英文模板里也能找到 `heading 1`。

## 铁律（都是踩坑换来的，不要绕过）

1. **直接操作 OOXML**，不走 python-docx / HTML 中转。
   HTML 中转会丢域、书签、批注（实测 96 个 SEQ、101 个 REF 无法保真）。

2. **一切按 `w:name` 定位，绝不按 styleId，也绝不复制模板的 styles.xml。**
   不同文档的 styleId 是两套编号；WPS 保存还会把语义 styleId 重写成数字。
   直接复制会导致「styleId=5 的段落从题注变成标题 4」这种语义错位。

3. **树是逻辑视图，不重建 OOXML。** 节点持有原 lxml 元素引用，只改不改写。
   一旦按树重新生成 XML，域/书签/批注/图片关系全丢。

4. **改完必须真机打开验证。** OOXML 写对了 ≠ 渲染对。
   直接格式残留、继承链断了，只有打开才知道。

5. **改编号后必须重算域。** 见上面第 3 步。

## 常见问题

| 症状 | 原因 / 修法 |
|---|---|
| 目录不更新、还是老页码 | 没跑 `fields`。目录若是 `WPSOffice手动目录` 静态文字（无 TOC 域），程序刷不了，需人工在 WPS 更新 |
| 图目录是空的 | 题注的 SEQ 标识符必须与 `\c` 一致。工具用 `SEQ 图`/`SEQ 表`，对应 `\c "图"`/`\c "表"` |
| 标题样式没套上 | 看日志里 `[render] ⚠ 方案里有但文档没有的样式`。把该样式名加进 `styles.map`，或用 `new_styles` 新建 |
| 图片位置跑出页面 | 换边距/装订线后版心变了，制表位不会自动跟着变。工具已按节缩放；自定义边距请在 `page.margin` 里给全 |
| 只有最后一节页面设置生效 | `body/sectPr` 只是最后一节，前面的节在 `w:p/w:pPr/w:sectPr` 里。用工具的 `layout.all_sectprs()` |
| 内容少了 | 跑 `verify` 看 `[完整性]` 那几行。媒体/表格/图片数必须与源一致 |
| 交叉引用显示「错误！未定义书签」 | 原文档把书签挂在空段落上。工具会自动认领（`adopt_orphan_bookmarks`）；仍有残留时看日志 |
| 公式编号乱/公式消失 | 工具只删末尾编号 run，不整段重写。**永远不要整段重写公式段**，会把 oMath 删掉 |
| 横向宽表被压扁 | 保持 `section.orientation = "keep"`，横向节别动 |

## 目录结构

```
docxfmt/
├── __main__.py     python -m docxfmt 入口
├── cli.py          命令行
├── spec.py         规范定义（工具的输入契约）
├── docxio.py       解包/打包/OOXML 构造
├── styles.py       样式表（按 w:name 索引、移植、几何）
├── recognize.py    语义识别 -> 文档树
├── propose.py      ⭐ 挑出疑点 + 上下文；应用人工/AI 的 overrides
├── render.py       套样式 + 抹直接格式
├── toc.py          目录 / 图目录 / 表目录
├── layout.py       页面 / 分节 / 分页符 / 制表位
├── captions.py     题注 SEQ + 交叉引用 REF + 公式编号
├── pipeline.py     主流水线（顺序有讲究，别改）
├── verify.py       三层校验
└── winfield.py     COM 重算域（独立进程，因为 win32com 只装在 anaconda3）

spec/                spec 样例
├── from-template.json       照模板排
└── chapter-numbering.json     分章编号 + 公式编号
```

## 调用契约

- `inspect` 永远先跑，返回结构统计供判断
- `run` 不修改输入文件，输出到 `-o`
- `run` 之后**必须** `fields`，除非用户明确不要域
- 交付前跑 `verify`，把 `[完整性]` 那几行报给用户
- 报告要给出**具体数字**（套了多少条样式、题注多少条、引用多少处、零丢失几项），
  不要只说「已完成」
