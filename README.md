# docxfmt —— docx 排版工具

一套**规范驱动**的 docx 排版工具。一份 JSON spec 描述「要排成什么样」，
工具照做。换个模板 = 换一份 spec，代码一行不改。

```
tool/
├── SKILL.md              封装成 WorkBuddy Skill 的说明
├── README.md             本文件
├── spec/
│   ├── from-template.json        样例：照模板排
│   └── chapter-numbering.json      样例：分章编号 + 公式编号
└── docxfmt/
    ├── __main__.py       python -m docxfmt 入口
    ├── cli.py            命令行
    ├── spec.py           规范定义 + 默认值 + 样例生成
    ├── docxio.py         解包/打包/OOXML 构造
    ├── styles.py         样式表（按 w:name 索引、移植、几何、新建）
    ├── recognize.py      语义识别 -> 文档树
    ├── render.py         套样式 + 抹直接格式 + 样式降级
    ├── toc.py            目录 / 图目录 / 表目录
    ├── layout.py         页面 / 分节 / 分页符 / 制表位
    ├── captions.py       题注 SEQ + 交叉引用 REF + 公式编号
    ├── pipeline.py       主流水线
    ├── verify.py         三层校验
    └── winfield.py       COM 重算域（独立进程）
```

---

## 五个子命令

| 命令 | 作用 |
|---|---|
| `inspect` | **只侦察不改文件**。段落/表格/图片数、封面区、目录区、节点统计、样式使用分布 |
| `spec` | 生成 spec 样例 json，照着改 |
| `run` | 排版。核心命令 |
| `fields` | 重算全部域（目录/图表目录/题注编号/交叉引用） |
| `verify` | 校验：域完整性、编号连续性、断链、内容零丢失、真机样式对比 |

---

## 能力清单（老板要的那些全在）

| 能力 | 实现 | 备注 |
|---|---|---|
| **套样式（自动）** | `render.py` | 按语义角色自动分配：正文/标题 1-6/题注/**图片**/表格文字/目录标题/目录 1-3/附录/公式编号 |
| **图片样式** | `spec.styles.new_styles` | 独立样式，默认**无首行缩进 + 居中**（不是正文样式） |
| **题注样式** | 同上 | 独立样式，模板没有就新建 |
| **目录标题样式** | 同上 | 独立样式，默认**居中 + 无首行缩进**，与正文区分 |
| **插入目录** | `TOC \o "1-3" \h \z \u` | 层级可配，插在第一个一级标题前 |
| **插入图目录** | `TOC \h \z \c "图"` | 按 SEQ 标识符收集，与题注自动对齐 |
| **插入表目录** | `TOC \h \z \c "表"` | 同上 |
| **分页符** | `pPr/pageBreakBefore` | 按标题级别配，不用插空段落 |
| **分节符** | `pPr/sectPr` | 支持指定级别前分节、指定标题单独成横向节 |
| **页面设置** | 边距/方向/首页不同/奇偶页 | 横向节可保留（宽表改纵向会溢出） |
| **页眉页脚** | `header`/`footer` 样式 | 随模板移植 |
| **题注自动编号** | `SEQ 图` / `SEQ 表` | global（图1）/ chapter（图2-1）可切换 |
| **交叉引用** | `REF _figN \h` | 按作者原编号精确定位，40 处实测零错位 |
| **公式自动编号** | `SEQ 式` | 兼容「编号同段」与「编号独立成段」两种形态 |
| **分章编号** | `SEQ 章 \h` + `\s 1` | 章号锚点只插在正文章节，引言/附录不占号 |

---

## 用法

```bash
cd tool

# 0. 先侦察（只读）
python -m docxfmt inspect 输入.docx

# 1. 生成 spec 样例
python -m docxfmt spec --init my.json --mode template

# 2. 排版
python -m docxfmt run 输入.docx -o 输出.docx --spec my.json

# 3. ⚠️ 必须重算域
python -m docxfmt fields 输出.docx -o 输出_域重算.docx

# 4. 校验
python -m docxfmt verify 输出_域重算.docx --template 模板.docx
```

### 快捷方式（不想写 spec 时）

```bash
--template PATH         照这个模板的样式排
--caption-scheme chapter   分章编号
--number-sep -          分章分隔符
--pagebreak-levels 1,2  几级标题前分页
--toc / --figure-list / --table-list
--no-captions / --no-xref / --no-toc
--dry-run               只统计不落盘
--tree-json / --tree-txt   导出文档树
```

---

## spec 速查

```jsonc
{
  "source":   { "style_from": "template|requirement|keep",
                "template": "参考模板.docx" },
  "styles": {
    "map":       { "paragraph": "正文", "heading": {"1":"标题 1"},
                   "caption": "题注", "image": "图片",
                   "table_text": "表格表头文字", "toc_title": "目录标题",
                   "toc": {"1":"目录 1"}, "formula_number": "公式编号" },
    "new_styles": { "图片": {"align":"center","first_line":0},
                    "目录标题": {"align":"center","first_line":0} },
    "geometry":  { "题注": {"size":15,"font":"仿宋"} }
  },
  "numbering": { "caption": "global|chapter|none",
                 "formula": "none|global|chapter",
                 "sep": "-", "skip_keywords": ["引言","附录"] },
  "toc":         { "enabled": true, "levels": "1-3", "title": "目录",
                   "after": "first_heading", "replace_static": false },
  "figure_list": { "enabled": true, "title": "插图清单" },
  "table_list":  { "enabled": true, "title": "附表清单" },
  "captions":    { "enabled": true, "kinds": ["图","表"] },
  "xref":        { "enabled": true },
  "pagebreak":   { "before_heading": [1], "keep_with_next": true },
  "section":     { "insert": [1], "landscape_nodes": ["附表"] },
  "page":        { "margin": {"top":1440,"bottom":1440,
                              "left":1701,"right":1701} }
}
```

`styles.map` 写的是 **`w:name` 语义名，不是 styleId**。工具会自动做别名解析，
写「标题 1」在英文模板里也能找到 `heading 1`。

---

## 实测结果

在两份真实文档上跑通（**均为脱敏统计，不含原文内容**）：

### 文档 A：39MB / 8 节（含 4 横向）/ 33 表 / 49 图 / 2564 段

```
[styles] 移植模板样式：覆盖 13，新增 12，basedOn 重映射 7
[parse]  封面区 [0,21)；9 个一级标题 / 15 二 / 13 三 / 35 四
[render] 套样式 1552 段
[caption] 题注 59 条（图 38 / 表 21），认领孤儿书签 1
[xref] 交叉引用 40 处；定位方式：原编号命中 33 / 占位符就近 4 / 方位就近 3
[toc] 插图目录 + 表目录 + 正文目录（真 TOC 域，WPS 识别为 1 目录 + 2 图表目录）
[page] 8 个节全部应用，横向节保留方向
[tabs] 缩放 5 个段落 6 个制表位（按节版心，不是全局 max）
```
重算域后：图目录 38 条、表目录 21 条、正文目录三级全带页码；
SEQ 图 1..38 连续、SEQ 表 1..21 连续、REF 134 处**零断链**；
媒体/表格/图片/分节**零丢失**。

### 文档 B：63MB / 4 节 / 52 表 / 45 图 / 4385 段（启用分章编号 + 公式编号）

```
[styles] 新建样式 ['目录标题','图片','公式编号']
[parse]  封面区 [0,154)；8 个一级标题 / 15 二 / 28 三 / 41 四
[render] 套样式 1021 段（3273 段表格文字自动降级到正文 —— 该文档无表格文字样式）
[caption] 分章编号：插入隐藏 SEQ 章 锚点 4 个（引言/总结/附录不占号）
[caption] 题注 97 条（图 45 / 表 52）
[formula] 公式编号转 SEQ 域 75 处（原文是独立成段的 (1)(2)(3)…）
[toc] 删除旧目录块 2 处 → 插图目录 + 表目录 + 正文目录
```
完整性：媒体/表格/图片/分节**零丢失**，REF 248 处**零断链**。

### 走完整判断回路（AI 亲自定层级）后的增益

| | 纯规则 | AI 判断后 |
|---|---|---|
| 识别出的标题数 | 67 | **131** |
| 层级分布 | L1=9 L2=10 L3=13 L4=35 **L5=0** | L1=8 L2=18 L3=27 L4=58 **L5=20** |
| 正文目录 | 只有章节 | **完整三级 + 页码** |
| 交叉引用 | 40 处零错位 | 40 处零错位 |

L5 从 0 涨到 20 —— 那些是「某某平台 / 某某系统」这类技术名，
规则完全认不出来，是判断环节补上的。

---

## 五条铁律（都是踩坑换来的）

1. **直接操作 OOXML**，不走 python-docx / HTML 中转。
   HTML 中转会丢域、书签、批注（实测 96 个 SEQ、101 个 REF 无法保真）。

2. **一切按 `w:name` 定位，绝不按 styleId，也绝不复制模板的 styles.xml。**
   不同文档的 styleId 是两套编号；WPS 保存还会把语义 styleId 重写成数字。
   直接复制会让「styleId=5 的段落从题注变成标题 4」。

3. **树是逻辑视图，不重建 OOXML。** 节点持有原 lxml 元素引用，只改不改写。
   一旦按树重新生成 XML，域/书签/批注/图片关系全丢。

4. **改完必须真机打开验证。** OOXML 写对了 ≠ 渲染对。

5. **`run` 之后必须 `fields`。** 目录、图表目录、题注编号、交叉引用全是域，
   不重算就还是旧值。

---

## 本轮挖出并修掉的坑

| 坑 | 症状 | 修法 |
|---|---|---|
| 封面区按「第一个一级标题」算，但只认「一、」 | 纯数字编号的论文（`1.` `1.1`）一个都找不到 → 退化成 `len(paras)` → **整篇 4136 段被当封面跳过** | 优先认「已套 heading 1 样式 / outlineLvl=0」，其次「一、」，再次「1.」+ 加粗佐证 |
| 空标题段当封面终点 | p153 是**空的** heading 1，封面区被截断 | 封面终点要求段落非空 |
| `1.` 型弱信号误判 | 表格里「0.7084」这种数字被当一级标题 | 加 `_looks_like_heading()`：必须通体加粗（≥80%）或用黑体 |
| 公式编号两种形态 | 一份是「公式+编号同段」，一份是「整段就是 (1)」 | 两种都支持：同段的**只删末尾编号 run**（整段重写会删掉 oMath，踩过 417→342）；独立段的整段重写 |
| 表格文字样式不存在 | 3273 段全不套样式 | 自动降级到正文样式，日志报 `_退回正文=N` |
| 旧目录没删 | 新旧目录并存，正文里指向 `_Toc*` 的 PAGEREF 断链 96 处 | 删真 TOC 域目录；**静态文字目录保留并警告**（域刷不了它，删了内容就没了） |
| lxml 元素真值判断 | `elem or other` 在新版 Python 会变 False | 全改 `is not None` |
| 输出文件被占用 | `PermissionError` 让人摸不着头脑 | 捕获并提示「文档正被 Word/WPS 打开」 |
| 章节归属写成 O(n²) | 2500 段直接卡死 | 一次遍历同时算「最近标题」和「最近一级标题」 |
| 校验脚本自己写错口径 | 报「43 处引用全不一致」 | SEQ 域结果是纯数字 `1`，REF 显示是「图 1」→ 只能比数字 |

---

## 已知限制

- **静态文字目录刷不了**：源文档目录若是 `WPSOffice手动目录` 样式（纯文字、无 TOC 域），
  程序无法更新页码。工具会识别出来并警告，需人工在 WPS 更新，
  或设 `toc.replace_static=true` 让工具删掉重建。
- **WPS 劫持 Word COM**：本机 `Word.Application` 被 WPS 接管（`Name` 伪装成
  `Microsoft Word`，只有 `Path` 暴露），只能走 WPS。渲染与真 Word 有差异。
- **63MB 级文档重算域慢**：WPS 要几分钟，命令行会超时，建议后台跑。
- **尚未实现**：插入批注（comments.xml）、页眉文案自动生成、图表编号用
  `TOC \t "题注,1"` 旧式收集。
