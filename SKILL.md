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
  本 skill 是 **docx 处理链路的权威实现（单一真源）**：
  一份 JSON spec 描述「要排成什么样」，换模板=换 spec、代码零改动。
  四个常用入口：`styles`（样式基准确认单，只出单不套用）、
  `run`（排版）、`lint`（正文校验：缺字、漏字、语句不通、术语与体例不一致、
  数据自洽，只出清单不改文档）、`revise`（按批准项**以修订模式**落盘并逐条批注）。
  ⚠️ 对文档做任何写盘修改之前，必须先产出「改动事项单（CRS）」可勾选清单，
  经调用方逐条批准后才允许执行——这是本 skill 的硬门禁，不可绕过。
  触发词：排版 docx、Word 文档格式、按模板排版、统一文档格式、套样式、
  插目录、加图目录、生成题注、交叉引用、批量修改 Word、docx 格式化、
  文档层级不对、标题没样式、目录没页码、
  校对文档、检查哪里不通顺、哪里缺字、错别字、术语不一致、
  给我改动清单、改动事项单、改前确认、审批后再改、可勾选清单、
  样式基准、样式确认单、修订模式、留痕修改、批注修改原因。
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

## 四个入口（先认清你该敲哪个）

| 命令 | 阶段 | 干什么 | 会不会改文档 |
|---|---|---|---|
| `styles` | S1 样式基准 | 逐部分列出目标样式，出**样式确认单**（可手改） | ❌ 只出单 |
| `run` | S2/S3 排版 | 按 spec 套样式、补题注、建交叉引用、插目录 | ✅ 受门禁 |
| `lint` | S5 正文校验 | 错字/缺字/不通顺/术语/体例/数据自洽，出**改动事项单** | ❌ 只出单 |
| `revise` | S6 修订落盘 | 只执行 **已批准** 的项，**修订模式**写入并逐条批注 | ✅ 受门禁 |

`run` 是"按规范排"，`revise` 是"按批准改"。两者是**不同的动作**：
排版改的是结构与样式，修订改的是文字内容。也可以在链路里被
`docx-unify-pipeline` 编排着依次调用。

---

## 单一真源（本 skill 在生态里的位置）

**`docx-format` 是 docx 处理链路的权威实现。** 具体约定：

1. **`spec`（JSON）是"要排成什么样"的唯一真源。** 换模板、换规范、换编号口径
   都通过换/改一份 JSON 完成，**不新增代码分支、不新增命令行开关去硬编码规则**。
   新增一类排版规则 = 在 spec 里加一节 + 在 `docxfmt/` 里加一个消费者。
2. **默认模板是 `spec/default.json`**（源 `source.style_from = "requirement"`）。
   没给任何格式要求时 `run`/`styles`/`lint` 都回退到它，**不要**在代码里另写一套默认值。
   给了参考模板（`--template tpl.docx`）则以模板为准。
3. **样式一律按 `w:name` 语义名定位**，绝不按 `styleId`、绝不复制模板 `styles.xml`。
   spec 里写中文名（「正文」「标题 1」），工具用 `NAME_ALIASES` 做别名解析
   （「标题 1」↔`heading 1`、「正文」↔`Normal`），所以在英文模板里同样命中。
4. **CRS 的读写只有一份实现**：`docxfmt/crs.py`。`lint` / `polish` / `propose` /
   外部编排层都从这里拿能力，协议细则在 `references/change-request-sheet.md`。
   **不许各写一套 CRS 读写**——否则 `.html` 与 `.json` 必然漂移。
5. **历史资产 `D:/Code/快速统一docx格式/`（YAML 模板驱动的 CLI）不再并行演进**，
   只作为**规格来源**：它的 `templates/*.yaml` 里的字号/行距/缩进/编号口径，
   按需移植进 `spec/*.json`。两套实现同时演进会造成同一份文档两种结果。
   需要那份仓库里的分析型工具（`compare_docx_format.py`、`extract_format_spec.py`、
   `render_docx_report.py`）时，直接调用，不复制代码。

---

## ⭐ 核心回路：propose → 出 CRS → 🛑 批准 → run

**这是本 skill 与普通脚本最重要的区别。** 不要试图用规则解决所有层级问题，
更不要在没拿到批准的情况下往文档里写。

```bash
cd <skill 目录>

# ① 先侦察（只读）
python -m docxfmt inspect 输入.docx

# ② 挑出「拿不准的」+ 全部待改动项，带上下文
python -m docxfmt propose 输入.docx -o 待拍板.json --report 待拍板.md

# ③ 把待改动项归集成「改动事项单 CRS」（可勾选 .html + 机器读 .json）
#    规则见下一节「改动审批门禁」。工具做确定性判断并分档，
#    语义与取舍由调用方拍板。

# 🛑🛑 闸门：把 CRS 交给调用方，停下来等他勾选 / 回话
#    没拿到批准之前，③ 之后一步都不许走。

# ④ 只把「已批准」的项写回（未勾选项原样不动）
python -m docxfmt run 输入.docx -o 输出.docx --spec my.json --overrides 我的判断.json

# ⑤ ⚠️ 必须重算域
python -m docxfmt fields 输出.docx -o 输出_域重算.docx

# ⑥ 校验
python -m docxfmt verify 输出_域重算.docx --template 模板.docx
```

**③ 与 ④ 之间那一步不能省。** 它是调用方的否决权：
「先给我看清单，我勾了我认的，你再动」。详见下面的「改动审批门禁」。

**⑤ 不能省。** 目录、图表目录、题注编号、交叉引用全是域，不重算还是旧值。

---

## 🛑 改动审批门禁（CRS）—— 硬规矩，不可绕过

**任何写盘修改（改样式、调层级、改文字、改编号、改域）之前，
必须先产出「改动事项单（CRS）」并拿到调用方的逐条批准。**

这条规矩的适用对象是**所有**会改文档的入口：`run`、`fields`，以及
`docx-formula-image-to-omml` / `docx-numbering-xref-pipeline` 那边的修复脚本。
**只读命令**（`inspect`、`propose`、`lint`、各类 audit）不受限，随时可跑。

### CRS 是什么

一次改动 = 一份 CRS。CRS 是「我打算改什么」的**逐条可勾选清单**，
既给人看（勾），也给机器读（执行）。

| 产物 | 用途 |
|---|---|
| `<文档名>_改动事项.html` | **可勾选页**：逐条复选框、按档筛选、实时统计、一键导出 `approved.json`。给人用 |
| `<文档名>_改动事项.json` | **唯一裁决源**：执行侧只读它。`decision` 字段是批准与否的唯一依据 |

协议细则（JSON schema、HTML DOM 契约、导出格式、批量授权话术解析表、
风险分级执行策略、回执表模板）见 `references/change-request-sheet.md`；
可勾选页模板见 `templates/crs.html`。

### 表格长什么样

每条一行，**勾选列在最左**：

| 勾选 | ID | 档 | 位置 | 现状 | 建议改为 | 依据 | 风险 | 影响面 |
|---|---|---|---|---|---|---|---|---|
| [ ] | CR-001 | **S 必改** | 3.4.1 · 段 214 | 「如图 45 所示」 | 「如图 44 所示」 | 该图实际编号为图 44 | 低 | 无 |
| [ ] | CR-004 | **B 建议** | 全文 · 68 处 | 「图 N所示」 | 「图 N 所示」 | 体例统一 | 低 | 无 |

- **ID**：`CR-00N`，全局唯一、跨轮次稳定。下一轮只说「CR-007 还没改」就能对上。
- **档**：`S 必改`（错/缺字/数据矛盾）｜`A 应改`（不通顺/体例不一致）｜
  `B 建议`（风格统一）｜`C 待核`（口径需调用方拍板）
- **位置**：必须能定位——章节号 + 段号 +（图号/表号/书签名/域 ID）。
  「第 3 章某处」不算位置。
- **风险**：`低`＝只动文字｜`中`＝动样式/层级｜`高`＝动编号、SEQ/REF 域、书签、交叉引用
- **影响面**：会连带改到哪些（写"无"就明确写"无"，不要留空）

### 四条硬规矩

1. **没出 CRS → 拒绝写盘。** 不许"我先改了再说""改完再给你看清单"。
   调用方要的是**事前**否决权，不是事后通知。
2. **CRS 出了但零勾选 / 没有明确回话 → 拒绝执行**，并显式回话"清单已就绪，在等你勾选"。
   不要把沉默当默认同意。
3. **只执行 `decision=approved` 的项。** 未勾选项必须原样不动，
   且必须出现在交付回执的「未勾选未改」栏里——**不得静默丢弃**，
   否则下一轮又要重新发现一遍。
4. **批准 ≠ 跳过安全措施。** 拿到批准后照旧：备份（带时间戳）→ `--dry` 干跑 →
   `verify` 全绿 → 才写盘。批准的是"改什么"，不是"跳过验证"。

### 风险分级 → 执行策略

| 风险 | 典型改动 | 执行方式 |
|---|---|---|
| 低 | 纯文字替换，不碰域/书签 | 可整批执行，事后给逐项 diff |
| 中 | 动样式、层级、标点体例 | 逐条 `--dry`，给「改前 / 改后」对照 |
| 高 | 动编号、SEQ/REF 域、书签、交叉引用 | **单条 dry + `verify_all.py` 全绿**才允许写入 |

### 只有 CRS 才能改动的三种来源

| 来源 | 入口 | 现状 |
|---|---|---|
| 排版/层级疑点 | `propose` → 归集成 CRS | 已有回路，接入 CRS 即可 |
| 编号 / 交叉引用病灶 | `docx-numbering-xref-pipeline` 的 S2 七类病灶表 | 天然就是 CRS 的行 |
| 正文文字问题（缺字、不通顺、术语、数据自洽） | `lint` | 📋 协议已定（见下），脚本待实现 |

`lint` 的定位：**只出 CRS，绝不改文档。** 发现的每条问题都落成 CRS 的一行，
由调用方勾选决定改不改。**不要**在 `lint` 里顺手把"明显的错字"直接改掉——
那正是这条门禁要防的事。

### 交付回执（拿到批准、执行完之后）

必须报三栏，缺一不可：

| 栏 | 内容 |
|---|---|
| 已改 | N 项，逐条列 `CR-ID + 改前 → 改后` |
| 未勾选未改 | M 项，逐条列 `CR-ID`（让调用方知道哪些还挂着） |
| 失败 | K 项，逐条列 `CR-ID + 失败原因` |
| 安全 | 备份路径 + md5 |

---

## 怎么判断（CRS 里的语义部分，别偷懒）

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

- **表格里的填写项不是标题**：`所属专项：xxx`、`日期：年月日`、`项目编号` ——
  长得像短标题，其实是表格里等人填的空格。判成正文。
- **注释不是标题**：`注：请补充说明` —— 同上。
- **附表类不要当 L1**：附表一般挂在正文最后一个章节之下，判 L2 才不会把目录撑爆，
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
而且**纯机械模式也受门禁约束**——规则判断出来的改动同样是改动，
照样要出 CRS、等批准。门禁拦的是"写盘"，不是"语义判断"。

---

## S1 样式确认单：`styles`

**先出单、后套用。** 老板没给格式要求就用默认模板；给了模板就照模板列。
出的是**逐部分的目标样式表**（部分 / 现有样式 / 目标样式 / 解析为 / 缺失），
这份表在 HTML 页上**可直接手改**（编号方式、目录层级、是否分页、表格边框、空格开关）。

```bash
# 用默认模板
python -m docxfmt styles 输入.docx
# 用老板给的参考模板
python -m docxfmt styles 输入.docx --template 标杆.docx
# 已有 spec 则优先用 spec
python -m docxfmt styles 输入.docx --spec my.json
```

产物（落在被改文档同目录）：

| 文件 | 用途 |
|---|---|
| `<文档名>_样式确认单.html` | 人看的可编辑确认页，改完导出 `_样式确认单.json` |
| `<文档名>_样式确认单.json` | 解析后的 spec，**确认后原样当 `--spec` 喂回 `run`** |

**顺序是硬的**：`styles` → 🛑 等老板确认/手改 → `run --spec <确认单>.json`。
**没有确认就 `run` 等于替老板定了样式。**

`styles` 会打印两行关键信息，要转述给老板：

- `目标样式走别名命中的`：spec 写中文名、文档里是英文名，工具自动对上的（如 `标题 1→heading 1`）。命中=安全。
- `目标样式在文档里找不到的部分`：这部分会用 `new_styles` 新建。要提醒老板"这些样式文档里原本没有，
  我会新建；如果界面上看着不对，在确认页里改"。

---

## S5 正文校验：`lint`

**只出 CRS，绝不改文档。** 这是本 skill 最容易被误用的一步——
**不要**在 `lint` 里顺手把"明显的错字"直接改掉，那正是门禁要防的事。

```bash
python -m docxfmt lint 输入.docx --spec 样式确认单.json
```

产物：`<文档名>_改动事项.html`（可勾选页）+ `<文档名>_改动事项.json`（唯一裁决源）。

检查项与执行器的对应关系（**执行器决定这条能不能自动改**）：

| 检查 | 类别 | 默认档 | executor | 说明 |
|---|---|---|---|---|
| P1 半角标点混用 | 体例 | B | `text_replace` | 中文句子里混了 `,` `.` `;` `:` |
| P2 标点后空格 | 体例 | B | `text_replace` | `，` 后面跟空格 |
| P3 中文↔拉丁/数字空格 | 空格 | B | `text_replace` | `精度 0.8554` → `精度0.8554` |
| P4 公式与中文空格 | 空格 | B | `manual` | 跨越 oMath 的空格，位置敏感 |
| P5 叠字 / 连续标点 | 标点 | A/C | `text_replace`(同一标点连续) / `manual`(混合标点) | 见下 |
| P6 缺字模式 | 缺字 | S | **`manual`** | 补什么字只有作者知道 |
| P7 相邻重复 run | 多录 | S | `text_replace` | 拆 run 导致的假重复要排除 |
| P8 术语变体 | 术语 | A | **`manual`** | 全文级改动 |
| P9 引用落空 | 引用 | S | `manual` | 指向不存在的图/表 |
| P10 百分点自洽 | 数据 | A | **`manual`** | 见下 |

**两条必须守住的判据**（都是实战误报换来的）：

- **混合标点（`：，`「：。」）不自动改**：`RE_MIX_PUNCT` 命中的一律 **C 档 + `manual`**。
  把 `：，` 改成 `：` 是自作主张，实际可能是作者要写两个并列结构而漏了字。
  只有**连续同一个标点**（`，，`）才是确定的冗余，走 `text_replace`。
- **数据自洽要容忍两种基准**：`由0.8554提升至0.8884，提高3.30%` 里 3.30% 是**相对增幅**，
  是对的。判据用 `min(abs(d-c), abs(d*100-c))` 做单位换算，且**跳过含「较/相比/相对/比」的表述**
  （那是另一个基准）。这类一律 `manual`。

---

## S6 修订落盘：`revise`

**只执行 `decision=approved` / `modified` 的项**，以 **Word 修订模式**（`w:ins` / `w:del`）
写入，并给每一处挂**批注说明改动原因**。未勾选项原样不动。

```bash
# 干跑：只算不改盘，先看会改什么
python -m docxfmt revise 输入.docx --crs 输入_改动事项.json --dry
# 老板在 HTML 页勾完、导出 _approved.json 之后
python -m docxfmt revise 输入.docx --crs 输入_改动事项.json \
    --approved 输入_approved.json -o 输出_修订.docx
```

产物：`输出_修订.docx` + `<输入名>_改动回执.md`（三栏回执，见门禁章节）。

**revise 的不变量**（`_validate()` 强校验，任一不满足即不落盘）：

| 量 | 要求 | 为什么 |
|---|---|---|
| `w:p` / `w:tbl` / `w:drawing` 计数 | **必须不变** | 只改文字，不许增删结构 |
| `bookmarkStart` / `bookmarkEnd` | **必须不变** | 书签是引用锚点，动了就断链 |
| `fldChar` 三态（begin/separate/end） | **必须配平且计数不变** | 域是编号与目录的命根子 |
| `document.xml` XML 可解析 | 必须 | 不产出坏文件 |
| `settings.xml` `trackChanges` | 置为 true | 否则改完没留痕 |

**跨 run 的匹配会抛 `_Skip`**（同一条 before 文本被拆在多个 `<w:t>` 里时，
不硬拼、不强改，记为该条失败），这样宁可不改也不改坏。`executor=manual` 的项
**一律跳过**——它们本来就是要人来做。

---

## S3 题注补插：`captions`

**先出单，后补插。** 找出「有图/表但没有题注」的位置，按老板选的**编号形态**给建议文本，
出 CRS 等批准；`--apply` 才真正插段。

```bash
# 只出单（不改文档）
python -m docxfmt captions 输入.docx --out 题注事项.json --html 题注事项.html
# 老板勾完、导出 _approved.json 之后补插
python -m docxfmt captions 输入.docx --apply \
    --crs 统一改动事项.json --approved 输入_approved.json \
    --out-docx 输出_补题注.docx --receipt 题注回执.md
```

**编号形态跟着 spec 的 `numbering.caption` 走**（老板在 S1 确认单里选的）：

| `numbering.caption` | 建议题注长这样 | 说明 |
|---|---|---|
| `global` | `图 1`、`表 12` | 全文连续，编号形态 `1` |
| `chapter` | `图 2-1`、`表 3-4` | 分章，编号形态 `章节号-序号`，分隔符取 `number_sep` |

**判据**（认的是"附近有没有题注"，不是"图号连不连续"）：

- **图**：非表格内的段落含 `w:drawing` → 往下 **1~2 个非空段**内找 `图\s*\d` 打头的题注。
  找到就判"已有"，找不到才建议插。
- **表**：`w:tbl` 的**上下紧邻段**（各看 1~2 段）找 `表\s*\d` 打头的题注。

⛔ **题注插在图/表下方**（图）或**上方**（表），插的是**新段落**并套 `题注` 样式。
所以它是**唯一会改变 `w:p` 计数**的执行器——`revise` 不认领它，
`captions --apply` 认领（见下"执行器分工"）。

---

## S4 排版体检：`polish`

老板的**六条排版硬要求**的检测与执行。**必须先量页**——③④⑤ 三条全靠渲染级的页码/坐标，
静态读 `document.xml` 判不出来。

```bash
# 只出单（不给 --pages 会自动调 paginate.py 量页，64MB 文档约 1 分钟）
python -m docxfmt polish 输入.docx --pages 量页.json \
    --out 排版事项.json --html 排版事项.html
# 执行结构类改动（分页符 / 表格边框）
python -m docxfmt polish 输入_修订.docx --apply \
    --crs 统一改动事项.json --approved 输入_approved.json \
    --out-docx 输出_排版.docx --receipt 排版回执.md
```

| # | 硬要求 | 类别 | 默认档 | executor | 判据要点 |
|---|---|---|---|---|---|
| ① | **>100 页时一级标题分页** | 版式 | A | `pagebreak` | **仅当总页数 > 100**；给各级标题设 `pageBreakBefore` |
| ② | 数字/英文/公式与中文**之间不留空格** | 空格 | B | `text_replace` | 与 `lint` 的 P3/P4 同源，本命令只做**判据复核** || ③ | 正文页**不出现 4 行以上大范围空白** | 版式 | S | `manual` | 见下（排除章末/表格页/图表后留白/空页/末页） |
| ④ | **不出现单字加标点独占一行** | 版式 | S | `manual` | 段末是标点 + 末字换行了 + `Δx ≤ 1.6×字号` |
| ⑤ | **题注与对应图表同页** | 版式 | A | `manual` | 题注在下方看图形段；**排除目录条目** |
| ⑥ | **表格美化**（数据表三线表 / 长文本表全框） | 表格 | B | `table_border` | 平均格字数 ≥ 阈值 或 单格 ≥ 120 字 → 全框，否则三线 |

**四条实测换来的判据（不要放宽）**：

- **② 的例外落在「模式」上，不落在「段落类型」上**：`图|表|式|公式|附图|附表 + 空格 + 数字`
  一律保留空格，**不管它出现在题注段还是正文**。原因：正文 `如图 1所示` 里的 `图 1`
  同时是 **REF 域的缓存文本**；只按"题注段"放行会让缓存变成 `图1`、与题注 `图 1` 不一致，
  一按 F9 被改回，`verify_all` 第 6 项（REF 缓存 == 语义重算值）直接 FAIL。
  另外 `revise` 有一道兜底护栏：**域结果区（`separate`~`end` 之间）的 run 一律不动**。

- **① 的段号必须与检测同一棵树**：另起 `etree.parse()` 会造成第二棵树，
  `paras.index(p)` 用 `==` 是身份比较、永远失配 → 段号全变 0 → 执行时全报"越界"。
  正确做法是用同一次解析的树 + `pos = {id(p): i for i, p in enumerate(paras)}`。
- **③ 的行高取众数**：行距实测值取**众数**（本机实测 29.0pt），
  不用中位数（29.1，被少数异类拉偏）、也不能取低分位（退化成 12.1，判据过严、误报翻倍）。
  同时**空段落 `ye` 不可信**（图形段的 `Range.Text` 是 `/`，`_norm()` 会清成空串 → 末坐标偏小 → 假报空白），
  要先做 `_has_content()` 过滤。
- **④ 判据**：`段末是标点` **且** `末字行号 > 首字行号`（确实换行了）**且** `xe - 左边距 ≤ 1.6×字号`。
  本机实测 28 个候选段 `Δx` 最小 60pt（≈5 字）→ 该文档确实没有孤字行，**报 0 处是正确结论**，别放宽阈值硬凑。
- **⑤ 必须排除目录条目**：`表 22 学习者画像域增量更新方法…\t49` 是目录行不是题注。
  凡含制表符 `\t` 或命中目录行正则的一律跳过。

**⑥ 的边框真相**：Word/WPS 的表格边框**多数写在单元格级 `w:tcPr/w:tcBorders`**，
而**单元格级优先级高于表级 `w:tblBorders`**。所以：

- 判定要**先看 `tblBorders`，null 时再汇总各格 `tcBorders`**（本机 52 张表里多张连 `tblBorders` 都没有）。
- 执行时**必须先清掉所有 `tcBorders`**，否则改表级等于没改。
- `w:tblStyle` 是 WPS 重写过的**数字 styleId**（如 41/29/30），**不可反查语义名**，别指望从它读"三线表"。

几何量（三线 / 全框）：

| 样式 | top / bottom | insideH | insideV / left / right |
|---|---|---|---|
| **三线表** | 单线 1.5pt（val=12） | 0.75pt（val=6） | `nil` |
| **全框** | 单线 0.5pt（val=4） | 0.5pt（val=4） | 0.5pt（val=4） |

---

## S5b 事实核查：`facts`

把文档里的**可核查断言**抽出来（奖项 / 官方通报 / 新闻发布），
经联网核查后把判定回填，落成 CRS。**本轮实测结论：该报告 0 条**——
说明抽取器"宁缺毋滥"，不把"落地路径""实施任务卸载""获得性能奖励"这类词当事实。

```bash
# ① 抽取（只出表，不落 CRS）
python -m docxfmt facts 输入.docx --out 待核查事实.json --md 事实核查表.md
# ② 人/Agent 联网核查后写 verdicts.json，再回填
python -m docxfmt facts 输入.docx --verdicts verdicts.json \
    --to-crs 事实核查.json --html 事实核查.html
```

**两道门禁（缺一个就误报成灾）**：

- **只走文档树 `parse()`，只处理 `type == "paragraph"`**。
  ⛔ 不要"扫到附录就停"——目录里就有 `附录A：发表论文列表158` 这一行，
  一旦按文本命中就把后面 4000 多段整片跳过（实测只剩 6 段）。附录起点要取
  **真正的 `sec_kind == "appendix"`** 的最小 idx。
- **award / news 必须有锚点**：句子里得出现**年份**或**机构/载体**
  （`RE_YEAR` / `RE_ENTITY`），否则丢弃。关键词要收紧——
  `发布|实施|落地|入选|认定为` 这类宽词会把技术描述当新闻。

判定档位：`partially_supported`→S/mid、`contradicted`→S/mid、`insufficient`→C/low。

---

## 执行器分工（为什么改文字的和改结构的要分开跑）

同一个批准清单里混着不同性质的动作，**一条命令改不完，也不该改完**：

| executor | 认领方 | 是否改计数 | 留痕方式 |
|---|---|---|---|
| `text_replace` / `run_dedup` | `revise` | **不改**（`w:p`/`w:tbl`/`w:drawing` 恒定） | Word 修订模式 `w:ins`/`w:del` + 批注 |
| `pagebreak`（①） | `polish --apply` | 不改 | 分页符/`pageBreakBefore`（Word 修订**不记录**分页） |
| `table_border`（⑥） | `polish --apply` | 不改 | 边框属性（Word 修订**不记录**边框） |
| `insert_caption` | `captions --apply` | **会增 `w:p`** | 直接插段（放最后一步） |
| `manual` | 谁都不认领 | — | 交老板人工 |

**执行顺序是硬的**：`revise` → `polish --apply` → `captions --apply`。
理由：前两个不改计数（可互相兼容），`captions` 会增段，**必须放最后**，
否则段号漂移会让前面算好的定位全部失效。

**不认领 ≠ 失败**：每个执行器有 `supported` 声明，遇到不归自己的 executor 就
**原样退回**，记入回执的"**转交他处**"栏，而不是记成失败。回执因此是**五栏**：
已改 / 未勾选未改 / 转交他处 / 失败 / 安全。

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
  "page":       {"margin":null},            // null=用模板边距

  // ↓ 下面三节给 polish / lint 消费（S4 排版与 S5 校验的口径来源）
  "tables": {                               // 表格美化口径
    "border_style": "auto",                 // auto=按内容判 | none | full
    "three_line_data_table": true,          // 数据表 → 三线表
    "full_grid_text_table": true,           // 文本占比大的表 → 全框
    "text_table_min_chars": 60              // 「文本占比大」的判定阈值
  },
  "spacing": {
    "normalize_cjk_latin": true,            // 中文↔拉丁/数字 去空格
    "keep_caption_space": true              // 题注保持「图 1」，不参与去空格
  },
  "lint": {
    "scope": "body",                        // 校验范围：body | all
    "exclude": ["formula", "appendix"]      // 正文校验不碰公式与附录
  }
}
```

`styles.map` 里写的是 **`w:name` 语义名**，不是 styleId。工具会自动做别名解析，
所以写「标题 1」在英文模板里也能找到 `heading 1`。

**默认模板**：`spec/default.json`。不传 `--spec` / `--template` / `--requirement` 时，
`run`、`styles`、`lint` 都回退到它（`source.style_from` 默认 `"requirement"`）。
默认值定义在 `docxfmt/spec.py` 的 `DEFAULT_REQUIREMENT`——**改默认口径改这里，不要在别处另写一份**。

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
| 我还没勾选，文档就被改了 | 违反门禁。停下来、回滚到备份、重新走「出 CRS → 等批准」 |
| CRS 里列了已经改好的条目 | 文档在出清单之前又被动过。**出 CRS 前必须重新核对现状**（重新提取 + 与清单基准指纹比对），旧轮次结论一律作废 |
| 清单条目定位不到 | 位置栏写得不够细。CRS 的位置必须是「章节号 + 段号 + 图号/表号/书签名」，能被独立复核 |
| 样式确认单里 15 个部分全报「找不到」 | 文档实际样式名是 `Normal`/`heading 1`/`caption`，spec 写的是中文名。已用 `NAME_ALIASES` 做别名解析；若仍 MISS，说明该样式文档里确实没有，会走 `new_styles` 新建 |
| `revise` 报某条 `_Skip` | 该条的 before 文本被拆在多个 `<w:t>` 里（常见于含公式/上下标的句子）。**不是 bug，是宁可不改也不改坏**。该条进回执「失败」栏，需人工处理 |
| `revise` 报「无 CRS / 指纹不符」 | `exit 3`=没出单就执行；`exit 5`=文档在出单后被重存过，段落号全变。**重新出单**，不要套用旧单 |
| `revise` 之后目录还是旧页码 | 修订产物同样是旧缓存，**必须再跑一次 `fields`** |
| 题注的空格被 `lint` 报成问题 | 题注「图 1」的空格是**体例要求**，由 `spacing.keep_caption_space` 保护。若报了，检查该段是否被误判成正文 |
| `lint` 报「提高3.30%」数据不符 | 相对增幅 vs 百分点两种基准。判据已做单位换算并跳过含「较/相比/相对/比」的表述；仍报则 `manual` 交人工 |

## 与兄弟 skill 的关系

| Skill | 层 | 职责 |
|---|---|---|
| **本 skill** | **权威实现层** | spec 真源 + `run`/`styles`/`lint`/`revise` + CRS 唯一实现 |
| `docx-numbering-xref-pipeline` | 编排与门禁层（编号域） | 5 阶段编号/引用治理，每步断言 + `verify_all.py` 13 项验收（12/13 需 `--base`，`--stage s6` 记 EXPECT） |
| `docx-formula-image-to-omml` | 工具实现层 | 图片公式→OMML、题注/引用/清单/字符级修复脚本（§9–§17 技术细节） |
| `docx-unify-pipeline` | **总编排层** | 七阶段 S0~S6 端到端串起上面三者，统一 CRS 与门禁 |

被上面这些 skill 调用时按 `## 调用契约` 走；它们也要遵守本文件的 CRS 门禁。

## 目录结构

```
docxfmt/
├── __main__.py     python -m docxfmt 入口
├── cli.py          命令行（run / styles / lint / revise / propose / inspect / spec / fields / verify / polish / captions / facts）
├── spec.py         规范定义（工具的输入契约）+ DEFAULT_REQUIREMENT
├── docxio.py       解包/打包/OOXML 构造
├── styles.py       样式表（按 w:name 索引、移植、几何）
├── recognize.py    语义识别 -> 文档树
├── propose.py      ⭐ 挑出疑点 + 上下文；应用人工/AI 的 overrides
├── render.py       套样式 + 抹直接格式
├── toc.py          目录 / 图目录 / 表目录
├── layout.py       页面 / 分节 / 分页符 / 制表位
├── captions.py     题注 SEQ + 交叉引用 REF + 公式编号
├── insert_captions.py  S3：给「缺题注的图/表」补插题注段（只出单，不含域）
├── polish.py       S4 排版六条硬要求的检测与执行
├── pipeline.py     主流水线（顺序有讲究，别改）；含默认模板回退
├── crs.py          ⭐ CRS 唯一实现：指纹 / CrsBuilder / 落盘 / 渲染 / 裁决读取
├── applykit.py     结构类改动的通用执行框架（supported 分流 + 五栏回执）
├── confirm.py      S1 样式确认单（collect + 单文件可编辑 HTML）
├── factcheck.py    S5b 事实核查（抽取 → 判定回填 → 落 CRS）
├── lint.py         S5 正文校验 P1~P10（只出 CRS，绝不改文档）
├── revise.py       S6 修订落盘（w:ins/w:del + 批注 + 不变量校验 + 回执）
├── paginate.py     渲染级量页（WPS/Word COM，Range.Information 六常量，独立进程）
├── verify.py       三层校验
└── winfield.py     COM 重算域（独立进程，因 win32com 未必在主解释器里）
```

> **⚠️ 两个易混模块（别搞错）**：
> `captions.py`（598 行）= 题注 SEQ / 交叉引用 REF / 公式编号的**生成**；
> `insert_captions.py`（390 行）= **给缺题注的图/表补插题注**（S3 阶段）。

```

spec/                spec 样例
├── default.json               ⭐ 默认模板（无格式要求时的回退基准）
├── from-template.json         照模板排
└── chapter-numbering.json     分章编号 + 公式编号

references/          协议细则（写盘门禁的唯一权威定义）
└── change-request-sheet.md    CRS：schema / HTML 契约 / 授权话术 / 回执模板

templates/           给人用的可勾选页
└── crs.html          单文件交互式勾选页（零外部依赖）
```

## 调用契约

- `inspect` 永远先跑，返回结构统计供判断
- **写盘前必须先出 CRS 并拿到逐条批准**（见「改动审批门禁」）。这是硬门禁
- `styles` / `lint` 是**只读命令**，随时可跑，不产出 docx，只产出清单
- `styles` 之后必须**等老板确认**才 `run`；`lint` 之后必须**等老板勾选**才 `revise`
- `run` 不修改输入文件，输出到 `-o`
- `run` 之后**必须** `fields`，除非用户明确不要域
- `revise` 之后**也要** `fields`（修订里的域缓存同样是旧值）
- 交付前跑 `verify`，把 `[完整性]` 那几行报给用户
- 报告要给出**具体数字**（套了多少条样式、题注多少条、引用多少处、零丢失几项），
  不要只说「已完成」
- 交付必须包含**三栏回执**：已改 / 未勾选未改 / 失败，外加备份路径与 md5
