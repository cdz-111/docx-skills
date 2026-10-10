# 七阶段详细剧本（S0 – S6）

每个阶段统一按 **目标 / 输入 / 命令 / 产物 / 闸门 / 断言 / 失败处理** 七栏写。
**闸门（gate）有两个档次**：

- `boss` —— 阶段结束**必须拿到老板明确批准**才能进下一阶段。没批准就停住，不许"顺手继续"。
- `self` —— 阶段自带机器可验的断言（`verify` / `verify_all.py` 全绿）即可推进。

```
S0 侦察与备份 ──┬─→ S1 样式基准确认 ──(boss)──┐
               │                              │
               └─→ S2 编号与引用治理 ─(self)─┤
                                              ↓
                        S3 题注补插与引用落地 ─(boss)─┐
                                                      ↓
                        S4 排版（6 条硬要求）────(boss)─┐
                                                      ↓
                        S5 内容校验 + 事实核查 ─(boss)─┐
                                                      ↓
                        S6 修订落盘 + 交付 / 验收
```

S1 与 S2 是**并行支路**（样式基准与编号治理互不依赖）。S3 起必须串行：
题注编号定了，正文引用才落得下去；引用落完，排版量页才准；排版完成，才轮得到内容校验。

---

## S0 · 侦察与备份（gate = auto）

**目标**：在动任何东西之前，把"这是什么文档、现在什么样、改坏了能回哪去"三件事锁死。

### 命令

```bash
PY=~/.workbuddy/binaries/python/envs/default/Scripts/python.exe
D="C:/Users/14711/.workbuddy/skills/docx-format"
U="C:/Users/14711/.workbuddy/skills/docx-unify-pipeline/scripts"
DOCX="课题3-中期科技报告.docx"

# 0.1 备份（带时间戳，永不覆盖原文件）
cp "$DOCX" "${DOCX%.docx}_备份_$(date +%Y%m%d_%H%M%S).docx"

# 0.2 建账本（记指纹，后面每一步写盘前都要对得上）
"$PY" "$U/pipeline_state.py" init "$DOCX"

# 0.3 侦察（只读）
"$PY" -m docxfmt inspect "$DOCX" --tree-txt "$DOCX.树.txt"
```

### 产物

| 产物 | 说明 |
|---|---|
| `<文档名>_备份_<时间戳>.docx` | 回滚点，**必须有**，md5 要记下来 |
| `_unify/pipeline_state.json` | 状态账本，含文档指纹 |
| `<文档名>.树.txt` | 段落树，供人眼扫结构和层级 |

### 断言（进入 S1 前必须全过）

1. 备份文件存在，且 md5 与原文件**不同路径、内容一致**（`md5sum` 对得上）。
2. 账本里 `doc.md5` 与当前文件一致。
3. `inspect` 报出的段落数 / 表格数 / 图片数 / 公式数，与老板口述的规模量级吻合。
   （差很多说明拿错文件了。）

### 失败处理

- **文件被占用**：桌面存在 `~$<文档名>.docx`，或 `open(路径,'r+b')` 抛 `Permission denied`
  → 老板正开着它。**禁止写回原路径**，全程 `-o` 另存新名，并在交付时说清楚。
- **文档是 `.doc`**：先转 `.docx` 再进流程，否则解包即失败。

### 坑

- 备份一定要带**秒级时间戳**。同一分钟跑两次会互相覆盖。
- 账本指纹是**整包 md5**。Word 打开再另存就会变——这不是错误，是 S0 要捕捉的信号。

---

## S1 · 样式基准确认（gate = boss）

**目标**：把"这份文档最终长什么样"在白纸黑字上定死，**老板点头之后**才去套。

老板没给格式要求 → 用**默认模板** `docx-format/spec/default.json`。
老板给了参考文档 → 用 `--template` 照那份列。

### 命令

```bash
D="C:/Users/14711/.workbuddy/skills/docx-format"

# 默认模板
"$PY" -m docxfmt styles "$DOCX"

# 老板给了标杆文档
"$PY" -m docxfmt styles "$DOCX" --template "标杆.docx"

# 已有 spec 则优先用 spec
"$PY" -m docxfmt styles "$DOCX" --spec my.json
```

### 产物

| 产物 | 谁用 |
|---|---|
| `<文档名>_样式确认单.html` | **老板**：逐部分看目标样式，直接改（编号方式/目录层级/分页/表格边框/空格开关），改完导出 JSON |
| `<文档名>_样式确认单.json` | **机器**：确认后原样当 `--spec` 喂回 `run` |

### 必须转述给老板的两行

`styles` 会打印两行，**不能吞掉**：

- `目标样式走别名命中的`：说明 spec 里写的中文名在文档里实际叫什么
  （`标题 1→heading 1`、`正文→Normal`、`题注→caption`）。命中 = 安全。
- `目标样式在文档里找不到的部分`：这部分会走 `new_styles` **新建样式**。
  要明确告诉老板"这些样式你文档里原本没有，我会新建；界面上看着不对就在确认页里改"。

### 闸门

```bash
"$PY" "$U/pipeline_state.py" set S1 blocked --note "样式确认单已出，等老板勾"
# ... 老板回话 ...
"$PY" "$U/pipeline_state.py" gate S1 approved --by "老板" \
      --note "全部同意 / 改：目录只到 2 级，正文改成小四"
"$PY" "$U/pipeline_state.py" set S1 done
"$PY" "$U/pipeline_state.py" artifact S1 "<文档名>_样式确认单.json"
```

### 失败处理

- **15 个部分全报 MISS**：说明 spec 的中文名没被别名表覆盖。检查
  `docxfmt/styles.py` 的 `NAME_ALIASES`，或改成本文档实际用的 `w:name`。不要瞎猜。
- **老板说"你看着定"**：把默认模板的**具体数值**念给老板确认（正文宋体小四 / 1.5 倍行距 /
  首行缩进 2 字符 / 标题 1 黑体三号等），拿到"可以"再推进。**"你看着定"不是批准。**

### 坑

- 样式**一律按 `w:name` 语义名**，绝不按 `styleId`（WPS 保存会把语义 styleId 重写成数字，
  直接复制 `styles.xml` 会导致「题注」段变成「标题 4」）。
- 这份确认单只列**不套**。看到"哪儿不对"就在确认页改，**不要先去 `run` 一遍试试**。

---

## S2 · 编号与交叉引用治理（gate = self）

**目标**：把 SEQ / 书签 / REF / TOC 四层耦合的编号系统捋顺。
**完整剧本在 `docx-numbering-xref-pipeline`，本阶段只做调度与验收。**

### 命令

```bash
N="C:/Users/14711/.workbuddy/skills/docx-numbering-xref-pipeline"
S="$N/../docx-formula-image-to-omml/scripts"

# 2.1 四条审计全跑（只读）—— 少跑一条就会漏一整类病灶
"$PY" "$S/audit_captions.py"        "$DOCX"
"$PY" "$S/audit_ref_semantics.py"   "$DOCX"
"$PY" "$S/recon_hardref.py"         "$DOCX"
"$PY" "$S/build_caption_ref_map.py" "$DOCX"

# 2.2 分层修复（顺序不可换，每步 --dry 先跑）
"$PY" "$S/fix_captions.py"            ...   # ① 重编号 ② 硬写题注补 SEQ ③ 断链重指向
"$PY" "$S/fix_ref_bookmark_span.py"   ...   # ④ _Ref 书签收窄
"$PY" "$S/fix_hardref_to_ref_field.py ..."  # ⑨ 正文纯文本图表号 → 真 REF 域（必须最后）

# 2.3 验收（13 项，任一 FAIL → 退出码 2，禁止交付）
#      --stage s6 时，第 11/12/13 项在给了 --base 且与基线存在差异时记为 EXPECT（不计 FAIL，退出码仍 0）
"$PY" "$N/scripts/verify_all.py" "$DOCX" --base "<备份>.docx" --stage s6
```

### 产物

| 产物 | 说明 |
|---|---|
| 修复后的 docx | 编号自洽 |
| `verify_all.py` 输出 | 13 项检查结果，**必须全绿**（`--stage s6` 时，第 11/12/13 项在给了 `--base` 且与基线有差异时记为 EXPECT，不计 FAIL，仍算全绿） |

### 断言（gate = self，靠 verify_all 把关）

| # | 断言 | 判据 |
|---|---|---|
| 1 | 书签配平 | `bookmarkStart == bookmarkEnd` |
| 2 | 域三态配平 | `begin == separate == end` |
| 3 | 域嵌套合法 | 结束深度 0、最小深度 ≥ 0、`separate` 不重复 |
| 4 | REF 断链 0 | 每个 REF 的目标书签都存在 |
| 5 | REF 缓存 == 语义重算值 | **不能只比"缓存文本是不是 `图 N`"** |
| 6 | `_Ref` 过宽书签 0 | 跨度内不得同时含 SEQ 与 `_Toc` 的 bookmarkStart |
| 7 | 题注编号连续 | 图 `1..N`、表 `1..M` 无重复无跳号 |
| 8 | 清单条目数 == 题注数 | 且每条 `PAGEREF _Toc####` 锚点存在、显示号 == 锚点题注号 |
| 9 | 结构计数 | `w:tbl` / `w:drawing` / SEQ / 段落数（给 `--base` 则逐项对比） |
| 10 | 除 `document.xml` 外部件逐字节一致 | 这是"没动过图片/样式"的唯一硬证据 |

### 失败处理

- **`verify_all` 报部件不一致但差异只集中在 `docProps/*`、`fontTable.xml`、
  `word/media/*` 的重命名**：那不是坏了，是 Word/WPS"另存为"的必然副作用。
  详见 `docx-numbering-xref-pipeline/SKILL.md` §4.1。
- **`SEQ` 域计数变少**：题注被删了。**必须重跑①②③⑥⑨**，不能只补一部分。

### 坑（浓缩版，完整版见编号 skill §7）

1. **域可以跨段落**，「某段文字是否在域里」必须按全篇字符级区间用栈判定。
   逐段 `find(separate)` 会把图目录条目当正文，把统计全带偏。
2. **域指令会被 Word 拆进多个 `<w:instrText>` run**（`' SEQ '` + `'图'` + `' \* ARABIC'`）。
   用 `re.findall(r'SEQ\s+图', xml)` 扫原文会**整批漏判**。必须按 run 顺序拼接重建。
3. **重建 XML 不能 `''.join(段落切片)`**（丢 `<w:body>` 包装 → `unbound prefix`）。
   必须按 `(start, end, newtext)` **倒序区间替换**。
4. **`paragraph._p.clear_content()` 会静默铲掉段落里的书签** → 已有交叉引用当场全断。
   重写题注必须**原地替换**或重建时补回书签。

---

## S3 · 题注补插与引用落地（gate = boss）

**目标**：给"该有题注但没有"的图/表补上题注；把正文里手打的「图N / 表N」纯文本
换成真正能自动更新的 REF 域。

### 命令

```bash
# 3.1 先出清单：哪些图/表没有题注（只出 CRS，不落盘）
#     ⚠️ 已并入 docx-format，命令名是 captions（不是旧的 insert_missing_captions.py）
( cd "$D/skill" && "$PY" -m docxfmt captions "$DOCX" \
    --out "$DOCX_题注事项.json" --html "$DOCX_题注事项.html" )

# 3.2 归集成 CRS（合进统一清单）
"$PY" "$U/crs_merge.py" --docx "$DOCX" \
    --in "$DOCX_题注事项.json" --in "$DOCX_编号治理.json" \
    -o "$DOCX_改动事项.json" --html "$DOCX_改动事项.html"

# 🛑 等老板勾选，导出 approved.json

# 3.3 执行：补题注会增 w:p，必须放在 revise / polish 之后（最后一步）
"$PY" -m docxfmt captions "$DOCX_排版.docx" --apply \
    --crs "$DOCX_改动事项.json" --approved "$DOCX_approved.json" \
    --out-docx "$DOCX_最终.docx" --receipt "$DOCX_题注回执.md"
"$PY" "$N/scripts/verify_all.py" "$DOCX_最终.docx" --base "$DOCX"
```

> 编号形态（`1-1` 分章 / `1` 连续）来自 spec 的 `numbering.caption` + `numbering.sep`，
> 命令本身不传参，跟 S1 确认单保持一致即可。

### 缩进/编号口径（老板的原始要求）

- 编号形态**要能选**：`1-1`（分章）还是 `1`（全文连续）。
  由 spec 的 `numbering.caption`（`chapter` / `global`）+ `numbering.sep` 决定。
- 「图」「表」与编号之间**保留一个空格**（`图 1`），不参与正文去空格。见 S4-②。

### 断言

1. 补的题注数 == 清单里批准的条数。
2. 每条新题注**带 SEQ 域**（不是死文本），且带 `_Ref` 书签。
3. 新增 REF 域数 == 基线 + 批准的硬写转域数；**可见文本只允许空格变化**。
4. `verify_all` 全绿。

### 失败处理

- **图/表在图目录里已有条目，但正文没有题注**：先查是不是"题注挂在图内文本框里"。
  文本域里的题注 `SEQ` 域不参与主文档编号 → 需要人工挪出来，工具改不了。
- **编号后图目录刷新出新旧两套**：`TOC \c "图"` 的域范围被撑破了。重跑 S2 的 ⑥。

### 坑

- **永远不要整段重写公式段**，会把 `m:oMath` 删掉。只删末尾编号 run。
- 补题注必须在**图下方**（`caption_position: below`），插在图上方会让题注跑到上一页。

---

## S4 · 排版：六条硬要求（gate = boss）

**这是老板最在意的阶段，也是唯一需要"渲染级"量测的阶段。**
六条里有四条（③④⑤⑥）**静态 XML 量不出来**——必须知道"这一页有多少行、
这一行有没有塞下、题注和它对应的表在不在同一页"。所以实现走
**静态 + 渲染级混合**：能静态判的静态判（快、可回归），
必须看分页的走 Word/WPS COM 导出 PDF 后量页。

### 六条要求 → 检测方法 → 修法

| # | 要求 | 检测（静态） | 检测（渲染级） | 修法 |
|---|---|---|---|---|
| ① | 文档 > 100 页时，**一级标题分页** | 数 `w:p` + `w:br type=page` 概算页数 | PDF 实际页数 | `pagebreak.before_heading=[1]`；≤100 页则不插 |
| ② | **数字/英文与中文之间无空格**；公式与中文之间无空格 | 正则（见下） | — | `spacing.normalize_cjk_latin`；题注例外 |
| ③ | 正文每页**不出现 ≥4 行空白** | 判不了 | **PDF 逐页量行**：页内最后一行 y 坐标 → 页底距离 | 就近扩/缩正文；或调该页图片高度 |
| ④ | **不出现单字加标点独占一行** | 判不了 | PDF 文本层：行内字符数 == 1 且以标点结尾 | 同上（微调段前段后、字距、图片宽度） |
| ⑤ | **题注与对应部分同页** | 判不了 | PDF：题注段 y 与其目标图/表的 y 是否同页 | 调相对位置（题注与图之间加 `keep_with_next`）；扩缩正文 |
| ⑥ | **表格美化**：数据表三线表，文本表全框，行宽高合审美 | 数每格字符数、算文本占比 | — | 改 `tblBorders` 几何 |

### 命令

```bash
# 4.0 先量页（③④⑤ 全靠它；不给 --pages 时 polish 会自动调，64MB 约 1 分钟）
"$PY" "$D/skill/docxfmt/paginate.py" "$DOCX" -o "$DOCX_量页.json"

# 4.1 出六条的单（只出 CRS，不落盘）
( cd "$D/skill" && "$PY" -m docxfmt polish "$DOCX" --pages "$DOCX_量页.json" \
    --out "$DOCX_排版事项.json" --html "$DOCX_排版事项.html" )

# 🛑 等老板逐条勾选（六条各自成行，可只选 ②⑥）

# 4.2 执行结构类（分页符 + 表格边框）——必须在 revise 之后、captions 之前
"$PY" -m docxfmt polish "$DOCX_修订.docx" --apply \
    --crs "$DOCX_改动事项.json" --approved "$DOCX_approved.json" \
    --out-docx "$DOCX_排版.docx" --receipt "$DOCX_排版回执.md"
```

**为什么必须先量页**：③④⑤ 的判据是「这一页有多少行、这一行塞下没有、题注和它在不在同一页」，
静态读 `document.xml` 判不出来。量页走 Word/WPS COM（`Range.Information` 六常量），
产出的 json 里每段带 `page/y/ln/x/page_end/ye/xe/ln_end`——**末字符坐标**是判④的关键。

---

### ① 一级标题分页：**先看总页数**

**仅在总页数 > 100 时**才给一级标题设 `w:pageBreakBefore`；≤100 页不插
（插了反而把短文档撑散）。`PolisherApplier` 支持 `pagebreak` 执行器，`allowed_delta`
里 `w:p`/`w:tbl`/`w:drawing` 全 0（分页符不算增删结构）。

⚠️ **检测与执行必须用同一棵 lxml 树**：另起 `etree.parse()` 会造成第二棵树，
`paras.index(p)` 用 `==` 是身份比较、永远失配 → 段号全变 0 → 执行时全报越界。
正确做法：同一次解析 + `pos = {id(p): i for i, p in enumerate(paras)}`。

### ② 空格规则的确切边界（**必须按这个来，别扩大打击面**）

去空格的正则只有两条，方向对称：

```python
CJK = "\u4e00-\u9fff\u3400-\u4dbf"          # 汉字
RE_CJK_SPACE   = re.compile(r"([%s])([ \u3000]+)([0-9A-Za-z])" % CJK)  # 中文 + 空格 + 拉丁
RE_LATIN_SPACE = re.compile(r"([0-9A-Za-z])([ \u3000]+)([%s])" % CJK)  # 拉丁 + 空格 + 中文
```

**例外清单（这些空格是体例要求，不许去）**：

| 场景 | 例 | 处理 |
|---|---|---|
| 图表引用「图/表 + 编号」 | `图 1`、`表 2-1`、`式 (3)`、`附图 3` | **保留**（`RE_KEEP_SPACE`，**任何位置**，不只题注段） |
| 公式编号的括号与编号 | `（1）` | 保留 |
| 英文单词之间 | `deep learning` | 本来就该有 |
| 单位与数字之间（老板明确要求时） | `5 kg` | 按老板口径，默认不去 |
| 参考文献著录 | `[1] 张三.` | 保留 |

> ⛔ **实测教训（别重犯）**：第一版只把「题注段」列为例外，结果正文里的
> `如图 1所示` 被去成 `如图1所示`。而 `图 1` 同时是 **REF 域的缓存文本** ——
> 缓存变成 `图1`、题注仍是 `图 1`，两者不一致：
> 一按 F9 就被 Word 改回来，且编号治理的 `verify_all` 第 6 项
> （REF 缓存 == 语义重算值）**直接 FAIL 95 处**。
> 所以例外必须落在**模式**上（`图|表|式|公式|附图|附表 + 空格 + 数字`），
> 与它出现在哪一段无关。`revise` 另有一道兜底护栏：**域结果区内的 run 一律不动**。

> ⚠️ **历史冲突提醒**：上一轮的 CR-004 曾建议 `图 1所示` → `图 1 所示`（**加**空格）。
> 本次规则要求"数字与中文之间**不要**空格"。两者精确矛盾。
> **已确认的口径是：正文去空格，题注与图表引用保持「图 1」**。
> 即：`精度 0.8554` → `精度0.8554`（去），`图 1` 一律不动（含正文交叉引用）。
> 执行前把这条口径念一遍给老板，因为它是**唯一一条会推翻上一轮结论的规则**。

### ⑥ 表格美化的判定与几何

**判定**（`tables` 节）：

```jsonc
"tables": {
  "border_style": "auto",          // auto=按内容判
  "three_line_data_table": true,   // 数据表 → 三线表
  "full_grid_text_table": true,    // 文本表 → 全框
  "text_table_min_chars": 60       // 单格平均字符数超过它 → 算"文本表"
}
```

判据：**算全表每格的平均字符数与最大字符数**。

- 平均 < 阈值 且 单元格多为数字/短词 → **数据表 → 三线表**
  （`tblBorders`：`top`/`bottom` 单线 1.5pt，`insideH` 0.75pt，其余 `nil`；
  首行 `bottom` 加粗作表头横线）
- 平均 ≥ 阈值 或 存在长句单元格 → **文本表 → 全框**
  （`tblBorders` 四边 + `insideH` + `insideV` 全 0.5pt 单线）

**审美几何**（三条经验值，别拍脑袋）：

1. **列宽按内容配比**：给每列算 `max(表头字符数, 该列最长单元格字符数)`，
   按比例分配总宽（`w:tblW` 百分比），**不要均分**——均分会让长文本列挤成竖排。
2. **行首列窄、说明列宽**：第一列（序号/项目名）通常 ≤ 20% 总宽。
3. **行高**：单行行高 = 正文行距 + 2pt；有换行的单元格允许自然撑高，
   但**不要给 `w:trHeight` 写死固定值**（`hRule="exact"`），内容会被裁掉。

**横向宽表**：保持 `section.orientation = "keep"`，**横向节别动**。

> ⚠️ **边框的真实来源是 `w:tcPr/w:tcBorders`（实测）**。
> 单元格级边框**优先级高于**表级 `w:tblBorders`。本机 52 张表实测：多张表连
> `tblBorders` 都没有，边框全在单元格上；`w:tblStyle` 是 WPS 重写过的数字 styleId
> （41/29/30），**不可反查语义名**。
> 所以：**判定**要先看 `tblBorders`、null 时再汇总各格 `tcBorders`；
> **执行**必须先**清掉所有 `tcBorders`**，否则改表级等于没改。
>
> 几何量：三线表 `top`/`bottom` 单线 1.5pt(val=12) + `insideH` 0.75pt(val=6)、其余 `nil`；
> 全框表六边 0.5pt(val=4)。

### 闸门

六条硬要求**逐条出 CRS 行**（不合成一条"排版优化"）。老板可能只要 ②⑥、不要 ①③④⑤，
所以每条必须可单独勾。风险分级：

| 要求 | risk | 理由 |
|---|---|---|
| ① 一级标题分页 | `mid` | 动分页符，会牵连页码与目录 |
| ② 去空格 | `low` | 纯文字 |
| ③ 空白/④ 孤字行 | `mid` | 会改段前段后与图片尺寸 |
| ⑤ 题注同页 | `mid` | 动相对位置 |
| ⑥ 表格边框 | `mid` | 动 `tblBorders`，不改内容 |

### 失败处理

- **COM 拿不到 Word/WPS**：③④⑤ 降级为"只出建议位置清单，让老板在 Word 里手工挪"。
  **不要**瞎猜分页结果然后自动改。
- **一页空白 ≥4 行但原因是"章末"**：章末留白是正常排版。判据要加
  "该页不是本章最后一页"。

---

## S5 · 内容校验与事实核查（gate = boss）

**目标**：正文（**不含公式、附录**）的文字与数据自洽；事实性内容（奖项、新闻、官方通报）
联网核对。**只出 CRS，绝不改文档。**

### 命令

```bash
# 5a 正文校验（P1~P10，只出清单）
"$PY" -m docxfmt lint "$DOCX" --spec "$DOCX_样式确认单.json"

# 5b 事实核查（流程见 references/fact-check-protocol.md）
#     ① 抽取：只走文档树，只处理 paragraph；award/news 必须有年份或机构锚点
( cd "$D/skill" && "$PY" -m docxfmt facts "$DOCX" \
    --out "$DOCX_待核查事实.json" --md "$DOCX_事实核查表.md" )
#     ② 人/Agent 联网核查后写 verdicts.json，回填落 CRS
( cd "$D/skill" && "$PY" -m docxfmt facts "$DOCX" \
    --verdicts verdicts.json --to-crs "$DOCX_事实核查.json" --html "$DOCX_事实核查.html" )

# 5c 合并所有来源为唯一一份清单
"$PY" "$U/crs_merge.py" --docx "$DOCX" \
    --in "$DOCX.改动事项.json" --in "$DOCX_事实核查.json" \
    -o "$DOCX_统一改动事项.json" --html "$DOCX_统一改动事项.html"
```

### 校验范围（老板的原始要求）

| 项 | 范围 | 说明 |
|---|---|---|
| 符号用全角 | 正文 | 不含公式、附录 |
| 符号与文本间无空格 | 正文 | 同上 |
| 错字 / 多字 / 少字 | 正文 | P6/P7 |
| 正文与图表内容相符、数据相符 | 正文 | P9/P10 |
| 交叉引用相符 | 正文 | P9，与 S2 的 REF 验收呼应 |
| 事实性内容联网校验 | **主要是奖项、新闻、官方通报**（**不含小范围事件**） | 见 fact-check-protocol |
| 内容通顺性 | 正文 | 归 P8 术语 + 人工读 |

### 断言

1. `lint` 跑完输出的分档计数与清单条目数一致。
2. **凡 `executor=manual` 的项，`revise` 必须跳过**（它们本来就要人来做）。
3. 事实核查的每条结论都带**可点击 URL + 访问日期**；查不到就明确写"未找到公开来源"。

### 失败处理

- **误报**：上一轮报过、本轮不复现的，用 `decision=closed` + `notes="上一轮误报：…"` 留档，
  **不要悄悄删掉**（老板记得你提过）。
- **同一处在两份来源单里建议不同改法**：`crs_merge.py` 会保留档位最高的一条并把冲突
  写进 `baseline.conflicts`。**报给老板裁**，不要自己选。

---

## S6 · 修订落盘与交付（gate = boss）

**目标**：按批准的清单，以**修订模式**写入，逐条**批注修改原因**，产出三栏回执。

### 命令

```bash
"$PY" "$U/pipeline_state.py" check --stage S6     # 写盘前自查（指纹+闸门+产物）

"$PY" -m docxfmt revise "$DOCX" \
    --crs "$DOCX_统一改动事项.json" \
    --approved "$DOCX_approved.json" \
    -o "${DOCX%.docx}_修订.docx" \
    --author "docx-format 校对"

"$PY" -m docxfmt fields "${DOCX%.docx}_修订.docx" -o "${DOCX%.docx}_修订_域重算.docx"
"$PY" -m docxfmt verify "${DOCX%.docx}_修订_域重算.docx"
```

### 产物

| 产物 | 说明 |
|---|---|
| `<文档名>_修订.docx` | 修订模式产物，可在 Word 里逐条接受/拒绝 |
| `<文档名>_改动回执.md` | 三栏回执 |

### 断言（`revise._validate()` 强校验，任一不满足即不落盘）

| 量 | 要求 |
|---|---|
| `w:p` / `w:tbl` / `w:drawing` 计数 | **必须不变** |
| `bookmarkStart` / `bookmarkEnd` | **必须不变** |
| `fldChar` 三态 | 配平**且计数不变** |
| `document.xml` | 可 `ET.fromstring` |
| `settings.xml` `trackChanges` | 已置 true |
| 新增部件 | 只允许 `word/comments.xml`（+ 必要的 Content_Types / rels 行） |

### 三栏回执（**缺一不可**）

```markdown
# 改动回执 —— <文档名>

- 依据：<文档名>_统一改动事项.json（target md5 `...`）
- 备份：<文档名>_备份_20261009_101530.docx　md5 `...`
- 结果：已改 14 / 未勾选未改 5 / 失败 0

## 一、已改（14）
| CR-ID | 位置 | 改前 | 改后 |
## 二、未勾选未改（5）
| CR-ID | 档 | 位置 | 摘要 |
## 三、失败（0）
| CR-ID | 原因 |
## 四、安全
- 备份路径与 md5（上）
- 域未重算：需在 Word 里 Ctrl+A → F9 一次
```

### 失败处理

- **`_Skip`（跨 run 匹配失败）**：该条进"失败"栏。这是**宁可不改也不改坏**，
  不是 bug。需人工处理。
- **`exit 3`**：没出单就执行 → 拒绝。
- **`exit 5`**：文档在出单后被改过 → **重新出单**，禁止套用旧单。
- **交付后老板说"目录页码不对"**：修订产物同样是旧域缓存，**必须再跑一次 `fields`**。

### 坑

- `revise` 只改文字。**它不做排版、不改编号结构**——那些是 S1/S2/S3/S4 的事。
- 交付前**一定**把「备份路径 + md5」写进回执。老板回滚时唯一能靠的就是这一行。
