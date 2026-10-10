# 改动事项单（CRS）协议细则

本文件是**写盘门禁的唯一权威定义**。`SKILL.md` 里的门禁章节是本文件的摘要；
两者冲突时，以本文件为准。

**不变量**：任何对文档的写盘修改，都必须先有一份 CRS，且其中每一项都有明确的
`decision`。没有 CRS 就没有写盘。

---

## 1. 产物与命名

一次改动 = 一份 CRS，落在**被改文档的同目录**：

| 文件 | 角色 | 谁用 |
|---|---|---|
| `<文档名>_改动事项.json` | **唯一裁决源**。执行侧只读它 | 机器 |
| `<文档名>_改动事项.html` | 可勾选页。复选、筛选、统计、导出 | 人 |
| `<文档名>_改动回执.md` | 执行后生成的回执（见 §10） | 人 |

`.html` 与 `.json` 内容同源：`.html` 只是 `.json` 的一个视图。
两者不一致时，**以 `.json` 为准**，并重新生成 `.html`。

---

## 2. CRS JSON schema

```jsonc
{
  "crs_version": "1.0",

  // 出单时被改文档的指纹。执行前必须复核，不符即拒绝（见 §9）
  "target": {
    "path": "C:\\...\\示例文档.docx",
    "size": 64689234,
    "md5": "68130176a8b31fb3c9aed3e589e2faab",
    "document_xml_md5": "53bbb5bf1a2979be31e8547b339b3747",
    "paragraphs": 1646,
    "captured_at": "2026-10-09T09:55:00+08:00"
  },

  // 出单依据：本单针对上面那个版本
  "baseline": {
    "note": "目标文件指纹若变化，本单全部作废，必须重新出单",
    "previous_round": "2026-10-08 21:08 → 21:43"
  },

  "sources": ["propose", "numbering-xref", "lint"],

  // 闸门状态。执行侧只认这个字段
  "gate": {
    "state": "awaiting_approval",   // awaiting_approval | approved | rejected
    "total": 21,
    "approved": 0,
    "rejected": 0
  },

  "items": [ /* 见 §3 */ ]
}
```

---

## 3. item 结构

```jsonc
{
  "id": "CR-001",              // 必填。CR- 前缀 + 3 位序号，全局唯一、跨轮次稳定
  "grade": "S",                // S | A | B | C，见 §4
  "source": "lint",            // propose | numbering-xref | lint
  "category": "错字",           // 自由短标签：错字 | 缺字 | 语序 | 引用 | 体例 | 数据 | 结构 | 术语

  "where": {                   // 必填。必须能被独立复核
    "section": "2.4.1",
    "para": 1034,              // 段落序号（以本单 target 指纹那一版为准）
    "anchor": "「无b符号比特量化」", // 可搜索的定位串，供人肉搜索
    "count": 1                 // 该项在全文命中几处
  },

  "before": "对于无b符号比特量化",
  "after":  "对于无符号 b 比特量化",
  "basis":  "b 是公式对象，原文把 b 插在「无」与「符号」之间，且多一个「范」字",
  "risk":   "low",             // low | mid | high，见 §5
  "impact": "无",               // 会连带改到哪些；没有就写"无"，不要留空

  "executor": "text_replace",  // 执行器标识：text_replace | style_map | heading_level | renumber | ...
  "decision": "pending",       // pending | approved | rejected | modified | closed
  "after_override": null,      // decision=modified 时，用户改写的建议
  "notes": ""                  // 备注（如"上一轮误报"）
}
```

**`decision` 取值**：

| 值 | 含义 | 执行侧 |
|---|---|---|
| `pending` | 还没表态 | 不动 |
| `approved` | 批准 | **执行** |
| `rejected` | 明确否掉 | 不动，进回执「未勾选未改」 |
| `modified` | 批准但改法被调用方改过 | 用 `after_override` 执行 |
| `closed` | 已修 / 已撤回（误报） | 不动，不进待勾区 |

---

## 4. 档位定义

| 档 | 含义 | 判据 |
|---|---|---|
| **S 必改** | 影响正确性 | 错字、缺字、数据自相矛盾、引用指向错、编号断链 |
| **A 应改** | 影响可读性 / 体例 | 语句不通、搭配不当、术语不统一、体例不一致且已有明确多数的 |
| **B 建议** | 风格优化 | 空格/标点体例统一、措辞润色、可选优化 |
| **C 待核** | 口径未定 | 需要调用方拍板（如"某个方法没有验证实验，补实验还是写明待下一阶段"） |

档位只影响**呈现顺序与默认勾选建议**，不改变门禁：
**S 档也一样要勾，不许"这种明显的错我直接改了"。**

---

## 5. 风险分级

| risk | 定义 | 执行要求 |
|---|---|---|
| `low` | 只动文字，不碰域 / 书签 / 编号 / 样式 | 可整批执行，事后给逐项 diff |
| `mid` | 动样式引用、标题层级、自动编号列表、段落样式 | 逐条 `--dry`，给「改前 / 改后」对照 |
| `high` | 动 SEQ/REF/TOC 域、书签、题注编号、交叉引用 | **单条 dry + `verify_all.py` 全绿**才允许写入 |

风险由**改动会触及什么**决定，不由改动的大小决定。
改一个错别字是 `low`；改一个小标题的样式（会影响目录层级）就是 `mid`。

---

## 6. HTML 勾选页的 DOM 契约

`templates/crs.html` 是模板。生成实例时**只替换数据块**，不要改结构——
`crs_apply` 与任何人写的解析脚本都依赖以下属性。

```html
<!-- 每条必须带这些 data-* -->
<tr class="cr-row"
    data-cr-id="CR-001"
    data-grade="S"
    data-risk="low"
    data-source="lint"
    data-decision="pending">

  <!-- 勾选框：唯一的人机输入点 -->
  <td><input type="checkbox" class="cr-check" data-cr-id="CR-001"></td>
  ...
</tr>

<!-- 按档分组，便于筛选 -->
<tbody data-grade-group="S"> ... </tbody>

<!-- 顶部汇总与控件 -->
<span id="cr-summary"></span>
<button id="cr-export">导出 approved.json</button>

<!-- 数据源：内联 JSON，供页面自举 -->
<script id="crs-data" type="application/json">{ ...整个 CRS json... }</script>
```

**必须支持的交互**（缺一项，人就会用错）：

1. 逐条勾选
2. `全选` / `全不选` / `按档勾选（S / A / B / C）`
3. 按 `grade`、`risk`、`source` 筛选行
4. 实时汇总：`已勾 N / 共 M`、按档计数
5. 导出 `approved.json`（下载 + 复制到剪贴板两条路，因为有的环境不给下载）

**主题**：跟随 `prefers-color-scheme`，深浅两套都要能读。
深色下正文必须浅色、浅色下正文必须深色；不要写死单一配色。

---

## 7. 导出格式（`approved.json`）

导出的是**裁决结果**，不是整份数据。文件名 `<文档名>_approved.json`：

```json
{
  "crs_version": "1.0",
  "target_md5": "68130176a8b31fb3c9aed3e589e2faab",
  "exported_at": "2026-10-09T10:12:33+08:00",
  "items": [
    { "id": "CR-001", "decision": "approved" },
    { "id": "CR-005", "decision": "modified", "after_override": "改写后的建议" },
    { "id": "CR-007", "decision": "rejected" }
  ]
}
```

**未出现在 `items` 里的 ID，一律视为 `pending`（不执行）。**
这样"只导出勾了的"就天然安全。

---

## 8. 批量授权话术 → decision 映射

调用方不必逐条点，一句口头授权同样有效，但**必须落成同一份 decision**：

| 调用方说 | 落成 |
|---|---|
| 「全部同意」「都改」「全勾」 | 全部 `approved`，`gate.state=approved` |
| 「除 CR-007 外都同意」 | 除该项 `rejected`，其余 `approved` |
| 「只改 S 档」 | `grade=S` → `approved`，其余 → `rejected`（明确否掉，不留在 pending） |
| 「先不改」「再想想」 | 全部 `rejected`，`gate.state=rejected` |
| 「CR-005 改成 XXX」 | 该项 `decision=modified`，`after_override="XXX"` |
| 「CR-003 不用了」 | 该项 `rejected` |

**沉默不是同意。** 没有回话就保持 `awaiting_approval`。

---

## 9. 执行侧契约

执行器（`run` / `fields` / 各修复脚本）在写盘前必须依次检查：

| # | 检查 | 不通过时 |
|---|---|---|
| 1 | CRS 文件存在 | `exit 3` —— 「未出改动事项单，拒绝写盘」 |
| 2 | `target.md5` == 当前文件 md5 | `exit 5` —— 「文档已变，本单作废，请重新出单」 |
| 3 | `gate.state == "approved"` 且 `approved > 0` | `exit 4` —— 「清单已就绪 / 已全部否掉，没有可执行项」 |
| 4 | 逐项按 §5 的风险分级执行 | 任一项 `verify` 不绿 → 停止，不写盘 |
| 5 | 写盘前备份 | 无备份不写盘 |

**检查 2 很重要**：文档在出单之后又被 Word/WPS 重存过，段落序号全变，
照旧单执行会改错地方。**指纹不符一律重新出单，不要"差不多"地套用。**

---

## 10. 回执表模板（`<文档名>_改动回执.md`）

执行完必须产出，三栏缺一不可：

```markdown
# 改动回执 —— <文档名>

- 依据：<文档名>_改动事项.json（target md5 `...`）
- 备份：<文档名>_备份_20261009_101530.docx　md5 `...`
- 结果：已改 14 / 未勾选未改 5 / 失败 0

## 一、已改（14）

| CR-ID | 位置 | 改前 | 改后 |
|---|---|---|---|
| CR-001 | 2.4.1 · 段 1034 | 对于无b符号比特量化 | 对于无符号 b 比特量化 |

## 二、未勾选未改（5）

| CR-ID | 档 | 位置 | 摘要 |
|---|---|---|---|
| CR-018 | C | 2.3 | 附录 B 混入 1 条软著 |

## 三、失败（0）

| CR-ID | 原因 |
|---|---|

## 四、安全

- 备份路径与 md5（上）
- 域未重算：需在 Word 里 `Ctrl+A` → `F9` 一次
```

---

## 11. 跨轮次规则

1. **CR-ID 稳定**：同一问题在后续轮次沿用同一 ID；修好的标 `decision=closed`，
   移出待勾区，进「已修」区。这样调用方说"CR-007 还没改"永远对得上。
2. **旧单作废**：目标文档指纹一变（被 Word 重存、被手工改过），
   上一份 CRS 的 `where.para` 全部失效。**重新出单，不要打补丁。**
3. **出单前先核现状**：上一个轮次说"待改"的项，这一轮出单前必须重新验证它是否还在。
   否则清单里会混进已经改好的条目，调用方会失去信任。
4. **误报要显式撤回**：上一轮报错的项，本轮以 `decision=closed` + `notes="上一轮误报：…"`
   留档，别悄悄删掉——调用方记得你提过。
