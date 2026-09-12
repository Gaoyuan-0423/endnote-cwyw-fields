# EndNote CWYW Fields

**[English](#english)** | **[中文](#中文文档)**

Audit, rebuild, and verify editable EndNote **Cite While You Write (CWYW)** citation fields in Microsoft Word `.docx` files.

A visible `[12]` is **not** proof of a working citation. After a round of revisions, tracked changes, or a copy-paste between documents, the numbers often survive as plain text while the EndNote fields behind them are gone — the manuscript still *looks* correct, but Word and EndNote no longer manage it. This toolkit detects that condition, and rebuilds genuine CWYW fields from records it can prove are correct.

---

## English

### The problem

| Symptom | What it means |
| --- | --- |
| Citations print as `[12]` but the EndNote tab cannot reformat them | The field is gone; only text remains |
| **Update Citations and Bibliography** does nothing, or empties the bibliography | Field linkage was destroyed |
| The `EndNoteBibliography` paragraph style is still applied | Style ≠ field; this is not evidence of a live citation |
| The reference list looks right, the in-text numbers do not renumber | Bibliography is plain text, not an `ADDIN EN.REFLIST` field |

### What the toolkit does

1. **Audits** a DOCX and counts, separately: active `ADDIN EN.CITE` fields, paired `ADDIN EN.CITE.DATA` fields, `ADDIN EN.REFLIST` bibliography fields, and *plain* numeric citation groups.
2. **Resolves** each cited number → bibliography entry → EndNote record, matching on normalized DOI rather than on numbering (bibliography number 27 is **not** EndNote record 27).
3. **Rebuilds** plain groups into real Word complex fields carrying embedded traveling-library `<record>` XML, so they are portable and editable in EndNote.
4. **Fails closed.** If any citation cannot be resolved uniquely and deterministically, nothing is written.

### Features

- Read-only audit with optional JSON report and per-citation locations.
- Dry-run planning mode that reports every unresolved citation before it writes anything.
- Record sources: embedded traveling-library records in the target or earlier DOCX files, and `.enl` EndNote libraries.
- Repeatable `--source-docx` / `--library` flags with explicit precedence for duplicate DOIs.
- Optional JSON DOI map for bibliography entries that carry no DOI and cannot be matched automatically.
- Never overwrites the input; refuses to run on documents containing tracked changes.
- Post-write self-audit that aborts if field counts do not match the visible citation groups.
- Packaged as an agent Skill (`SKILL.md` + `agents/openai.yaml`) alongside the plain CLI scripts.

### Repository layout

```
.
├── LICENSE                         # MIT license
├── SKILL.md                        # Agent skill definition and 7-step workflow
├── agents/openai.yaml              # Skill interface metadata
├── references/ooxml-safety.md      # OOXML field structure, failure conditions, manual recovery
├── requirements.txt                # lxml>=4.9
└── scripts/
    ├── audit_endnote_fields.py     # Read-only structural audit
    ├── rebuild_endnote_fields.py   # Dry-run planner + fail-closed rebuild
    └── endnote_ooxml.py            # Shared DOCX / record / matching / OOXML implementation
```

### Requirements

- Python 3.10 or newer
- [`lxml`](https://pypi.org/project/lxml/) (`pip install -r requirements.txt`)

```bash
git clone https://github.com/Gaoyuan-0423/endnote-cwyw-fields.git
cd endnote-cwyw-fields
pip install -r requirements.txt
```

> `.gitignore` excludes `*.docx`, `*.enl`, and `*.Data/`. Manuscripts and EndNote libraries are never committed — supply your own paths.

### Quick start

Locate four things first: the manuscript to fix, the best pre-edit DOCX that still contains EndNote fields, the `.enl` library that should manage the manuscript, and a clean (changes-accepted) copy of the target.

```bash
cd scripts

# 1. Audit every candidate document.
python audit_endnote_fields.py manuscript.docx --locations --json audit.json

# 2. Dry-run the plan. Nothing is written.
python rebuild_endnote_fields.py manuscript_clean.docx \
  --source-docx manuscript.docx \
  --library manuscript.enl

# 3. If every citation resolves uniquely, write a new file.
python rebuild_endnote_fields.py manuscript_clean.docx \
  --source-docx manuscript.docx \
  --library manuscript.enl \
  --output manuscript_endnote.docx

# 4. Audit the output.
python audit_endnote_fields.py manuscript_endnote.docx
```

Then open the output in Word with the EndNote CWYW add-in and complete the round-trip described below. **Structural validation is not application validation.**

### Command reference

#### `audit_endnote_fields.py`

| Option | Meaning |
| --- | --- |
| `docx` | Word DOCX to audit (required) |
| `--json PATH` | Write the full report as JSON |
| `--locations` | Print every visible citation with paragraph index, body/table, and parsed numbers |

Reported fields: `visible_citation_groups`, `active_endnote_cite_fields`, `active_endnote_data_fields`, `plain_citation_groups`, `endnote_reflist_fields`, `tracked_changes_present`, `embedded_record_candidates`, `embedded_unique_dois`, `mixed_paragraphs`, `citation_locations`.

#### `rebuild_endnote_fields.py`

| Option | Meaning |
| --- | --- |
| `docx` | Clean DOCX containing numeric citations (required) |
| `--source-docx PATH` | DOCX holding original traveling-library records; repeatable |
| `--library PATH` | EndNote `.enl` library, opened read-only via a temporary copy; repeatable |
| `--reference-map PATH` | JSON map from bibliography number to DOI |
| `--output PATH` | New output DOCX; **omit for dry-run planning** |
| `--report PATH` | Write the plan/result JSON report |

Order matters: **add the intended managing library before any supplemental library**, because for a duplicate DOI the first source supplied wins.

### Reference map format

Use this only for entries that carry no DOI in the bibliography text. Both forms are accepted:

```json
{
  "12": "10.1038/s41586-021-00000-0",
  "27": { "doi": "doi:10.1016/j.cell.2019.01.001" }
}
```

### Exit codes

| Script | Code | Meaning |
| --- | --- | --- |
| audit | `0` | No plain citation groups or mixed paragraphs found |
| audit | `1` | Plain citation groups and/or mixed paragraphs present |
| audit | `2` | The document could not be read as a supported DOCX |
| rebuild | `0` | Dry-run ready to write, or write completed |
| rebuild | `2` | Aborted by an integrity error (e.g. tracked changes, unreadable input) |
| rebuild | `3` | Unresolved citations; stopped without writing |
| rebuild | `4` | Field-flattening or conversion error; stopped without writing |

### Fail-closed conditions

The rebuild stops without writing when any of the following applies:

- the document contains tracked insertions, deletions, or moves;
- an existing EndNote field is incomplete or has no numeric display;
- a bracketed group cannot be parsed as a positive integer list or range;
- a cited number has no bibliography entry;
- a bibliography entry has neither a DOI nor an explicit map entry;
- no unique record matches that DOI;
- a citation sits inside a hyperlink or a non-splittable run;
- the output path equals the input path or already exists.

These are integrity failures. They are not invitations to guess a record or fabricate a field.

### `.enl` support and limits

- The library is opened through a **temporary copy** and queried with SQLite; the scripts never modify an EndNote library.
- The reader builds records for EndNote's **Journal Article** reference type only. Other reference types are skipped rather than coerced into a generic type — supply a DOCX that still holds their traveling-library record instead.
- Direct edits to the `.enl` SQLite database are not supported.

### Validation levels

| Level | Meaning |
| --- | --- |
| **Audit verified** | Field instructions and visible citation groups were counted correctly |
| **Structurally rebuilt** | Every target group has paired `EN.CITE` / `EN.CITE.DATA` markup with embedded record XML, and visible text is unchanged |
| **EndNote validated** | The file survived an Update → Save → Close → Reopen cycle in Word with the EndNote add-in, and passed a second audit |

Claim the last level only after the application round-trip.

### Word / EndNote round-trip

1. Open the intended EndNote library, then the rebuilt DOCX.
2. In Word's **EndNote** tab, run **Update Citations and Bibliography** once.
3. Resolve any matching-reference prompt by DOI and title; reject incorrect duplicates.
4. Save, close, reopen, update once more, and audit the saved DOCX.

If Word or EndNote drops citations, changes numbering unexpectedly, or cannot read the traveling records, discard the generated copy and return to the untouched input.

If records are missing, import them into the intended library via DOI, PubMed, RIS, or EndNote XML, verify the metadata inside EndNote, and rerun the dry plan.

### Further reading

Field structure, failure conditions, and the manual recovery procedure are documented in [`references/ooxml-safety.md`](references/ooxml-safety.md).

### License

Released under the [MIT License](LICENSE). Copyright (c) 2026 Gaoyuan.

---

## 中文文档

**[English](#english)** | **[中文](#中文文档)**

审计、重建并验证 Microsoft Word `.docx` 文档中**可编辑的 EndNote 边写边引（Cite While You Write, CWYW）引文字段**。

文档里显示为 `[12]` **并不代表**这条引文是可用的。经过一轮修订、开启修订痕迹、或在文档之间复制粘贴之后，编号常常只是以**纯文本**形式保留下来，而其背后的 EndNote 字段已经丢失——稿件看起来依然正确，但 Word 与 EndNote 已经无法再管理它。本工具用于检测这种状态，并在**能够证明记录正确**的前提下重建真正的 CWYW 字段。

### 要解决的问题

| 现象 | 实际含义 |
| --- | --- |
| 引文显示为 `[12]`，但 EndNote 选项卡无法重新格式化 | 字段已丢失，只剩下文本 |
| 点击 **Update Citations and Bibliography** 无反应，或参考文献表被清空 | 字段链接已被破坏 |
| 仍保留 `EndNoteBibliography` 段落样式 | 样式 ≠ 字段，不能作为引用仍存活的证据 |
| 参考文献表看似正常，但正文编号无法重新编号 | 参考文献表是纯文本，而非 `ADDIN EN.REFLIST` 字段 |

### 工具做什么

1. **审计**：分别统计有效的 `ADDIN EN.CITE` 字段、配对的 `ADDIN EN.CITE.DATA` 字段、`ADDIN EN.REFLIST` 参考文献字段，以及**纯文本**数字引文分组。
2. **解析**：正文编号 → 参考文献条目 → EndNote 记录。匹配依据是**规范化 DOI**，而不是编号本身（参考文献第 27 条**不等于** EndNote 记录 27 号）。
3. **重建**：把纯文本分组转换为真正的 Word 复杂域，内嵌 traveling library 的 `<record>` XML，因此在 EndNote 中可移植、可编辑。
4. **失败即停**：只要有任何一条引文无法唯一且确定地解析，就不写入任何文件。

### 功能特性

- 只读审计，支持导出 JSON 报告与逐条引文位置。
- 试运行（dry-run）规划模式，在真正写入之前列出全部未解析引文。
- 记录来源：目标或早期 DOCX 中内嵌的 traveling library 记录，以及 `.enl` EndNote 文库。
- `--source-docx` / `--library` 可重复指定，并为重复 DOI 规定了明确的优先顺序。
- 可选 JSON DOI 映射表，用于参考文献条目本身不含 DOI、无法自动匹配的情况。
- 绝不覆盖输入文件；文档存在修订痕迹时拒绝运行。
- 写入后自动复检，字段数量与可见引文分组不一致时直接中止。
- 既可作为命令行脚本使用，也打包为 agent Skill（`SKILL.md` + `agents/openai.yaml`）。

### 目录结构

```
.
├── LICENSE                         # MIT 许可证
├── SKILL.md                        # Agent skill 定义与 7 步工作流
├── agents/openai.yaml              # Skill 接口元数据
├── references/ooxml-safety.md      # OOXML 字段结构、失败条件、人工恢复流程
├── requirements.txt                # lxml>=4.9
└── scripts/
    ├── audit_endnote_fields.py     # 只读结构审计
    ├── rebuild_endnote_fields.py   # 试运行规划 + 失败即停的重建
    └── endnote_ooxml.py            # DOCX / 记录 / 匹配 / OOXML 共享实现
```

### 环境要求

- Python 3.10 或更高版本
- [`lxml`](https://pypi.org/project/lxml/)（`pip install -r requirements.txt`）

```bash
git clone https://github.com/Gaoyuan-0423/endnote-cwyw-fields.git
cd endnote-cwyw-fields
pip install -r requirements.txt
```

> `.gitignore` 已排除 `*.docx`、`*.enl` 与 `*.Data/`。稿件与 EndNote 文库不会被提交，请自行提供路径。

### 快速开始

先准备好四样东西：待修复的稿件、仍含有 EndNote 字段的最佳**修订前 DOCX**、本应管理该稿件的 `.enl` 文库，以及一份**已接受修订的干净副本**。

```bash
cd scripts

# 1. 审计每个候选文档
python audit_endnote_fields.py manuscript.docx --locations --json audit.json

# 2. 试运行规划，不写入任何文件
python rebuild_endnote_fields.py manuscript_clean.docx \
  --source-docx manuscript.docx \
  --library manuscript.enl

# 3. 若所有引文都能唯一解析，则写入新文件
python rebuild_endnote_fields.py manuscript_clean.docx \
  --source-docx manuscript.docx \
  --library manuscript.enl \
  --output manuscript_endnote.docx

# 4. 审计输出文件
python audit_endnote_fields.py manuscript_endnote.docx
```

随后请在装有 EndNote CWYW 插件的 Word 中打开输出文件，完成下文所述的往返验证。**结构验证不等于应用层验证。**

### 命令参考

#### `audit_endnote_fields.py`

| 参数 | 说明 |
| --- | --- |
| `docx` | 待审计的 Word DOCX（必填） |
| `--json PATH` | 将完整报告导出为 JSON |
| `--locations` | 逐条打印可见引文：段落序号、正文/表格、解析出的编号 |

报告中包含：`visible_citation_groups`、`active_endnote_cite_fields`、`active_endnote_data_fields`、`plain_citation_groups`、`endnote_reflist_fields`、`tracked_changes_present`、`embedded_record_candidates`、`embedded_unique_dois`、`mixed_paragraphs`、`citation_locations`。

#### `rebuild_endnote_fields.py`

| 参数 | 说明 |
| --- | --- |
| `docx` | 含数字引文的干净 DOCX（必填） |
| `--source-docx PATH` | 含原始 traveling library 记录的 DOCX；可重复指定 |
| `--library PATH` | EndNote `.enl` 文库，通过临时副本以只读方式打开；可重复指定 |
| `--reference-map PATH` | 参考文献编号到 DOI 的 JSON 映射 |
| `--output PATH` | 新的输出 DOCX；**省略即为试运行规划** |
| `--report PATH` | 导出规划/结果 JSON 报告 |

参数顺序很重要：**应先指定本应管理该稿件的文库，再指定补充文库**，因为对于重复 DOI，最先提供的来源优先。

### 映射表格式

仅用于参考文献正文中不含 DOI 的条目。以下两种写法均可：

```json
{
  "12": "10.1038/s41586-021-00000-0",
  "27": { "doi": "doi:10.1016/j.cell.2019.01.001" }
}
```

### 退出码

| 脚本 | 退出码 | 含义 |
| --- | --- | --- |
| audit | `0` | 未发现纯文本引文分组或混合段落 |
| audit | `1` | 存在纯文本引文分组和/或混合段落 |
| audit | `2` | 文档无法作为受支持的 DOCX 读取 |
| rebuild | `0` | 试运行可写入，或写入已完成 |
| rebuild | `2` | 因完整性问题中止（如存在修订痕迹、输入不可读） |
| rebuild | `3` | 存在未解析引文，未写入即停止 |
| rebuild | `4` | 字段扁平化或转换出错，未写入即停止 |

### 失败即停的条件

出现以下任一情况时，重建过程将**不写入任何文件**并停止：

- 文档包含修订插入、删除或移动；
- 已有 EndNote 字段不完整，或没有数字形式的显示文本；
- 方括号分组无法解析为正整数列表或范围；
- 某个引用编号在参考文献表中不存在；
- 参考文献条目既无 DOI，也无显式映射；
- 没有任何唯一记录与该 DOI 匹配；
- 引文位于超链接或无法安全拆分的数据块内；
- 输出路径与输入路径相同，或该文件已存在。

这些属于完整性问题，**不允许**通过猜测记录或伪造字段来绕过。

### `.enl` 支持与限制

- 文库通过**临时副本**打开并使用 SQLite 查询，脚本**从不会修改** EndNote 文库。
- 记录读取仅支持 EndNote 的 **Journal Article** 文献类型。其他类型会被跳过，而不会被强行转换为通用类型——请改用仍保存其 traveling library 记录的 DOCX。
- 不支持直接编辑 `.enl` 的 SQLite 数据库。

### 验证层级

| 层级 | 含义 |
| --- | --- |
| **审计通过（Audit verified）** | 字段指令与可见引文分组已正确统计 |
| **结构重建（Structurally rebuilt）** | 每个目标分组都具备配对的 `EN.CITE` / `EN.CITE.DATA` 标记与内嵌记录 XML，且可见文本保持不变 |
| **EndNote 验证（EndNote validated）** | 该文件在装有 EndNote 插件的 Word 中经历了 Update → Save → Close → Reopen 往返，并通过了二次审计 |

只有在完成应用层往返之后，才可以宣称最后一层。

### Word / EndNote 往返验证

1. 先打开目标 EndNote 文库，再打开重建后的 DOCX。
2. 在 Word 的 **EndNote** 选项卡中运行一次 **Update Citations and Bibliography**。
3. 依据 DOI 与标题处理匹配文献对话框，拒绝错误的重复项。
4. 保存、关闭、重新打开，再更新一次，并对保存后的 DOCX 重新审计。

如果 Word 或 EndNote 删除了引文、编号异常变化，或无法读取 traveling library 记录，请丢弃生成的文件，回到未经改动的原始输入。

若记录缺失，请通过 DOI、PubMed、RIS 或 EndNote XML 将它们导入目标文库，在 EndNote 内核对元数据后重新运行试运行规划。

### 延伸阅读

字段结构、失败条件与人工恢复流程详见 [`references/ooxml-safety.md`](references/ooxml-safety.md)。

### 许可证

本项目基于 [MIT 许可证](LICENSE) 发布。版权所有 (c) 2026 Gaoyuan。
