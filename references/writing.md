# 规范写作

## 写作顺序

先确认内容、用户模板和编号/引用要求，再写正文。没有模板时轻量写作入口采用 A4、四边约 2.54 cm、正文 12 pt，中文宋体/西文 Times New Roman；这不是论文通用强制标准，有模板就按模板修改。

在起草阶段为章节、图表、公式和文献分配**稳定目标名**，与显示序号解耦。例如 `FigureArchitecture` 而不是依赖当前序号命名。收集全文所有引用位置，包括重复引用和先引用后出现的目标。先建立编号源，再插入引用，最后更新字段。只给确实需要引用的内容建立引用句子。

## JSON 写作入口

顶层是 `title`（可选）与 `blocks`。所有示例 JSON 使用 UTF-8。

| block.kind | 必填及常用参数 |
|---|---|
| `paragraph` | `parts`；可选 `bookmark` |
| `heading` | `level`（1–9）、`parts`；`numbered` 默认 true，可选 `bookmark` |
| `caption` | `label`、`bookmark`、`cached`；可选 `sequence`、`chapter`、标题 `parts` |
| `equation` | `math`、`bookmark`、`cached`；可选 `sequence`（默认 Equation）、`chapter` |
| `bibliography` | `parts`、用于引用的 `bookmark` |
| `toc` | 默认标题目录；图/表目录用 `mode: figures/tables` 和实际 `sequence` |
| `table` | `rows`：矩形二维数组，每个单元格是一个 `parts` 数组 |

`parts` 数组中允许以下项目：

```json
[
  "普通文字 mm",
  {"text": "2", "script": "superscript"},
  "；电流 I",
  {"text": "cell", "script": "subscript"},
  "；见",
  {"ref": "FigureArchitecture", "cached": "图1"},
  "；第",
  {"ref": "ChapterMethod", "mode": "number", "cached": "1"},
  "章；文献[",
  {"ref": "ReferenceAlpha", "mode": "number", "cached": "1"},
  "]"
]
```

引用 `mode` 为 `text`（默认，取书签内容）、`number`（取自动列表编号）、`page`（PAGEREF）。`cached` 是未重新计算前的显示缓存，不是自动编号源；编写时按实际顺序设置，编辑顺序改变后重新计算。不能凭空编造页码缓存，未知页码可写“待更新”，交付时明确未分页。

书签名为以 ASCII 字母开头、最长 40 字符的字母数字/下划线标识符。SEQ 名可用中文，但不含空格和字段开关。脚本拒绝重复书签和缺失引用目标，允许目标后出现。

标题每个编号级别对应 Heading1–Heading9，绑定同一多级列表。文献标题等用 `numbered: false`；正文标题层级不能跳跃，不能把参考文献正文也设成章节标题。轻量入口目录覆盖 1–3 级，其他范围需按用户需要调整真实 TOC 指令。

图表题注的书签只包住“标签＋编号”，题名在书签外；公式书签只包住括号编号。按章编号设置 `chapter: true`，`cached` 用 `"1.1"`；前提是此前已有带编号的一级标题。普通连续编号使用 `"1"`。正文图表序列、附录序列要分开；字母附录和特殊学校规则用原生编号/XML 精确实现，不用本入口的十进制标题冒充。

文献正文用独立自动列表，显示为 `1 文献信息`；REF `\n` 返回数字，正文另加方括号。这避免列表括号和引用括号重复。需要文后 `[1]` 等格式时，修改列表与 REF 的显示策略并实际检查更新结果，不能只改缓存。

## 原生数学结构

`math` 是文本、项目数组或单一操作对象，不是 LaTeX 解析器：

```json
[
  {"sub": ["I", {"roman": "cell"}]},
  "=",
  {"frac": ["V", "R"]},
  "+",
  {"sum": ["i=1", "N", {"sup": ["x", "2"]}]}
]
```

支持 `sub`、`sup`、`frac`（两个参数），`sum`（下限、上限、求和主体），`paren`、`bracket`、`abs`（单个表达式）和 `roman`（正体字面文本）。不支持的操作报错；矩阵、积分、联立上下标等用明确的 OMML 扩展或已有数学工具，不用字符串伪装支持。

“cell”等描述性下标和函数名需要正体时显式使用 `roman`，不要自动推断。数学中的指数/下标使用 OMML，正文单位和符号才用 Word 字符上下标。

## 配合通用 DOCX 工具

轻量入口适合结构化初稿。用户指定模板、图片、复杂分节等时，使用本机通用 `docx` Skill 的生成工具，直接在生成阶段实现同等结构，不必绕到“先写错再修复”。

- `docx-js` 的 TextRun 设置 `superScript: true` / `subScript: true`；不能以缩小字号代替。
- 用真正 Heading 样式和多级编号，题注使用实际 SEQ 域，引用使用 REF/PAGEREF 域。只有 InternalHyperlink 不够。
- 使用 Python/XML 时可导入 `scripts\word_native.py`：`run`、`field`、`Bookmarks.wrap`、`math_nodes`、`add_numbering` 构造对应结构。插入后维护包关系、内容类型及 XML 顺序。
- 不通过 `python-docx` 重保存一个未知复杂文档来修复全部问题；未知部件、域、对象和修订可能丢失。既有文档优先保留 ZIP 中未修改部件。

目录创建后，字段计算仍需要 Word 或经过实际验证的兼容引擎。在 Word 中分别更新目录内容、编号源、REF 与页码；检查前向引用、重复文献、跨节页码和题注顺序。未经重新计算的示例缓存不能当成最终排版结果。
