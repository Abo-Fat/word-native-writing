# word-native-writing

让 AI 撰写的 Word 文档具备可编辑、可更新的原生结构，而不只是视觉上像论文或报告。这个 Skill 同时支持**规范写作**与**已有文档修复**。

## 功能

- 普通文字使用真实 Word 上下标；公式使用原生 OMML 上下标、分式、求和等结构。
- 使用真正的标题样式、大纲层级与多级自动编号。
- 图、表、公式使用 SEQ 编号源和精确的编号书签。
- 图表、公式、章节和页码引用使用 REF/PAGEREF，不以手写数字或超链接代替交叉引用。
- 文献列表自动编号，正文全文复用对应的文献引用目标，支持前向引用。
- 创建真实 TOC 目录，并区分目录条目更新和页码更新。
- 先审计、后按精确计划修复；绑定原稿 SHA256 与段落原文，拒绝覆盖原稿和歧义替换。

Skill 入口：[SKILL.md](SKILL.md)。详细规则见[规范写作](references/writing.md)和[文档修复](references/repair.md)。

## 安装

将本仓库放入 AI 客户端的 Skill 发现目录，并保持目录名为 `word-native-writing`。例如 Windows 上使用 `.agents\skills` 的客户端：

```powershell
git clone https://github.com/Abo-Fat/word-native-writing.git "$env:USERPROFILE\.agents\skills\word-native-writing"
```

目标目录必须尚不存在；已经安装时不要重复克隆或覆盖自己的修改。其他客户端请按其实际 Skill 目录安装。若当前会话没有重新发现技能，请重新打开会话；不会替换已有通用 `docx` Skill。

Python 3.9+。审计、检查、基础解包/打包仅需 Python 标准库；原生写作/修复和 Flat OPC 打包需要 `lxml`。缺失依赖时，在仓库目录执行：

```powershell
python -m pip install -r .\requirements.txt
```

## 使用 Skill

可以直接让 AI 使用 `word-native-writing`：

> 使用 word-native-writing 撰写 Word 报告，正确处理上下标、标题编号、图表题注和文献交叉引用。

> 使用 word-native-writing 审计并修复这份 DOCX，保留原稿和技术内容，将静态编号及引用转换为可更新的 Word 原生结构。

AI 应先阅读对应指南，再根据本次内容生成写作规格或修复计划；示例不能原样套用到任意文档。

## 直接运行工具

以下命令假定当前目录为本仓库根目录。

### 规范写作

```powershell
python -X utf8 .\scripts\docx_native.py write .\assets\writing-example.json "D:\work\Word结构示例.docx"
python -X utf8 .\scripts\check_docx.py "D:\work\Word结构示例.docx"
```

写作规格支持段落、真实上下标、标题、题注、文献、引用、公式、目录和简单文字表格。示例文献是明确标注的占位信息，不可作为真实文献交付。

### 定向修复

```powershell
python -X utf8 .\scripts\audit_docx.py "D:\work\原稿.docx" --output "D:\work\audit.json"
# 根据审计和 assets\repair-plan.example.json 编写本次 repair-plan.json。
python -X utf8 .\scripts\docx_native.py repair "D:\work\原稿.docx" "D:\work\repair-plan.json" "D:\work\修复暂存.docx"
python -X utf8 .\scripts\check_docx.py "D:\work\修复暂存.docx" --source "D:\work\原稿.docx"
```

有模板时，审计命令可增加 `--template "D:\work\模板.docx"`。修复计划必须使用本次审计的 SHA256、段落索引及完整原文，不自动猜测科学符号、引用对象或公式含义。

### 可选 Word 域刷新

在 Windows 安装了桌面 Word，且允许 Word 自动化时：

```powershell
powershell -NoProfile -File .\scripts\refresh_fields.ps1 -InputPath "D:\work\修复暂存.docx" -OutputPath "D:\work\修复完成.docx"
```

逐类刷新已有编号、REF 与页码，保留 Flat OPC 恢复状态。它不会补建缺失引用，也不会重建目录条目；新增或变更目录需要在 Word 中定向更新整个目录。

## 文件结构

| 路径 | 用途 |
|---|---|
| `SKILL.md` | 技能触发条件、规范及执行入口 |
| `references` | 写作与修复指南 |
| `assets` | JSON 写作规格及修复计划示例 |
| `scripts\docx_native.py` | 原生写作与精确计划修复入口 |
| `scripts\word_native.py` | 字符格式、OMML、编号、书签及域的构造函数 |
| `scripts\audit_docx.py` / `check_docx.py` | 只读审计、字段与引用检查、源内容辅助比较 |
| `scripts\docx_package.py` | 标准库解包/打包，不是完整 XSD 校验器 |
| `scripts\refresh_fields.ps1` / `package_word_state.py` | Word 逐类域刷新及 Flat OPC 打包 |
| `scripts\office_skill_bridge.py` | 可选调用本机已有 `docx` Skill 的校验器 |

## 边界

这不是任意文档的一键自动修复器，也不是完整排版引擎。复杂模板、图片、分节、附录及修订需配合通用 DOCX 工具进行精确处理。复杂目标段落不会被展平成纯文本以强行修复。

结构检查和 XSD 校验不能代替公式语义检查、Word 字段计算或逐页视觉复核。没有实际执行的项目必须明确标注未验证。没有 Word 时仍可生成原生结构，但目录、页码和重排后的显示缓存需之后重新计算。

本仓库不包含原始报告、报告正文审计数据、带专用路径的历史案例脚本或第三方 DOCX Skill 源码。不要将自己的原稿、审计结果或敏感修复计划提交到公开仓库。
