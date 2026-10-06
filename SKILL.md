---
name: word-native-writing
description: "规范撰写或修复 Word/.docx 文档：真实上下标、原生 OMML 公式、标题与多级自动编号、图表和公式 SEQ 题注、REF/PAGEREF 交叉引用、全文一致的文献引用与参考文献编号、自动目录。用户要求写 Word、论文或报告，或修复 AI 生成文档中的伪上下标、手写编号、缺失交叉引用、引用失效和标题目录不规范时使用。"
---

# word-native-writing：Word 规范写作与修复

目标：让文档具有 **Word 原生语义结构**，而不只是视觉上相似。支持新文档写作和已有文档修复；修复不得擅自改变技术内容。

这里的“全文引用”指全文一致的文献引用及图、表、公式、章节、页码引用管理，不是把整篇文献复制进文档。

## 入口选择

| 用户任务 | 执行方式 |
|---|---|
| 新写 Word、论文、报告 | 先读 `references\writing.md`；按内容和模板写作，使用原生结构 |
| 修复已有 DOCX | 先读 `references\repair.md`；审计原稿，生成针对该文档的修复计划，再处理副本 |
| 只检查问题 | 运行 `scripts\audit_docx.py` 和 `scripts\check_docx.py`，不修改文档 |
| 复杂排版、图片、页眉页脚或修订 | 与本机已有通用 `docx` Skill 配合；本 Skill 负责编号、引用和科学字符的语义规范 |

**所有脚本路径都相对本 Skill 安装目录，不相对用户文档目录。** 命令示例假设当前目录为 Skill 根目录；否则用绝对脚本路径。

## 不可违反的规范

1. **字符上下标**：普通数字/字母加 `w:vertAlign`，保持基准字号，不用 `²`、`₁` 等字符冒充格式。不把代码、型号、数组索引或音标自动改成上下标；先确认语义。
2. **数学公式**：使用 `m:oMath` 原生 OMML，上下标、分式、求和分别有相应结构。公式内容和编号分开；不能以线性公式解析“不报错”证明语义正确。
3. **标题**：真正的 Heading 样式、大纲层级和多级自动编号。标题文字不重复手写序号；文献标题等可入目录而不参与章节编号。
4. **题注**：图、表、公式建立各自的 `SEQ`。按章编号必须有正确的一级标题编号和 `STYLEREF`/重置策略；附录用独立编号规则，不能混入正文序列。
5. **交叉引用**：先建立编号源和唯一书签，再插入真实 `REF`/`PAGEREF`。普通超链接、书签和手写“见图1”都不能替代可更新的交叉引用。
6. **文献引用**：文后用自动编号列表，每条建稳定书签；正文每个数字用 `REF ... \n \h`。括号和分隔符是普通文字；全文复用同一目标，不重复造号。
7. **目录**：真实 `TOC`，目录内容和目录页码分开更新。新生成目录的缓存不是已计算结果；全文需覆盖前向、重复引用、表格以及适用的脚注/尾注、页眉页脚。
8. **内容保护**：不编造文献、不新增原稿没有的引用论断，不删除全部直接格式，不替换原图、不重建无关表格；保留代码、数字和公式含义。

## 可执行工具

Python 3.9+。审计、检查、基础解包/打包只需标准库；原生写作/修复及 Flat OPC 打包需要 `lxml`。仅缺失依赖时执行：

```powershell
python -m pip install -r ".\requirements.txt"
```

### 新文档写作

生成 UTF-8 JSON 内容规格，再写入新文件：

```powershell
python -X utf8 ".\scripts\docx_native.py" write "D:\work\content.json" "D:\work\报告.docx"
```

完整可运行规格见 `assets\writing-example.json`。支持富文本段落、真上下标、标题编号、题注、文献、前向引用、原生公式、目录和简单文字表格。示例文献是明确标注的占位内容，交付时必须换成真实来源。

这个轻量入口不是完整排版引擎。模板格式、图片、分节、复杂表格和其他正文结构，按 `references\writing.md` 使用通用 DOCX 工具或直接复用 `scripts\word_native.py` 的结构构造函数；不能为迁就入口删减用户内容。

### 文档修复

```powershell
python -X utf8 ".\scripts\audit_docx.py" "D:\work\原稿.docx" --output "D:\work\audit.json"
python -X utf8 ".\scripts\docx_native.py" repair "D:\work\原稿.docx" "D:\work\repair-plan.json" "D:\work\修复暂存.docx"
python -X utf8 ".\scripts\check_docx.py" "D:\work\修复暂存.docx" --source "D:\work\原稿.docx"
```

有模板则审计时增加 `--template "D:\work\模板.docx"`。修复计划示例见 `assets\repair-plan.example.json`，必须替换 SHA256、段落索引、原文字和目标规则，不能原样执行到其他文档。

修复入口只按明确计划编辑正文及表内的普通文本段落，不扫描猜测全部数学下标或引用对象。含已有域、超链接、图片、换行、修订等复杂内容的目标段落会拒绝处理，转为精确 XML/Word 范围编辑；不能把它们先展平成纯文本。

### 字段计算与检查

在用户允许 Word 自动化且 Windows 已安装桌面 Word 时：

```powershell
powershell -NoProfile -File ".\scripts\refresh_fields.ps1" -InputPath "D:\work\修复暂存.docx" -OutputPath "D:\work\修复完成.docx"
python -X utf8 ".\scripts\check_docx.py" "D:\work\修复完成.docx" --source "D:\work\原稿.docx"
```

刷新脚本只更新已有编号、引用及页码，**不会创建缺失引用，也不会重建目录条目**。新目录或条目变化需先在 Word 中定向“更新整个目录”，再检查；不要默认全局 `Fields.Update()` 或导出 PDF。没有 Word 时保留有效结构和已知缓存，如实标注页码、目录、字段重新计算未验证。

## 验收与交付

- 原稿未覆盖；输出是实际存在的新 DOCX，XML 可解析，字段闭合，书签成对且引用目标存在。
- 必要的标题、自动编号、原生上下标/OMML、题注、全文引用和目录都已建立；零个失效引用不等于缺失引用已补齐，须逐一对照原稿引用清单。
- 修复前后核对原图字节、表格数据、代码和正文。NFKC 比较只能辅助，不能证明没有内容变化。
- 逐条检查公式语义。只有实际在 Word 中计算/检查过，才能声称编号、目录和分页正确。
- 结构检查不是完整 XSD 校验，不是逐页视觉复核。需要 XSD 时使用已安装 `docx` Skill 的校验器及 `scripts\office_skill_bridge.py`，不复制其第三方代码。
- 不启动/关闭用户已有 Word，不按进程名批量结束进程。只管理自己创建且确认身份的实例；异常时保留 Flat OPC 状态。
- 对用户简洁说明交付路径、主要修复和未验证事项，不把历史案例的数量或验证结论当成本次结果。

## 参考资料边界

`references` 是日常执行规则；`assets` 是输入规格示例；`scripts` 是可执行工具。原工具中的历史案例仅保留在原工具目录，不属于本 Skill 的安装或公开发布内容；历史脚本包含专用路径和规则，禁止直接执行，案例审计正文不得上传。
