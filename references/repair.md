# 已有文档修复

## 先审计，后定向处理

保留原稿，输出新文件。`.doc` 需先转换为 `.docx`，转换副本也要检查内容；本入口不支持直接修改 `.doc`、加密文件或宏文档。用户未要求改变引用体例、标题层级、编号策略时保留原有规则；存在影响语义的歧义就询问，不凭正则猜测。

审计结果的 `REPORT.sha256` 用于计划绑定；`REPORT.paragraphs` 的 `index` 是 `word\document.xml` 内全部 `w:p` 的 **1 起始 XML 顺序**，包含表格段落，不等于 Word COM 的 Paragraphs 索引。`text` 是普通 Word 可见文字，OMML 另见 `math_text`。

审计当前输入及可选模板，检查样式、多级列表、真实上下标、题注编号源、书签、完整域指令、正文引用和目录。按用户模板的明确文字要求确定格式；模板内置样式与文字说明冲突时不能盲目全盘导入。

建立待修清单：哪些字符需要真实上下标，哪些图表/章节/文献被正文提及，以及各引用的唯一目标。检查失效域不能发现尚未建域的“图1”或“[1]”，必须从内容清单补齐缺失引用。没有正文引用的公式只建立编号与目标，不添加新的论断。

## 定向 JSON 计划

顶层必须有 `source_sha256`（来自本次审计）和 `operations`。每项必须有 `kind`、`paragraph`、`expect`；`expect` 与原稿该段完整普通文字逐字一致，防止误修其他版本。

| kind | 参数与行为 |
|---|---|
| `heading` | `level`；可选 `prefix` 删除手写标题编号、`bookmark`、`numbered`；绑定真正标题和自动列表 |
| `caption` | `prefix`（开头原手写标签/编号）、`label`、`cached`、`bookmark`；可选 `sequence`、`chapter`；将前缀替换成 SEQ 与编号书签 |
| `bibliography` | 可选原手写 `prefix`、`bookmark`；同一计划内的文献共用独立自动编号列表 |
| `script` | `match`、`script`；可选替换后的普通字符 `text`、1 起始 `occurrence`；仅修改明确片段 |
| `replace` | `match`、`parts`；可选 `occurrence`；通常用于把静态引用替换成 REF |
| `equation` | 明确的 `math`、`cached`、`bookmark`；可选 `sequence`、`chapter`；将纯文本公式段替换为已确认的 OMML 和编号 |

示例见 `assets\repair-plan.example.json`。它假定一个明确的四段输入，不是通用自动识别规则。必须按当前审计修改。

同一段可以有多个不重叠的字符/引用操作，坐标都基于**未修改的原稿**；入口先定位完整片段，再倒序处理，避免先改 `²` 后找不到整个 `⁻²`。匹配出现多次时必须给 `occurrence`，否则拒绝。公式整段替换不能与其他文本操作混用。

标题和文献修复会新建编号列表，计划必须覆盖需要共享该编号的所有目标，按实际文档顺序从 1 开始。只迁移列表中的一部分可能产生混用或重新起号；要保留既有列表、起始编号或复杂附录，使用其已有 numId/Word 范围进行精确编辑，不套这个简化入口。

计划只处理正文及表格内的**普通文本段落**；包含域、超链接、书签、图片、制表符、换行、修订等结构时拒绝文本替换。不能为“通过”检查把段落内容展平、删除对象或接受用户修订。正文外的页眉页脚、脚注/尾注采用精确 XML 或 Word story 范围处理，再用全域检查发现失效引用。

入口保留所有未修改 ZIP 部件的原始字节，维护新增编号的包关系和内容类型，禁用已有 `updateFields` 的打开时全局更新；输出拒绝覆盖源文件或已有目标。数字签名会因编辑失效，签名文档需单独确定签名处理方式。

## 超出简化入口时

```powershell
python -X utf8 ".\scripts\docx_package.py" unpack "D:\work\原稿.docx" "D:\work\unpacked"
# 精确编辑必要 XML；保持其他内容和命名空间不变。
python -X utf8 ".\scripts\docx_package.py" pack "D:\work\unpacked" "D:\work\staging.docx"
```

基础打包器只检查 XML 可解析。可选完整校验：

```powershell
python -X utf8 ".\scripts\office_skill_bridge.py" --office-dir "$env:USERPROFILE\.agents\skills\docx\scripts\office" validate "D:\work\staging.docx"
```

该路径只是常见安装示例；先确认本机实际路径。桥接不携带第三方工具，缺失就说明，不能声称 XSD 已通过。`--original` 会启用修订内容保护检查，普通未跟踪的格式修复不能随意把其失败忽略。

精确 XML/COM 编辑时保留原 run 的字体、字号、语言和其他属性，按 schema 顺序插入 `vertAlign`、`numPr`、`outlineLvl` 等。Word Find 返回的实际范围才可用于含隐藏域代码的段落；普通字符串偏移不等于 Word Range 位置。

科学下标采取明确白名单或逐处确认，保留 FP32、W8A8、代码名称、数组索引和真实音标。原生公式由用户内容逐条映射，不能从历史案例直接复制公式、物理参数或数量断言。

## 刷新和交付边界

`refresh_fields.ps1` 在自建 Word 实例中只读打开暂存 DOCX，先更新 STYLEREF/SEQ，再 REF，再分页、PAGEREF 和页脚，通过 Flat OPC 导出新文档。它不会修复缺失字段，主正文之外的引用更新需按需要定向处理；页脚原有更新策略不能证明所有 story 均已计算。

新目录、改名或增删的目录条目需定向更新整个目录；仅更新 PAGEREF 不会重建条目。不要无条件全局更新所有域、调用保存/导出 PDF 或关闭用户原有 Word。长时间阻塞先保留状态，只处理当前任务自己创建且确认身份的实例。

最终运行 `check_docx.py --source 原稿.docx`，它检查正文、页眉页脚、脚注/尾注的字段和引用目标、特殊脚本字符候选、书签边界；原图字节和表格文字比较仍基于正文包。候选脚本字符可能是合法音标，检查失败需要解释并逐处判断，不能为了“绿灯”删除合法字符。

另外按待修清单核对缺失引用确已补齐；逐条检查原生公式；对照原稿的代码、正文和数字。NFKC 会把某些兼容字符归一化，不能作为内容保持的唯一证据。只有实际检查过的范围才能写成通过，未做 Word 重算、分页或视觉检查要明确说明。
