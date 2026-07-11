# 下一步行动 NEXT_ACTIONS

> 生成日期：2026-07-06
> 目标：列出**可立即开工**的具体任务，具体到文件、方法与验收方式。
> 本次任务仅产出文档，以下均为“下一阶段编码建议”，尚未实施。

## A. 立即可做（无需求阻塞）

### N1. 纳入 Git 版本管理【必须做，先做】
- 动作：在仓库根 `git init`；新增 `.gitignore` 忽略 `bin/`、`obj/`、`.vs/`、`packages/`、`*.user`。
- 注意：**评估 `iWorkhelper_TemporaryKey.pfx` 是否提交**——含私钥，建议忽略并单独保管（发布用正式证书）。
- 验收：`git status` 干净，构建产物不入库，首次提交完成。

### N2. 搭建日志基础设施【必须做】
- 新增 `Core/Infrastructure/Logger.vb`：提供 `Info/Warn/Error(msg, ex)`，写入 `%AppData%\iWorkHelper\logs\yyyy-MM-dd.log`。
- 验收：调用后日志文件生成，含时间戳、级别、消息、异常堆栈。

### N3. 打通“读取当前邮件 PDF 附件”【必须做，MVP 第一环】
- 新增 `Core/MailAttachmentReader.vb`。
- 逻辑：通过 `Globals.ThisAddIn.Application.ActiveExplorer().Selection` 取选中项 → 转 `Outlook.MailItem` → 遍历 `Attachments` → 筛 `.pdf` → `SaveAsFile` 到临时目录。
- 边界处理：无选中/非邮件/无附件/非 PDF → 通过 N2 日志 + 用户提示。
- 接入点：先在 `MainRibbon.vb:9 ButtonArchive_Click` 里替换 `MsgBox` 占位，改为调用并弹出“找到 N 个 PDF 附件：…”。
- 验收：选中一封含 PDF 附件的真实邮件点击“归档”，弹出正确的附件文件名与临时路径。

### N4. PdfPig 抽取文本自测【必须做，验证依赖可用性】
- 新增 `Core/Local OCR/PdfPigTextParser.vb`：`Function ExtractText(pdfPath As String) As String`，用 `UglyToad.PdfPig.PdfDocument.Open` 遍历页面 `page.Text`。
- **同时验证**：项目引用的是 PdfPig 的 `net471` DLL，需确认在 net48 工程中正常加载运行（见风险 R-TECH-1）。
- 验收：对一份文本型 PDF 返回非空文本；对图片型/加密 PDF 返回空并记录明确原因（这将直接回答 Q1）。

## B. 需先确认才能做（阻塞项，建议同时推进确认）

### N5. 确认关键需求（发给需求方，越早越好）
需要对方回答（详见 [../requirements/REQUIREMENTS_BACKLOG.md](../requirements/REQUIREMENTS_BACKLOG.md) 的 Q1~Q10）：
1. **Q1** 滴滴发票 PDF 是文本型还是扫描图片型？（决定是否要真 OCR）
2. **Q2** 在线模式对接哪个 OCR/发票识别服务？有无账号/Key/预算？
3. **Q3** 需要提取哪些字段？
4. **Q4** 文件命名规范？
5. **Q5** 是否分类归档？规则？
6. **Q6** 处理当前选中单封，还是批量？
7. **Q8** 目标框架以 4.8 为准还是文档所写 8.0？
- 并请对方提供 **2~3 份脱敏样例 PDF**（否则 N6/N7 无法验证）。

### N6. 字段提取器（依赖 Q3 + 样例）【待确认后做】
- 新增 `Core/Extraction/DidiInvoiceExtractor.vb`：输入文本，输出字段对象（发票号/金额/日期/行程等，按 Q3）。
- 验收：对样例 PDF 提取字段与人工核对一致。

### N7. 命名 + 归档（依赖 Q4/Q5）【待确认后做】
- 新增 `Core/Naming/FileNamingService.vb`（生成合法文件名、冲突加序号）与 `Core/Archiving/ArchiveService.vb`（落盘到 `My.Settings.ArchiveFolderPath`，目录缺失自动创建）。
- 验收：文件按规范命名并出现在归档目录；重复执行不覆盖正确文件。

### N8. 串联端到端并替换归档占位【MVP 收口】
- 修改 `MainRibbon.vb:9 ButtonArchive_Click`：`读附件(N3) → 抽文本(N4) → 提字段(N6) → 命名(N7) → 归档(N7)`，全过程日志化，结束弹出结果汇总。
- 验收：对真实滴滴邮件一键归档成功，得到规范命名的 PDF 落到归档目录。

## C. 建议的最小编码顺序

```
N1 → N2 → N3 → N4   （立即可做，验证 Outlook 取件 + PdfPig 可用性）
        并行推进 N5（需求确认 + 样例）
N5 关闭 Q1/Q3/Q4 后 → N6 → N7 → N8（MVP 完成）
```

## D. 本阶段“不要做”

- 不要在需求（Q1~Q4）确认前动手写字段提取/命名规则——会返工。
- 不要提前引入 OCR 引擎或云服务 SDK——先用 PdfPig 验证是否够用。
- 不要重构现有 UI/设置代码（当前可用，无需改）。
