# 批量归档流程 PROCESS_FLOW

> 日期：2026-07-06
> 对应实现：`Core/Workflow/BatchArchiveWorkflow.vb` 及其调用的各模块。

## 1. 端到端流程图

```
用户在 Outlook 选中一封或多封邮件
        │
        ▼
点击 功能区「工作助手 > 发票归档 > 归档」
   (MainRibbon.ButtonArchive_Click)
        │
        ▼
[运行锁] ArchiveRunGuard.TryAcquire（Interlocked 原子获取，单一状态来源，防重复点击/并发）
        ├ 失败(已有任务真正在运行) → 提示“已有归档任务正在运行”，返回（不进入预检查）
        └ 成功(token) → 进入 Using token（保证任何路径都释放运行锁）
        │
        ▼
[预检查] ArchivePreflightChecker：选中邮件/归档目录(配置/存在/写权限)/临时/日志目录/命名模板/OCR 配置
        ⚠ 预检查【不再】检查“是否已有任务运行”——该职责已归 ArchiveRunGuard，避免自我阻断
          （历史缺陷：先置运行标志、预检查又检查该标志，导致每次点击都误报“已有任务运行”）
        ├ 有阻断项 → 一次性列出问题 + 日志位置，不建进度窗口，Using 释放运行锁
        └ 通过 → 打开进度窗口
        │
        ▼
BatchArchiveWorkflow.Run(application, reporter)  （批次 ID = B+时间戳）
        │
        ├─[1] 读取设置：ArchiveFolderPath、ParseMode(Local/Online)
        │        └ allowOnline = (ParseMode == "Online")
        │
        ├─[2] AppLogger.Initialize(归档目录)  → 日志写入 归档目录\logs 或 %AppData%\iWorkHelper\logs
        │
        ├─[3] ArchivePlanner.ValidateArchiveFolder
        │        ├ 目录为空 → 返回 ConfigurationMissing（提示先设置）——流程结束
        │        └ 目录不存在 → 尝试创建；失败 → Failure，流程结束
        │
        ├─[4] MailAttachmentReader.ReadSelectedPdfAttachments
        │        ├ 遍历 Explorer.Selection（1..Count）
        │        ├ 非 MailItem → 跳过并记日志（SkippedNonMailCount++）
        │        ├ MailItem → 遍历 Attachments，仅 .pdf
        │        ├ SaveAsFile 到临时目录（PathHelper.GetTempWorkDirectory）
        │        │     文件名经 GetNonConflictingPath 处理，不覆盖
        │        ├ 无 PDF 的邮件 → MailsWithoutPdfCount++，给出可读消息
        │        └ 释放所有 COM 对象（Marshal.ReleaseComObject）
        │        └ 无任何 PDF → 返回 Skipped，流程结束
        │
        ▼
   对每个 PDF 附件循环（单项失败不中断整批）：
        │
        ├─[5] RecognitionPipeline.Recognize(pdfPath, 原文件名)
        │        ├ (a) PdfTextExtractor.Extract → 页数/文本/是否疑似图片型 + 坐标行重建
        │        ├ (b) LocalTextInvoiceRecognizer.Recognize(文本)
        │        │        └ 非滴滴 VatInvoice → 委派 GeneralInvoiceLocalRecognizer（候选评分/分区/明细）
        │        │        ├ 文本为空 → NeedsOcr
        │        │        ├ 解析到关键字段且较完整 → Success
        │        │        ├ 解析到部分字段 → PartialSuccess
        │        │        └ 有文本但无关键字段 → NeedsOcr
        │        ├ (c) 若本地可用(Success/Partial) → 直接返回
        │        └ (d) 若 NeedsOcr：
        │                 ├ allowOnline 且 百度OCR配置齐全 → BaiduOcrInvoiceRecognizer.Recognize
        │                 │      （当前真实 API 为 TODO，返回 Failure/占位）
        │                 └ 否则 → NeedsOcr（附说明：未配置/本地模式）
        │
        ├─[6] ArchiveNamingRule.BuildFileName(invoice, 原名, 时间戳)
        │        ├ 有字段 → 开票日期_销售方_发票号_价税合计.pdf
        │        └ 无字段 → 回退：原名 或 未识别票据_时间戳（不中断）
        │        └ 全程 FileNameSanitizer 清理非法字符
        │
        ├─[7] ArchivePlanner.PlanTargetPath → GetNonConflictingPath（同名自动追加(1)(2)…）
        │
        ├─[8] ArchiveExecutor.Execute(临时文件, 目标路径)
        │        ├ 目标已存在 → 拒绝覆盖，Failure
        │        ├ File.Copy(overwrite:=False)
        │        └ 成功 → 删除临时文件；失败 → 保留临时文件
        │
        └─[9] 记录 ArchiveItemResult（状态/目标名/消息），写日志
        │
        ▼
   汇总 ArchiveBatchResult（成功/部分/需OCR/跳过/失败计数）
        │
        ▼
MainRibbon 展示一次性汇总 MessageBox（不逐附件弹窗）
   详细结果已在日志文件中
        │
        ▼
End Using / Finally → ArchiveRunToken.Dispose → ArchiveRunGuard 释放运行锁
   （预检查失败早退、业务异常、进度窗口异常关闭等任何路径都会释放）
```

## 2. 状态语义

| 单项状态 (ProcessStatus) | 含义 | 是否已落盘 |
|--------------------------|------|-----------|
| Success | 识别完整并归档 | 是 |
| PartialSuccess | 识别不完整但已归档（含识别异常但文件保住） | 是 |
| NeedsOcr | 本地无法识别，且在线 OCR 不可用/未配置 | 是（回退命名） |
| Failure | 归档复制失败（如目标已存在/无权限） | 否 |
| Skipped | 无 PDF 附件等（批次级） | — |

> 设计取向：**尽量不丢文件**。即便识别失败/需 OCR，PDF 仍以回退名归档，用户可后续人工处理；仅当"复制动作本身"失败才记 Failure。

## 2.3 分流处理（滴滴/常规/未识别/无 PDF）

预检查通过后，读取分组 → **对每封邮件的每个 PDF 识别并分类**（`MailProcessingClassifier`）→ 得出邮件类型 → 路由：

- **滴滴发票邮件** → 现有合并链路（发票+行程单合并→统一命名归档，1 个归档项）。
- **常规发票邮件 / 混合邮件** → 逐 PDF：常规发票用常规模板单独归档；未识别 PDF 用 `未识别_{原始文件名}` 单独归档（各产出 1 个归档项）。
- **仅未识别邮件** → 每个 PDF 按未识别命名归档。
- **无 PDF 邮件** → 跳过（分组阶段已过滤）。

进度以邮件为主维度；阶段新增“正在判断邮件类型/正在处理常规发票/正在处理未识别 PDF/当前邮件无 PDF 已跳过”。单封邮件/单个 PDF 失败不影响其它。详见 [../requirements/MAIL_PDF_CLASSIFICATION_SPEC.md](../requirements/MAIL_PDF_CLASSIFICATION_SPEC.md)。

## 2.2 第七阶段追加：每封邮件合并 PDF + 进度条（归档单位=邮件）

**重要变更：归档单位由“附件”改为“邮件”——每封邮件生成一份合并 PDF 再归档。**

```
读取选中邮件 → 按邮件分组（MailPdfGroupingResult，每封一组，含其全部 PDF）
   对每封邮件（进度以“邮件”为单位）：
     [ExtractingText/CallingOcr] 逐个 PDF 识别（RecognitionPipeline，本地优先/在线兜底）
          · 本地解析：PdfPig 文本 + **坐标行重建按列取行程单单元格**（起点/终点/金额，对地址长短/换行稳）
          · 本地缺命名核心字段(乘车日期/金额/出发地点/到达地点)→ 视为不足 → 回退在线 OCR
     [ParsingFields] 合并识别结果（InvoiceRecognitionMerger）：
          - 发票字段以“含发票号码”的结果为准；行程明细做并集；扩展字段并集
          - 来源综合：全本地=LocalText / 含在线=Mixed / 全在线=BaiduOcr / 全失败=Failed
          - 结果绑定到本封邮件的合并 PDF
     [Merging] 合并 PDF（PdfPig 自带 PdfMerger）：发票在前、行程单在后、未知最后；
          无法分类则按原始附件顺序；输出到临时目录、不覆盖原件、不冲突命名
     [Naming] 命名（**单一统一模板** {乘车日期}_{金额}_{出发地点}_{到达地点}；
              取值：乘车日期=行程出发日期>开票日期，金额=行程金额>价税合计>金额；
              字段严重不足→fallback 未识别票据_{邮件主题}_{时间戳}）
     [Archiving] 归档合并 PDF（同名追加序号，不覆盖）
     [WritingResult] 记录 ArchiveItemResult（邮件主题/PDF数/原始PDF/合并临时/最终路径/来源）
     成功后清理临时文件（原始 PDF + 合并 PDF）；失败保留供排查
   单封邮件失败不影响其它邮件
汇总（以“邮件”为单位）→ 关闭进度窗口 → 一次性汇总弹窗
```

**OCR 时机**：采用“**合并前分别识别、合并识别结果**”策略（对每个 PDF 分别本地/OCR，再用 `InvoiceRecognitionMerger` 合并为一个结果），结果绑定到合并 PDF。理由：各 PDF 在其原始上下文中识别质量更好，且无需依赖多页 OCR。

**PDF 合并库**：使用**已引用的 PdfPig 自带 `UglyToad.PdfPig.Writer.PdfMerger`**（`Byte() Merge(String())`），**无需引入新库**；合并后 PDF 可正常打开（实测 2×2 页→4 页）。

**进度条阶段**（`ArchiveStage`）：正在读取邮件附件 / 保存 PDF / 合并发票和行程单 / 提取 PDF 文本 / 调用百度 OCR / 解析票据字段 / 生成归档文件名 / 归档文件 / 写入处理结果 / 已完成。进度窗口显示总邮件数、已处理数、当前邮件序号+主题、当前阶段、百分比。VSTO 归档在 UI 线程执行，采用 `Application.DoEvents` 在关键阶段刷新（避免后台线程访问 Outlook COM）。**无取消按钮**（关闭窗口不取消任务）——取消成本较高，本阶段不做。

## 2.4 常规发票本地识别专项（2026-07-10）

`[5]` 识别环节：`LocalTextInvoiceRecognizer` 检测到**非滴滴的 VatInvoice** 时委派 `GeneralInvoiceLocalRecognizer`：

```
GeneralInvoiceLocalRecognizer.Recognize
  1) 归一化 + 逻辑行（坐标行重建 / 无坐标按换行）
  2) 发票类型识别（专票/普票/数电票/电子发票…）
  3) 字段候选 + 评分择优：发票号码/代码、开票日期、购/销名称、价税合计(小写)、税率、税额
  4) 商品明细解析（PdfTableRegionDetector + GeneralInvoiceLineItemParser，失败不影响头部）
  5) 状态：开票日期+金额+销售方齐全→Success；缺→PartialSuccess（回退 OCR）；无核心→NeedsOcr
```

`{金额}` 对常规发票取值优先级：**价税合计 > 金额**（不取明细行金额）。本地缺 开票日期/金额/销售方名称 任一 → 字段不足 → 启用且配置完整时回退百度 OCR（`RecognitionPipeline`）。滴滴/行程单识别与合并流程、未识别 `未识别_{原始文件名}` 均不变。详见 [../development/GENERAL_INVOICE_RECOGNITION.md](../development/GENERAL_INVOICE_RECOGNITION.md)。

## 2.1 第四阶段流程增量

- **命名**：`[6]` 命名环节改为**模板驱动**（`NamingTemplateEngine`），从设置读取发票/行程单/未识别三类模板；空占位符跳过并折叠分隔符；未知占位符提示；模板异常回退内置默认规则。占位符清单见 [../requirements/NAMING_TEMPLATE_SPEC.md](../requirements/NAMING_TEMPLATE_SPEC.md)。
- **多条行程**：`[5]` 识别环节对行程单产出多条 `InvoiceTripInfo`；命名默认取首条；归档结果记录 `TripCount`。
- **Secret Key**：`[1]` 读取设置时，SK 经 DPAPI 解密（`ProtectedSettingsProvider`）；启动时明文自动迁移为密文。
- **汇总**：`ArchiveItemResult` 增加 `RecognitionSource`、`NamingRule`、`MissingFields`、`TripCount`，日志与（后续）UI 可展示。

> 补充说明：`[5]` 在线 OCR 已用真实 AK/SK 联调成功（Mixed，31 字段），行程明细可从 OCR 的 `Passeng*` 字段提取；`[6]` 命名支持设置页"变量说明/预览"；`[1]` 配置来源支持加密配置文件（`BaiduXmlConfigStore`，SK DPAPI 加密），My.Settings 未填时读取。演进详情见 [../history/DEVELOPMENT_HISTORY.md](../history/DEVELOPMENT_HISTORY.md)，Outlook 端到端验证见 [../testing/OUTLOOK_MANUAL_TEST.md](../testing/OUTLOOK_MANUAL_TEST.md)。

## 3. 关键目录

| 用途 | 路径 |
|------|------|
| 临时工作目录 | `%AppData%\iWorkHelper\temp` |
| 日志目录 | `{归档目录}\logs`（可用时）否则 `%AppData%\iWorkHelper\logs` |
| 百度 OCR 配置 | `%AppData%\iWorkHelper\baidu-ocr.config.xml` |
| 归档目标 | `My.Settings.ArchiveFolderPath`（用户在设置中选择） |

## 4. 失败隔离与健壮性

- 每个附件在独立 `Try/Catch` 中处理，异常仅影响该项。
- 每个 COM 访问点都有安全包装（`SafeSubject` 等）并在 `Finally` 释放。
- 日志任何异常被吞掉，绝不影响归档主流程。
- 归档目录缺失是**显式错误**（ConfigurationMissing），不静默。
