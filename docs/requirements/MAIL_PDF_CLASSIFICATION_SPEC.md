# 邮件/PDF 分流处理规范 MAIL_PDF_CLASSIFICATION_SPEC

> 日期：2026-07-07
> 归档按邮件/PDF 类型分流：滴滴发票合并、常规发票单独、未识别按固定命名、无 PDF 跳过。

## 1. 分类枚举

**邮件级 `MailProcessingType`**：DidiInvoiceMail / GeneralInvoiceMail / MixedPdfMail / UnknownPdfOnlyMail / NoPdfMail
**PDF 级 `PdfAttachmentClassification`**：DidiInvoicePdf / DidiTripPdf / GeneralInvoicePdf / UnknownPdf

## 2. 分类依据与优先级（`MailProcessingClassifier`）

不只看邮件主题，优先级：
1. **已识别字段/票据类型**：`recog.Invoice`（发票号码、行程 Trips）、`DocumentType`；
2. **PDF 文本关键词**；
3. **OCR 结构化结果**（体现在 recog 字段）；
4. **邮件主题**；
5. **附件文件名**。

### PDF 判定
- **滴滴特征**：Trips>0 / 销售方含“滴滴” / DocumentType=行程单 / 文本含 滴滴出行科技有限公司·滴滴·行程单·网约车·上车时间·订单号·TRIP TABLE·笔行程 / 主题含 滴滴·行程单·出行发票·网约车 / 文件名含 滴滴·行程单。
  - 有滴滴特征 + 有发票号码 → **DidiInvoicePdf**
  - 有滴滴特征 + 行程单文本、无发票号 → **DidiTripPdf**
  - 有滴滴特征但含糊 → DidiInvoicePdf
- **非滴滴**：有发票号码 或 文本含 发票代码·发票号码·开票日期·购买方·销售方·纳税人识别号·价税合计·校验码·增值税电子普通发票·数电·电子发票 → **GeneralInvoicePdf**
- 其余 → **UnknownPdf**

> 滴滴发票本身也是发票，但同时具备滴滴/行程特征时**优先归滴滴**；普通发票无行程特征归常规；无法判断归 UnknownPdf（不强行归常规）。分类不确定记日志。

> **识别与分类的关系（2026-07-10）**：分类不变；但**常规发票（非滴滴）的本地识别已改由 `GeneralInvoiceLocalRecognizer` 处理**（`LocalTextInvoiceRecognizer` 在 `DocumentType=VatInvoice` 时委派）。这提升了 `recog.Invoice` 字段质量（购/销不混淆、价税合计不误取明细金额、发票号码不误取），从而让「已识别字段」这一最高优先级依据更可靠——滴滴优先、常规归常规、特征不足归 UnknownPdf 的判定更准。滴滴发票/行程单识别与分类**完全不变**。

### 邮件判定
- 含任一 Didi* → **DidiInvoiceMail**
- 否则：有常规 + 有未知 → **MixedPdfMail**；仅常规 → **GeneralInvoiceMail**；仅未知 → **UnknownPdfOnlyMail**；无 → **NoPdfMail**

## 3. 各类处理规则

### 滴滴发票邮件（不变）
以邮件为单位：发票 PDF + 行程单 PDF 合并为一个 PDF（发票在前、行程单在后），合并识别结果绑定合并 PDF，按**滴滴命名模板**归档。默认 `{乘车日期}_{金额}_{出发地点}_{到达地点}`。

### 常规发票邮件
以 **PDF 为单位**：每个常规发票 PDF 单独识别、命名、归档（不合并）。一封多张则分别归档。按**常规发票命名模板**。默认 `{开票日期}_{金额}_{销售方名称}`；`{金额}` = 价税合计 > 金额。字段不足用 fallback 名。单个失败不影响其它 PDF/邮件。

### 未识别 PDF
每个按 **`未识别_{原始文件名}.pdf`** 归档（`UnknownPdfNamingRule`）；清理 Windows 非法字符；同名追加序号、不覆盖；归档结果标记“未识别PDF”。不做复杂识别也可归档。

### 混合邮件（常规 + 未知）
常规按常规、未知按未识别，分别产出归档项；批次结果分别记录。

### 无 PDF 邮件
直接跳过：不报错、计入“跳过”、日志“无 PDF 附件，已跳过”、进度继续、不影响后续。（分组阶段已过滤无 PDF 邮件，计入 `MailsWithoutPdfCount`。）

## 4. 命名模板变量（常规发票可用）

`{开票日期} {金额} {价税合计} {销售方名称} {购买方名称} {发票号码} {发票代码} {税额} {税率} {原始文件名} {邮件主题} {时间戳} {识别来源} {票据类型}`

- `{金额}` 常规发票 = 价税合计 > 金额（无行程金额）。
- `{原始文件名}` = 当前 PDF 附件原始名（不含扩展名、已清理非法字符）。
- 未识别 PDF 的 `未识别_{原始文件名}` 为内部固定规则，不作为设置项。
- 变量单一来源 `ArchiveNamingRule.SupportedPlaceholders()`，与设置页“变量说明”一致。

## 5. 归档结果与汇总

- `ArchiveItemResult.ProcessingKind`：滴滴合并 / 常规发票 / 未识别PDF / 跳过 / 失败。
- `ArchiveBatchResult` 统计：MailCount（去重邮件数）、DidiCount、GeneralInvoiceCount、UnknownPdfCount，及成功/部分/需OCR/跳过/失败/合并数。
- 汇总弹窗：`共 N 封邮件（滴滴 x，常规发票 y，未识别PDF z）：成功…失败…` + 报告/日志路径。
- 报告（`archive-report-*.txt`）每项含：邮件主题、分类(ProcessingKind)、PDF 数、原始附件、合并临时、归档路径、命名规则、fallback、失败原因。

## 6. 进度阶段（新增）

正在读取邮件附件 / 正在判断邮件类型 / 正在识别滴滴发票与行程单 / 正在合并发票和行程单 / 正在处理常规发票 / 正在处理未识别 PDF / 正在生成归档文件名 / 正在归档文件 / 当前邮件无 PDF 已跳过 / 正在写入处理结果 / 已完成。进度以**邮件**为主维度；多 PDF 时逐个处理；不逐 PDF 弹窗。

## 7. 已验证 / 待验证

- **真实通过（离线 `--selftest` 54/54 + `--classify`）**：滴滴/常规/未识别 PDF 分类、邮件类型判定、常规发票默认命名、未识别命名、滴滴样例仍走合并流程。
- **待 Outlook 验证**：真实邮件分流全流程（常规多 PDF 分别归档、混合邮件、无 PDF 跳过、进度阶段显示、汇总统计）。
- 未提供真实常规发票/未识别 PDF 样例，分类以关键词/字段规则为准，真实样例可用 `--classify` 复核。
