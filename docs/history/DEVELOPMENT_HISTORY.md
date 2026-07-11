# 开发历程 DEVELOPMENT_HISTORY

> 本文由各阶段实施报告（PHASE_2~11）整合而成，只保留阶段目标、关键功能、重要设计决策、关键修复、验证结果与遗留问题；逐文件流水账已删除。当前实现请以源码与正式文档为准。
>
> 说明：项目起步于 2026-07-06 的接手分析（当时仅 UI/设置骨架、Core 为空），随后按阶段演进至完整归档链路。以下"未验证"均指无 Outlook 端到端环境所致，离线部分已真实执行。

---

## 阶段 2：最小可用链路（MVP）

- **目标**：打通 选中邮件 → 读 PDF 附件 → 抽文本 → 识别 → 命名 → 归档 的最小闭环。
- **关键功能**：基础设施（日志/统一结果/文件名清理/不冲突命名/异常格式化）；批量附件读取（COM 释放）；PdfPig 文本抽取 + 疑似图片型判定；识别抽象层 + 本地文本识别器（初版正则）；命名 + 归档 + 批量汇总（单项失败隔离）；Ribbon 串联替换占位。
- **设计决策**：`Core/` 采用无空格分层目录；本地优先调度；尽量不丢文件（识别失败也以回退名归档）。
- **验证**：主工程 Debug 构建 EXIT 0；PdfPig net471 DLL 在 net48 工程构建通过。
- **遗留**：百度 OCR 真实 API 当时为 TODO；本地正则待样例校准。

## 阶段 3：百度 OCR 真实接入 + 配置界面

- **目标**：接入百度 multiple_invoice、SettingsForm 加 OCR 配置、样例驱动字段校准。
- **关键功能**：`Core/Ocr/Baidu/*` 全套（token 进程内缓存+提前 5 分钟刷新、TLS 1.2、字段集中映射、未知字段入 ExtendedFields）；`KeyFieldEvaluator` 关键字段评估；SettingsForm 新增 12 项 OCR 配置（SK 密码遮罩）；OfflineTester 建立。
- **设计决策**：不引第三方 HTTP/JSON 库（`HttpWebRequest`+`JavaScriptSerializer`）；不硬编码密钥（AK/SK 来自设置或环境变量，日志仅脱敏摘要）。
- **验证**：5 份真实滴滴样例本地解析全部 Success/LocalText；主+离线工程构建 EXIT 0。修复本地税率误取相邻金额尾数（`134.433%`→`3%`）。
- **遗留**：真实 AK/SK 联调；SK 明文存储待加密。

## 阶段 4：命名模板 + 多行程 + DPAPI

- **关键功能**：命名模板引擎（占位符渲染/空字段折叠/未知占位符提示/异常回退）；本地多条行程解析（以上车时间为锚点）；**Secret Key DPAPI 加密**（`SecretProtector`/`ProtectedSettingsProvider` + 启动时明文迁移）；OfflineTester 增强（`--ocr`/`--force-ocr`/`--save-response`）。
- **验证**：5 份样例本地回归无回归；DPAPI 往返离线验证。`.gitignore` 建立。
- **遗留**：真实 OCR 联调；DPAPI 在 Outlook 会话内实测。

## 阶段 5：验证固化

- **验证**：OfflineTester `--selftest` 建立（DPAPI 往返 / 命名模板边界 / 合成 OCR JSON 解析映射）；命令路径与退出码；本地样例回归；双工程 Rebuild EXIT 0。诚实标注真实 OCR 与 Outlook 端到端未验证，未伪造通过。

## 阶段 6：真实百度 OCR 联调（成功）

- **关键成果**：用真实 AK/SK 对 5 份样例联调成功（`Success/Mixed/VatInvoice/31 字段`）。
- **关键修复（真实 bug）**：`pdf_file` 编码原用 `Uri.EscapeDataString`（.NET Framework 约 65520 字符上限）对超长 Base64 抛异常 → 改 `System.Net.WebUtility.UrlEncode`。详见 [../development/OCR_ONLINE_VALIDATION_LOG.md](../development/OCR_ONLINE_VALIDATION_LOG.md)。
- **字段校准**：复核人真实字段名为 `Checker`（原误用 `Reviewer`）；滴滴行程信息内嵌于 `vat_invoice.result` 的 `Passeng*` 字段（按 `row` 分组）。
- **功能**：设置页"变量说明/预览"按钮（单一来源 `SupportedPlaceholders`）。

## 阶段 7：配置安全落地 + 邮件合并归档 + 进度条

- **配置落地**：新增加密配置文件 `%AppData%\iWorkHelper\baidu-ocr.config.xml`（SK DPAPI 加密）；OfflineTester `--save-baidu-config`/`--use-config`；实测"加密落地→解密→真实 OCR"闭环。
- **邮件合并归档（重要变更）**：归档单位由"附件"改为"邮件"——每封邮件的发票 PDF + 行程单 PDF **合并为一份 PDF** 再识别/命名/归档。合并用 **PdfPig 自带 `PdfMerger`（无新库）**；采用"合并前分别识别、`InvoiceRecognitionMerger` 合并结果"策略。
- **进度条**：`ProgressForm`（10 阶段、`Application.DoEvents` 刷新、无取消按钮）。
- **发布前安全扫描**：仓库无明文 AK/SK/token/敏感文件。

## 阶段 8：合并为单一命名规则

- 命名配置**合并为一套统一模板**，默认 `{乘车日期}_{金额}_{出发地点}_{到达地点}`；旧三套（发票/行程/未识别）保留仅兼容、业务不再使用。业务变量优先级：乘车日期=行程出发日期>开票日期；金额=行程金额>价税合计>金额。fallback：`未识别票据_{邮件主题}_{时间戳}`。详见 [../requirements/NAMING_TEMPLATE_SPEC.md](../requirements/NAMING_TEMPLATE_SPEC.md)。

## 本地识别稳定性专项：坐标行重建

- **根因**：旧实现用 `page.Text` 全文拼接 + 强假设正则切分起点/终点，对"终点无 `区|`"、长地址换行、车型别名不稳（真实反馈 6 封中仅 3 封正确）。
- **修复**：改用 **PdfPig 词坐标按列取行程单单元格**（`PdfTextLayoutExtractor` + 列区间过滤 + 换行拼接 + 文本归一化）；命名核心字段不足时回退在线 OCR。5 份样例命名核心字段 **5/5 命中**。详见 [../development/LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT.md](../development/LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT.md)。

## 邮件/PDF 分流处理

- 归档按类型分流：滴滴发票（合并归档）/ 常规发票（每张单独归档）/ 未识别 PDF（`未识别_{原始文件名}`）/ 无 PDF 邮件（跳过）。分类器 `MailProcessingClassifier` 依据 已识别字段 > PDF 文本 > OCR > 主题 > 文件名，滴滴特征优先。详见 [../requirements/MAIL_PDF_CLASSIFICATION_SPEC.md](../requirements/MAIL_PDF_CLASSIFICATION_SPEC.md)。

## 阶段 9：容错处理与体验优化

- 统一错误分类（`AppErrorCode`/`AppError`/`UserFriendlyMessageProvider`，中文友好提示，日志不泄漏密钥）；归档前预检查（`ArchivePreflightChecker`）；设置页校验增强 + "测试 OCR 配置"按钮（仅连通性）；批量失败隔离；批次报告 `archive-report-*.txt`。详见 [../architecture/ERROR_HANDLING_DESIGN.md](../architecture/ERROR_HANDLING_DESIGN.md)。
- > 注：本阶段的预检查"防重入 TryBeginRun/EndRun"设计后被发现有自我阻断缺陷，已在下述缺陷修复中重构。

## 缺陷修复：点击归档立即提示"已有归档任务正在运行"（2026-07-10）

- **根因**：预检查前先 `TryBeginRun()` 置运行标志，预检查又检查该标志 → 每次点击都自我阻断（源码逻辑顺序错误，清缓存/正式部署/换调试方式均无效）。
- **修复**：新增单一来源运行锁 `ArchiveRunGuard`（`Interlocked` 原子获取、`Using` 释放、纯内存进程级）；预检查不再检查运行状态；`MainRibbon` 重排为"先获取运行锁→预检查→进度→归档→释放"。详见 [../development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md](../development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md)。

## 常规发票本地识别专项（2026-07-10）

- **根因**：常规发票原与滴滴共用一套滴滴校准正则，无坐标/分区/候选评分，购销易混淆、明细金额易误当总金额、缺销售方也不回退 OCR。
- **修复**：新增常规发票专用识别器 `GeneralInvoiceLocalRecognizer`（逻辑行 + 字段候选评分 + 分区解析 + 商品明细）；命名 `{金额}` 取价税合计不取明细金额；缺销售方回退 OCR。详见 [../development/GENERAL_INVOICE_RECOGNITION.md](../development/GENERAL_INVOICE_RECOGNITION.md)。

## 版本显示修复（2026-07-10）

- 设置页版本号一直显示 `1.0.0` → 改为优先读取 ClickOnce/VSTO 发布版本（`SettingsForm.GetApplicationVersion()`）。详见 [../development/VERSION_DISPLAY_BUG_REPORT.md](../development/VERSION_DISPLAY_BUG_REPORT.md)。

## 阶段 10/11：Outlook 启动缓慢（Resiliency 禁用）

- 阶段 10 优化 Startup（移除启动时 `My.Settings.Save()`、密钥迁移延后、新增启动耗时诊断）；阶段 11 建立启动缓慢根因诊断框架（Event ID 45 差值指标 + 问题分类 + 按需加载评估）。整合后详见 [../deployment/STARTUP_PERFORMANCE.md](../deployment/STARTUP_PERFORMANCE.md)。

---

## 遗留问题（截至 2026-07-11）

1. **Outlook 端到端人工验证**：设置页各校验弹窗、进度窗口、真实邮件分流全流程仍待在装有 Outlook 的机器验证（清单见 [../testing/OUTLOOK_MANUAL_TEST.md](../testing/OUTLOOK_MANUAL_TEST.md)）。
2. **常规发票真实样例校准**：评分阈值基于合成用例，真实票面大规模校准待更多样例。
3. **独立 taxi 类型**：`taxi_online_ticket/taxi_receipt` 真实字段名待样例（现有样例均为 vat_invoice 内嵌行程）。
4. **多页 PDF / 多条行程真实样例**、接口超时/额度不足/真实空字段场景待主动触发。
5. **发布**：正式代码签名证书替换、部署方式（ClickOnce/安装包）待定。
