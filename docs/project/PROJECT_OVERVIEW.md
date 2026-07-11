# 项目概览 PROJECT_OVERVIEW

> 更新日期：2026-07-11
> 本文以当前源码为准，反映项目实际实现状态。历史演进见 [../history/DEVELOPMENT_HISTORY.md](../history/DEVELOPMENT_HISTORY.md)。

## 1. 项目定位

**iWorkHelper** 是一个 **Outlook VSTO 加载项**（COM Add-in），使用 **VB.NET + .NET Framework 4.8** 开发，在 Outlook 功能区提供"工作助手"入口。核心功能：**批量选中邮件 → 读取 PDF 附件 → 识别票据（本地文本解析优先，百度 OCR 兜底）→ 按模板命名 → 归档到指定目录**。

主要面向滴滴出行发票/行程单与常规增值税发票的自动整理归档。

## 2. 核心业务能力（均已实现）

| 能力 | 实现模块 | 说明 |
|------|---------|------|
| Ribbon 入口 + 设置界面 | `MainRibbon`、`SettingsForm` | 归档目录、命名模板、OCR 配置 |
| 读取选中邮件 PDF 附件 | `Core/Mail/*` | 按邮件分组，跳过非邮件/无 PDF，释放 COM |
| PDF 文本抽取 + 坐标行重建 | `Core/Pdf/*`（PdfPig） | 文本层抽取、疑似图片型判定、按列取行程单单元格 |
| PDF 合并 | `Core/Pdf/PdfMergeService`（PdfPig PdfMerger） | 滴滴发票+行程单合并为一份 |
| 本地识别 | `Core/Recognition/*` | 滴滴行程单 + 常规发票（候选评分/分区/明细） |
| 在线 OCR | `Core/Ocr/Baidu/*` | 百度智能财务票据识别 multiple_invoice |
| 邮件/PDF 分流 | `Core/Workflow/MailProcessingClassifier` | 滴滴合并 / 常规单独 / 未识别 / 无 PDF 跳过 |
| 命名与归档 | `Core/Archive/*` | 单一模板 + fallback，同名不覆盖 |
| 密钥保护 | `Core/Security/*`（DPAPI） | Secret Key 加密存储 |
| 容错与预检查 | `Core/Workflow/*`、`Core/Common/*` | 归档前预检查、运行锁、统一错误分类、进度条 |
| 日志与报告 | `Core/Logging/AppLogger`、`ArchiveReportWriter` | 日志无 AK/SK/token 明文，批次报告 |

## 3. 端到端流程

选中邮件 → 点"归档" → 获取运行锁 → 归档前预检查 → 进度窗口 → 按邮件分组 → 逐 PDF 识别（本地优先/OCR 兜底）→ 分流路由（滴滴合并/常规单独/未识别）→ 命名 → 归档 → 汇总弹窗 + 报告。

详见 [../architecture/PROCESS_FLOW.md](../architecture/PROCESS_FLOW.md)。

## 4. 验证状态

- **离线验证（真实执行）**：OfflineTester `--selftest` 全通过（含 DPAPI、命名、分流、常规发票识别、运行锁、合成 OCR JSON）；5 份真实滴滴样例本地识别命名核心字段 5/5；百度 OCR 已用真实 AK/SK 联调成功（Mixed，31 字段）；三套构建（Debug/Release-Intranet/Release-Internet）通过。
- **待 Outlook 端到端人工验证**：设置页各校验弹窗、进度窗口显示、真实邮件分流全流程。清单见 [../testing/OUTLOOK_MANUAL_TEST.md](../testing/OUTLOOK_MANUAL_TEST.md)。

## 5. 已知限制

- 仅处理 **PDF 附件**，不处理图片附件（图片型 PDF 走百度 OCR 的 `pdf_file`）。
- 常规发票本地识别的评分阈值基于合成用例与通用规律，**真实票面大规模校准待更多样例**。
- 独立 `taxi_online_ticket/taxi_receipt` 行程单类型为兼容性预留，真实字段名待样例校准。
- 归档在 UI 线程执行（`Application.DoEvents` 刷新），无取消按钮。
- 发布使用临时签名证书，正式发布需替换正式代码签名证书。

## 6. 技术栈与构建

- VB.NET / .NET Framework 4.8 / VSTO 4.0 / Outlook 加载项；依赖 PdfPig 0.1.14（NuGet `packages.config`）。
- 三套发布配置：`Release-Intranet`（内网版，禁在线 OCR）、`Release-Internet`（外网版，启用在线 OCR）、`Release`/`Debug`。详见 [../architecture/BUILD_VARIANTS.md](../architecture/BUILD_VARIANTS.md) 与 [../architecture/TECH_STACK.md](../architecture/TECH_STACK.md)。
