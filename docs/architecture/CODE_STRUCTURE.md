# 代码结构 CODE_STRUCTURE

> 更新日期：2026-07-11
> 本文以当前源码为准（`Core/` 已完整实现，约 68 个业务 `.vb` 文件）。

## 1. 顶层结构

```
iWorkhelper/
├── iWorkhelper.sln / .vbproj / packages.config   # 解决方案 / 工程 / NuGet 依赖
├── ThisAddIn.vb (+ .Designer.vb/.xml)            # 加载项入口（Startup/Shutdown，轻量）
├── MainRibbon.vb (+ .Designer.vb/.resx)          # 功能区：归档 / 设置 按钮
├── SettingsForm.vb (+ .Designer.vb/.resx)        # 设置窗体：归档目录/命名模板/OCR 配置
├── ProgressForm.vb (+ .Designer.vb/.resx)        # 归档进度窗口
├── My Project/                                   # AssemblyInfo / Settings / Resources
├── Core/                                         # 业务核心（见下）
└── tools/OfflineTester/                          # 离线测试工具（不依赖 Outlook）
```

## 2. Core 业务模块

| 目录 | 文件数 | 职责 |
|------|-------|------|
| `Core/Common/` | 9 | 结果对象、路径/文件名工具、异常格式化、错误码/错误分级/友好文案、编译期开关（BuildFeatures） |
| `Core/Logging/` | 1 | `AppLogger` 线程安全文件日志（失败不崩溃主流程） |
| `Core/Diagnostics/` | 1 | `StartupPerformanceTracker` 启动耗时诊断 |
| `Core/Configuration/` | 3 | 百度 OCR 配置 POCO、XML 加密配置存储、配置读取门面 |
| `Core/Invoice/` | 4 | 发票/行程/明细字段模型 + 字段名常量 |
| `Core/Pdf/` | 7 | PdfPig 文本抽取、坐标行重建、表格区检测、PDF 合并 |
| `Core/Recognition/` | 16 | 识别管道、本地文本识别器、常规发票专用识别器（候选评分/分区/明细）、关键字段评估、识别结果合并 |
| `Core/Ocr/Baidu/` | 7 | 百度 token/HTTP/解析/类型映射/字段映射 |
| `Core/Mail/` | 5 | 邮件附件读取、按邮件分组 |
| `Core/Archive/` | 8 | 命名规则/模板引擎、归档规划/执行、报告写入 |
| `Core/Security/` | 2 | DPAPI 密钥加密（SecretProtector / ProtectedSettingsProvider） |
| `Core/Workflow/` | 12 | 批量归档编排、分流分类、预检查、运行锁、进度上报 |

> `Core/Online OCR/` 为历史遗留空文件夹（早期占位），实际 OCR 代码在 `Core/Ocr/Baidu/`。

## 3. 关键类与入口

| 类型 | 类 | 说明 |
|------|-----|------|
| 入口编排 | `Core/Workflow/BatchArchiveWorkflow` | 全链路串联：分组→识别→分流→命名→归档→汇总 |
| 识别调度 | `Core/Recognition/RecognitionPipeline` | 本地优先、必要时在线 OCR、来源标记（LocalText/BaiduOcr/Mixed/Failed） |
| 本地识别 | `LocalTextInvoiceRecognizer` | 滴滴行程单坐标行解析；非滴滴 VatInvoice 委派 `GeneralInvoiceLocalRecognizer` |
| 分流分类 | `MailProcessingClassifier` | 滴滴/常规/未识别/无 PDF |
| 命名 | `ArchiveNamingRule` + `NamingTemplateEngine` | 单一模板 + fallback；占位符单一来源 `SupportedPlaceholders()` |
| 运行锁 | `ArchiveRunGuard` + `ArchiveRunToken` | `Interlocked` 原子获取，`Using` 释放（见 [../development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md](../development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md)） |
| 预检查 | `ArchivePreflightChecker` | 静态检查配置/目录/权限（不检查运行状态） |
| 版本显示 | `SettingsForm.GetApplicationVersion()` | ClickOnce > InformationalVersion > FileVersion > AssemblyVersion（见 [../development/VERSION_DISPLAY_BUG_REPORT.md](../development/VERSION_DISPLAY_BUG_REPORT.md)） |

## 4. 命名与风格约定

- `Option Strict Off`、`Option Explicit On`、`Option Infer On`（`iWorkhelper.vbproj`）。
- 控件/成员 PascalCase 与 camelCase 混用（WinForms 惯例）；UI 与注释含中文。
- 离线安全原则：`Core/` 业务代码不依赖 Outlook/`My.Settings`，因此 `tools/OfflineTester` 可直接链接同一批源文件做离线测试（见 [../testing/OFFLINE_TESTER_GUIDE.md](../testing/OFFLINE_TESTER_GUIDE.md)）。
