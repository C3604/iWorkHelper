# iWorkHelper 架构文档

## 1. 项目定位

iWorkHelper 是一个 Outlook VSTO 加载项，基于 VB.NET / .NET Framework 4.8 开发。其核心功能是批量处理邮件中的 PDF 附件，完成发票识别、PDF 合并、文件命名和归档操作。用户在 Outlook 中选中一批邮件后，一键即可将附件中的发票 PDF 按照识别结果自动命名并归档到指定目录。

## 2. 技术栈

| 组件 | 技术 |
|------|------|
| 语言 | VB.NET |
| 运行时 | .NET Framework 4.8 |
| Office 集成 | VSTO 4.0 (Visual Studio Tools for Office) |
| UI 框架 | WinForms (Ribbon + 设置窗体 + 进度窗体) |
| PDF 处理 | PdfPig 0.1.14 (文本提取、坐标行重建、表格区域检测、PDF 合并) |
| OCR 识别 | 百度 OCR 多票识别 API (仅外网版本) |
| 密钥保护 | Windows DPAPI (Data Protection API, CurrentUser 范围) |
| 构建工具 | MSBuild / Visual Studio 2022 |

## 3. 代码结构

### 顶层文件

- **ThisAddIn.vb** — VSTO 加载项入口，负责加载项生命周期管理
- **MainRibbon.vb** — Outlook Ribbon 界面，提供用户操作按钮
- **SettingsForm.vb** — 设置窗体，管理百度 OCR 配置及归档路径等选项
- **ProgressForm.vb** — 进度窗体，展示批量处理进度

### Core 模块

#### Core/Common/ (通用基础设施，9 个文件)

- **Result.vb** — 泛型结果类型，封装操作成功/失败状态及错误信息
- **PathHelper.vb** — 路径处理辅助方法
- **FileNameSanitizer.vb** — 文件名清理，移除非法字符
- **PrivacySafeFormatter.vb** — 隐私安全格式化器，对邮件主题和附件名进行脱敏处理
- **ExceptionFormatter.vb** — 异常格式化，生成结构化异常信息
- **ErrorSeverity.vb** — 错误严重级别枚举
- **AppErrorCode.vb** — 应用错误代码枚举，标识各类错误场景
- **AppError.vb** — 应用错误类，结合错误代码与严重级别
- **UserFriendlyMessageProvider.vb** — 面向用户的友好错误消息生成器
- **BuildFeatures.vb** — 编译特性开关，控制内外网版本差异
- **ExplorerFolderService.vb** — 归档后打开或激活归档目录，使用 Shell.Application.Windows() 枚举已有窗口

#### Core/Logging/ (日志，1 个文件)

- **AppLogger.vb** — 线程安全的文件日志记录器

#### Core/Diagnostics/ (诊断，1 个文件)

- **StartupPerformanceTracker.vb** — 启动性能追踪器，记录加载项各阶段耗时

#### Core/Configuration/ (配置管理，3 个文件)

- **BaiduOcrOptions.vb** — 百度 OCR 配置选项（API Key、Secret Key、API 端点等）
- **BaiduXmlConfigStore.vb** — 基于 XML 文件的配置持久化存储
- **OcrConfigProvider.vb** — OCR 配置提供者，统一配置读取入口

#### Core/Invoice/ (发票数据模型，4 个文件)

- **InvoiceFieldNames.vb** — 发票字段名称常量定义
- **InvoiceTripInfo.vb** — 行程信息（出发地、到达地、乘车日期等，用于滴滴发票）
- **InvoiceLineItem.vb** — 发票明细行项目
- **InvoiceInfo.vb** — 发票信息聚合模型，包含金额、日期、销售方、行程明细等

#### Core/Pdf/ (PDF 处理，7 个文件)

基于 PdfPig 库实现 PDF 文本提取与处理：

- **PdfTextExtractor.vb** — PDF 文本提取器，从 PDF 中提取原始文本块
- **PdfTextLayoutExtractor.vb** — PDF 文本布局提取器，按坐标重建文本行
- **PdfTextBlock.vb** — PDF 文本块数据模型
- **PdfTextLine.vb** — PDF 文本行数据模型
- **PdfTextExtractResult.vb** — PDF 文本提取结果
- **PdfTableRegionDetector.vb** — 表格区域检测器，识别 PDF 中的表格结构
- **PdfMergeService.vb** — PDF 合并服务，将多个 PDF 文件合并为一个

#### Core/Recognition/ (识别引擎，16 个文件)

识别管线，负责从 PDF 文本或 OCR 结果中提取发票关键信息：

- 识别管线协调器，按照本地识别 → 百度 OCR 回退的顺序执行
- 本地文本识别器，基于 PdfPig 提取的文本进行正则匹配
- 滴滴行程发票识别器，专门处理滴滴出行电子发票
- 常规发票识别器，使用候选评分机制识别一般增值税发票
- 候选评分器，对多个候选字段值评分取最优
- 关键字段评估器，判断识别结果是否满足命名所需的最低字段要求
- 识别结果合并器，将多来源识别结果整合
- 百度 OCR 识别器，调用百度多票识别 API 进行在线识别

#### Core/Ocr/Baidu/ (百度 OCR 集成，7 个文件)

- **BaiduAccessTokenProvider.vb** — 百度 OCR 访问令牌获取与缓存
- **BaiduAccessTokenResult.vb** — 令牌请求结果
- **BaiduOcrHttpClient.vb** — 百度 OCR HTTP 客户端，发送 PDF 文件并获取识别结果
- **BaiduOcrRawResponse.vb** — 百度 OCR 原始响应数据模型
- **BaiduMultipleInvoiceResponseParser.vb** — 多票识别响应解析器
- **BaiduInvoiceTypeMapper.vb** — 百度发票类型到内部类型的映射
- **BaiduInvoiceFieldMapper.vb** — 百度发票字段到内部字段的映射

#### Core/Mail/ (邮件处理，5 个文件)

- **MailAttachmentReader.vb** — 邮件附件读取器，从 Outlook MailItem 中提取 PDF 附件
- **MailAttachmentItem.vb** — 邮件附件数据模型
- **MailPdfGroup.vb** — 单封邮件的 PDF 分组
- **MailPdfGroupingResult.vb** — 邮件 PDF 分组结果
- **MailReadResult.vb** — 邮件读取结果

#### Core/Archive/ (归档处理，8 个文件)

- **NamingTemplates.vb** — 命名模板定义（滴滴、常规发票、未识别各有对应模板）
- **NamingTemplateEngine.vb** — 模板引擎，将占位符替换为实际识别值
- **ArchiveNamingRule.vb** — 归档命名规则（滴滴发票、常规发票）
- **UnknownPdfNamingRule.vb** — 未识别 PDF 的命名规则
- **ArchivePlanner.vb** — 归档计划器，根据识别结果生成归档计划
- **ArchiveExecutor.vb** — 归档执行器，按计划将文件复制到目标目录
- **ArchiveResult.vb** — 归档结果数据模型
- **ArchiveReportWriter.vb** — 归档报告生成器，汇总处理结果

#### Core/Security/ (安全模块，2 个文件)

- **SecretProtector.vb** — 基于 DPAPI 的密钥加密/解密
- **ProtectedSettingsProvider.vb** — 受保护的配置读写提供者，自动处理加密/解密

#### Core/Workflow/ (工作流编排，12 个文件)

- 批量归档工作流编排器，协调整个处理流程
- 邮件分类器，将邮件附件分为滴滴/常规/未识别/无 PDF 四类
- 预检检查器，在流程启动前验证前置条件（目标目录存在性等）
- 运行锁（RunGuard），使用 Interlocked 原子操作防止并发执行
- 进度报告器，向 ProgressForm 推送处理进度

## 4. 核心处理流程

整个归档流程如下：

```
用户在 Outlook 中选中邮件 → 点击 Ribbon 上的归档按钮
    │
    ▼
RunGuard 获取运行锁（Interlocked 原子操作，防止重复执行）
    │
    ▼
预检检查（Preflight Check）
  - 验证是否选中了邮件
  - 验证归档目标目录是否存在
  - 验证 OCR 配置是否完整（仅外网版本）
    │
    ▼
BatchArchiveWorkflow 启动批量处理
    │
    ├─ 1. 读取邮件 PDF 附件
    │     MailAttachmentReader 从每封邮件中提取 PDF 附件
    │
    ├─ 2. 分类（Classification）
    │     将附件分为四类：
    │     - 滴滴发票（Didi）：根据文件名或内容特征识别
    │     - 常规发票（General）：一般增值税发票
    │     - 未识别（Unknown）：无法分类的 PDF
    │     - 无 PDF（NoPDF）：邮件中无 PDF 附件
    │
    ├─ 3. 识别（Recognition）
    │     │
    │     ├─ 滴滴发票：先合并同一邮件内多个 PDF，再整体识别
    │     │   └─ PdfMergeService 合并 → 识别管线
    │     │
    │     └─ 常规发票：逐个 PDF 独立识别
    │         └─ 识别管线
    │
    │     识别管线执行顺序：
    │     ① PdfPig 提取文本
    │     ② 本地文本识别器（正则匹配）
    │        - 滴滴：提取乘车日期、金额、出发地、到达地
    │        - 常规：提取开票日期、金额、销售方名称
    │     ③ 关键字段评估：判断本地识别结果是否充分
    │     ④ 若不充分，回退到百度 OCR（仅外网版本可用）
    │     ⑤ 合并多来源识别结果
    │
    ├─ 4. 命名（Naming）
    │     模板驱动的文件命名：
    │     - 滴滴发票：{乘车日期}_{金额}_{出发地点}_{到达地点}.pdf
    │     - 常规发票：{开票日期}_{金额}_{销售方名称}.pdf
    │     - 未识别：未识别_{原始文件名}.pdf
    │     NamingTemplateEngine 负责占位符替换，FileNameSanitizer 清理非法字符
    │
    └─ 5. 归档（Archive）
          ArchivePlanner 生成归档计划
          ArchiveExecutor 将文件复制到目标目录（含同名冲突解决）
          ArchiveReportWriter 生成处理报告
          ExplorerFolderService 打开或激活归档目录
          RunGuard 释放运行锁
```

## 5. 内外网版本差异

项目通过 `BuildFeatures.vb` 作为唯一的编译特性开关，控制内外网版本差异。

### 编译配置

项目定义了 4 种构建配置：

| 配置名 | 编译常量 | 用途 |
|--------|----------|------|
| Debug | 无 | 开发调试，默认离线模式 |
| Release | 无 | 通用发布，默认离线模式 |
| Release-Intranet | INTRANET_BUILD | 内网发布版本 |
| Release-Internet | INTERNET_BUILD | 外网发布版本，启用百度 OCR |

### 安全默认原则

未定义任何编译常量时，项目默认运行在离线模式下，不会尝试任何网络请求。只有显式定义 `INTERNET_BUILD` 编译常量时，百度 OCR 在线识别功能才会被启用。

### 条件编译控制

`BuildFeatures.vb` 中使用 `#If INTERNET_BUILD` 条件编译指令，暴露只读属性供其他模块查询。各模块无需关心编译常量细节，统一通过 `BuildFeatures` 类判断当前运行模式。

受影响的功能点：
- 百度 OCR 在线识别功能的启用与禁用
- 设置窗体中 OCR 相关配置项的显示与隐藏
- 预检检查中 OCR 配置完整性验证的开关

## 6. 错误处理设计

### 错误建模

错误处理采用结构化设计，由三个核心类型组成：

- **AppErrorCode 枚举** — 定义所有已知错误场景的唯一标识码，每个错误码对应一种明确的失败原因
- **AppError 类** — 组合错误码与严重级别，携带上下文信息，支持链式追溯原始异常
- **UserFriendlyMessageProvider** — 将 AppError 转换为面向用户的中文友好消息，隔离技术细节与用户展示

### 预检机制

在批量处理流程启动前执行预检检查（Preflight Check），提前验证各项前置条件是否满足。预检阶段不依赖运行时状态，仅检查静态配置和环境条件。预检失败时直接向用户报告问题，不进入处理流程。

### 运行锁

RunGuard 使用 `Interlocked.CompareExchange` 原子操作实现运行锁，确保同一时刻只有一个批量处理流程在运行。锁的获取和释放均为原子操作，无需传统的锁对象，避免死锁风险。

### 逐项故障隔离

批量处理中，单个邮件或单个附件的处理失败不会中断整个流程。每个处理项独立捕获异常，记录错误信息后继续处理下一项。最终通过归档报告汇总所有成功和失败的处理结果。

## 7. 安全设计

### 密钥保护

百度 OCR 的 API Key 和 Secret Key 通过 Windows DPAPI (Data Protection API) 加密存储：

- **加密范围**：使用 CurrentUser 范围，仅当前 Windows 用户可解密
- **存储格式**：加密后的密文以 `DPAPI:` 前缀存储在 XML 配置文件中，前缀用于区分明文与密文
- **SecretProtector** 负责加密和解密操作
- **ProtectedSettingsProvider** 封装配置读写，自动处理加密/解密透明转换

### 明文自动迁移

打开设置窗体时，若检测到配置文件中存在未加密的明文密钥（无 `DPAPI:` 前缀），系统自动将其加密后重新保存，实现透明的安全迁移。

### 日志脱敏

- API Key、Secret Key、Access Token 等敏感信息不会出现在日志中
- **PrivacySafeFormatter** 对邮件主题和附件名称进行脱敏处理，防止个人信息泄露到日志文件

## 8. 百度 OCR 集成

百度 OCR 集成仅在外网版本（`INTERNET_BUILD`）中可用，使用百度多票识别 API（`multiple_invoice`）实现在线发票识别。

### 调用流程

```
BaiduAccessTokenProvider 获取 Access Token
  - 使用 API Key + Secret Key 换取 Token
  - Token 带有效期缓存，过期自动刷新
      │
      ▼
BaiduOcrHttpClient 发送识别请求
  - HTTP POST 请求
  - 将 PDF 文件以 Base64 编码作为 pdf_file 参数提交
  - 调用百度 multiple_invoice API 端点
      │
      ▼
BaiduMultipleInvoiceResponseParser 解析响应
  - 解析 JSON 响应体
  - 提取每张发票的识别结果
      │
      ▼
BaiduInvoiceTypeMapper 映射发票类型
  - 将百度返回的发票类型编码映射为内部发票类型枚举
      │
      ▼
BaiduInvoiceFieldMapper 映射发票字段
  - 将百度返回的字段名称映射为内部 InvoiceInfo 模型字段
  - 统一数据格式（日期格式、金额格式等）
```

### 回退策略

百度 OCR 作为本地识别的补充手段，仅在以下条件同时满足时触发：

1. 当前为外网版本（`INTERNET_BUILD` 编译常量已定义）
2. 本地文本识别结果不充分（关键字段缺失，无法满足命名模板要求）
3. 百度 OCR 配置完整且有效（API Key、Secret Key 已配置）

本地识别结果充分时，不会发起百度 OCR 请求，避免不必要的网络调用和 API 配额消耗。
