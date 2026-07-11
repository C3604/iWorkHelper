# iWorkHelper

## 项目简介

iWorkHelper 是一个基于 Outlook VSTO、VB.NET 和 .NET Framework 4.8 的 Outlook 加载项，用于批量处理邮件中的 PDF 附件，完成识别、合并、命名和归档。项目重点覆盖滴滴发票/行程单场景，同时支持常规发票与未识别 PDF 的分类处理。

## 核心功能

- 在 Outlook Ribbon 中提供“归档”和“设置”入口。
- 按邮件批量读取 PDF 附件并执行归档流程。
- 滴滴发票与行程单按邮件合并为单个 PDF 后识别、命名和归档。
- 常规发票单独识别、命名和归档。
- 未识别 PDF 按固定规则保留原始文件名信息。
- 本地解析基于 PdfPig 文本抽取与规则识别。
- 外网版本支持百度 OCR 在线解析；本地字段不足时可自动回退在线 OCR。
- 支持自定义命名模板、预检查、进度展示与结果报告。
- 百度 OCR Secret Key 通过 Windows DPAPI 保护。
- 提供 `OfflineTester` 离线测试与诊断工具。

## 技术栈

- VB.NET
- .NET Framework 4.8
- VSTO / Microsoft Office Tools for Outlook
- WinForms
- PdfPig 0.1.14
- 百度 OCR（外网版本可选）
- Windows DPAPI
- MSBuild / Visual Studio 17 解决方案格式

## 项目结构

```text
iWorkhelper/
├── Core/                  # 归档、识别、OCR、配置、安全、日志、工作流核心代码
├── docs/                  # 架构、部署、测试、阶段报告与维护文档
├── My Project/            # VB.NET 项目设置、资源与程序集信息
├── Resources/             # Ribbon 图标等静态资源
├── tools/
│   ├── OfflineTester/     # 独立离线测试控制台工程
│   └── OutlookResiliency/ # Outlook 加载项诊断脚本
├── iWorkhelper.sln        # 主解决方案
├── iWorkhelper.vbproj     # Outlook VSTO 主工程
└── packages.config        # NuGet 依赖清单（旧式 packages.config）
```

## 构建要求

- 建议使用支持 .NET Framework 4.8 和 Office/VSTO 开发的 Visual Studio 版本；当前解决方案格式为 Visual Studio 17，推荐 Visual Studio 2022。
- 需要安装 Office/Outlook 桌面版、VSTO Runtime 和相应的 Office 开发组件。
- 项目目标框架为 `.NET Framework 4.8`。
- 主工程为 Outlook 加载项，实际调试和加载行为依赖 Windows + Outlook 桌面环境。
- 依赖通过 `packages.config` 管理，可使用 MSBuild Restore 恢复。

## 构建配置

- `Debug`：调试构建。
- `Release`：发布构建，但未定义 `INTERNET_BUILD`，按安全默认策略禁用在线解析。
- `Release-Intranet`：定义 `INTRANET_BUILD`，禁用在线 OCR，并在设置界面隐藏在线解析相关 UI。
- `Release-Internet`：定义 `INTERNET_BUILD`，允许配置本地解析与百度 OCR 在线解析。

项目当前编译常量策略以 `Core/Common/BuildFeatures.vb` 为准：只有定义了 `INTERNET_BUILD` 时才启用在线解析，否则默认关闭。

## 配置说明

用户设置保存在 `My.Settings`，主要包括：

- 归档目录：`ArchiveFolderPath`
- 解析模式：`ParseMode`
- OCR 开关：`OcrEnabled`
- 百度 OCR 接口地址与 Token 地址
- 自定义命名模板

敏感信息约定：

- `BaiduApiKey` 可保存为普通配置。
- `BaiduSecretKey` 通过 `ProtectedSettingsProvider` 使用 DPAPI 加密存储。
- 兼容的本机外部配置文件为 `%AppData%\iWorkHelper\baidu-ocr.config.xml`，不应提交入库。

示例占位：

```text
API Key: <your-api-key>
Secret Key: <your-secret-key>
```

## 默认命名规则

以当前代码行为为准：

- 滴滴统一归档命名：`{乘车日期}_{金额}_{出发地点}_{到达地点}`
- 常规发票命名：`{开票日期}_{金额}_{销售方名称}`
- 未识别 PDF：`未识别_{原始文件名}`

说明：

- `{乘车日期}` 优先取行程出发日期，其次取开票日期。
- `{金额}` 优先取行程金额，其次取价税合计。
- 当字段严重不足时，会回退到内部 fallback 规则：`未识别票据_{邮件主题或原附件名}_{时间戳}`。

## 使用流程

1. 在 Outlook 中选中需要处理的邮件。
2. 点击 Ribbon 中的“归档”按钮。
3. 插件读取邮件中的 PDF 附件并执行预检查。
4. 按邮件类型完成合并、识别、命名和归档。
5. 在进度窗口查看状态，并在结束后查看结果汇总、日志和归档报告。

## OfflineTester

`OfflineTester` 位于 `tools/OfflineTester/`，用于在不启动 Outlook 的情况下验证本地解析、OCR 集成和命名规则。

常见命令示例：

```text
OfflineTester.exe --selftest
OfflineTester.exe <pdf或目录> --local-only
OfflineTester.exe <pdf或目录> --ocr
OfflineTester.exe <pdf或目录> --general-invoice --dump-candidates
OfflineTester.exe --parse-ocr-json <脱敏json文件>
OfflineTester.exe --preflight [归档目录]
OfflineTester.exe --save-baidu-config
OfflineTester.exe <pdf或目录> --use-config --force-ocr
```

## 隐私与安全

- 发票、行程单和邮件附件可能包含敏感信息，不应提交真实票据样例。
- 百度 OCR 仅在外网版本且用户显式启用时使用。
- 内网版默认禁止在线解析，并隐藏在线 OCR 相关 UI。
- Secret Key 使用 Windows DPAPI 保护；日志中不应出现 AK、SK 或 Access Token 明文。
- 日志与归档结果报告默认对邮件主题、附件名和本机路径做脱敏处理，避免直接落盘真实内容。
- 本仓库不应包含真实证书私钥、真实用户配置或本机发布产物。

## 已知限制

- 仅支持 Windows 和 Outlook 桌面版。
- 依赖 VSTO、Office 运行时和 .NET Framework 4.8。
- 本地规则识别依赖 PDF 文本层质量，扫描件或图片型 PDF 可能需要 OCR。
- Outlook 可能因加载性能、签名或信任策略禁用加载项。
- 构建与发布仍受 VSTO 签名证书和目标环境配置影响。

## 开发状态

项目处于持续开发和内部验证阶段。源码已包含主流程、离线测试与多份设计/阶段文档，但仍需结合实际 Outlook 环境继续验证构建、签名、加载和归档效果。

## 许可证

当前仓库暂未声明开源许可证。未经授权，不得复制、修改或分发本仓库内容。
