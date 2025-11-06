# iWorkHelper 技术栈与兼容性文档

> 文档版本：v0.1.0  ·  更新日期：2025-11-06  ·  适用范围：Outlook VSTO 加载项（VB / .NET Framework 4.8）

## 0. 相关文档
- 开发框架与技术架构设计：[`Doc/01-Architecture.md`](01-Architecture.md)
- 开发流程与规范：[`Doc/03-DevelopmentProcess.md`](03-DevelopmentProcess.md)
- 开发任务 TodoList：[`Doc/04-TodoList.md`](04-TodoList.md)
- 开发注意事项与常见问题：[`Doc/05-Notes.md`](05-Notes.md)
- 变更记录：[`Doc/CHANGELOG.md`](CHANGELOG.md)

## 1. 技术组件与版本
- 平台与工具
  - 操作系统：Windows 11 23H2（64 位）
  - IDE：Visual Studio Community 2022（64 位）
  - Outlook：Office 365 版本 2502（内部版本 18526.20604）
  - 模板：Outlook VSTO 外接程序
  - 语言：Visual Basic
  - 框架：.NET Framework 4.8（开发者包需安装）
- 主要依赖（NuGet / 组件）
  - UglyToad.PdfPig ≥ 0.1.9（PDF 文本提取）
  - iTextSharp 5.5.13.3（PDF 合并与处理）
  - Microsoft.Office.Interop.Outlook（随 VSTO 安装）
  - System.Configuration（框架内置，Settings 支持）
  - 可选：Newtonsoft.Json（配置或日志扩展）

## 2. 技术选型依据
- VSTO（而非 Office.js）
  - 依据：需要离线能力、深度访问本地附件与 Outlook 对象模型、较低迁移成本。
- VB.NET（而非 C#）
  - 依据：团队技能与现有代码风格，配合 .NET Framework 4.8 保持成熟稳定。
- PdfPig（文本提取）
  - 依据：对文本型 PDF 提取效果稳定、API 简洁、无需外部服务；如遇扫描版 PDF，可按需扩展 OCR。
- iTextSharp（PDF 合并）
  - 依据：合并操作稳定，接口清晰；与 .NET Framework 4.8 兼容良好。

## 3. 兼容性矩阵
| 组件 | 最低版本 | 目标/验证版本 | 状态 | 备注 |
|---|---|---|---|---|
| 操作系统 | Windows 10 1809 | Windows 11 23H2 | 已验证 | 以 Win11 为主要测试环境 |
| Outlook | 2016 | 365 v2502 (18526.20604) | 已验证 | 支持 2016+，以 365 为主 |
| .NET | 4.7.2 | 4.8 | 已验证 | 项目目标版本为 4.8 |
| IDE | VS 2019 | VS 2022 | 已验证 | 建议使用 VS 2022 |
| PdfPig | 0.1.7 | ≥0.1.9 | 建议 | 文本提取能力稳定 |
| iTextSharp | 5.5.13.1 | 5.5.13.3 | 已验证 | 合并功能稳定 |

> 注：项目要求以 VS 2022、Outlook 365 2502 与 .NET 4.8 为基准环境；对 2016/2019 版本进行兼容性验证后再发布。

## 4. 环境配置要求
- 必备安装
  - 安装 Visual Studio 2022，并勾选「Office 开发工具」工作负载（VSTO）。
  - 安装 .NET Framework 4.8 Developer Pack。
  - 安装 Microsoft Office（Outlook 365 版本 2502）。
- 项目设置
  - 目标框架：`.NET Framework 4.8`
  - 语言版本：VB 默认（启用 `Option Strict On`、`Option Explicit On`）
  - VSTO 加载项：模板为「Outlook VSTO 外接程序」。
- NuGet 包
  - 推荐使用内置包管理器安装：PdfPig（UglyToad.PdfPig）、iTextSharp。
  - 若企业内网环境需配置镜像源，建议使用 `nuget.exe` 通过 CMD 进行源设置。
- 终端偏好
  - 开发与构建相关命令建议使用 `CMD` 而非 PowerShell（遵循项目规则）。

### 4.1 Settings（应用配置）示例
```ini
# My.Settings 关键项（示例）
ArchivePath = D:\WorkArchive\Invoices
MergeDidiFiles = true
EnableProgressUI = true
```

### 4.2 Outlook 信任中心与加载项加载
- 确保已信任加载项发布者证书（ClickOnce）。
- 如果加载项未加载，检查注册表项：`HKCU\Software\Microsoft\Office\Outlook\Addins\iWorkHelper` 中的 `LoadBehavior` 是否为 `3`。

## 5. 兼容性与风险说明
- 扫描版 PDF：PdfPig 对扫描版文本提取有限，需按需引入 OCR（可选）。
- 云附件/链接附件：仅对本地实际附件执行处理；云链接需先下载到本地。
- iTextSharp 许可：遵循对应版本的许可证约束（5.x 系列为 AGPL/商业；企业需合规评估）。

---

附录：参考链接
- PdfPig：https://github.com/UglyToad/PdfPig
- iTextSharp：https://github.com/itext/itextsharp
- VSTO 开发概述：https://learn.microsoft.com/office/dev/add-ins/（Office VSTO/COM 文档）