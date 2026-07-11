# 技术栈 TECH_STACK

> 更新日期：2026-07-11
> 依据：`iWorkhelper.vbproj`、`packages.config`、`iWorkhelper.sln`、当前源码。

## 1. 语言与框架

| 项 | 值 | 依据 |
|----|----|------|
| 语言 | VB.NET | 全部源码为 `.vb` |
| 应用类型 | VSTO Outlook 加载项（COM Add-in，OutputType=Library） | `iWorkhelper.vbproj`；ProjectTypeGuids 含 VSTO GUID |
| 目标框架 | **.NET Framework 4.8** | `<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>`；`packages.config` 全部 `net48` |
| VSTO 运行时 | VSTO 4.0 | BootstrapperPackage `Microsoft.VSTORuntime.4.0` |
| 加载行为 | LoadBehavior=3（随 Outlook 自动加载） | `iWorkhelper.vbproj` |

> 注：`.trae/rules/projectIntroduction.md` 曾写 ".NET Framework 8.0"，与工程实际 **4.8** 不符。以工程文件 v4.8 为准（VSTO 传统模型运行在 .NET Framework 上，不存在".NET Framework 8.0"）。

## 2. Office / Outlook / VSTO

| 组件 | 说明 |
|------|------|
| Host | Outlook（`<OfficeApplication>Outlook</OfficeApplication>`） |
| Interop | `Microsoft.Office.Interop.Outlook` 15.0 + `Microsoft.Office.Core`，`EmbedInteropTypes=true` |
| VSTO Tools | `Microsoft.Office.Tools(.Outlook/.Common)` v10.0 |
| Ribbon | 可视化设计器（`RibbonBase`），非 XML |
| 目标 Office | Office 365 桌面版（开发参考版本 2502） |

## 3. 依赖库（packages.config，NuGet 可恢复）

| 包 | 版本 | 用途 |
|----|------|------|
| **PdfPig（UglyToad.PdfPig）** | 0.1.14 | PDF 文本抽取、页面/坐标布局分析、**PDF 合并（PdfMerger）** |
| Microsoft.Bcl.HashCode | 6.0.0 | PdfPig 传递依赖 |
| System.Buffers / System.Memory / System.Numerics.Vectors / System.Runtime.CompilerServices.Unsafe | 4.6.0 / 4.6.0 / 4.6.0 / 6.1.0 | PdfPig 传递依赖 |

- 采用 `packages.config` 模式（非 PackageReference）。
- PdfPig 引用 `lib\net471` DLL（`net48` 向下兼容，构建与运行已验证）。
- **PdfPig 是文本抽取库，非 OCR**：文本型 PDF 用 PdfPig；图片/扫描型 PDF 走百度在线 OCR 的 `pdf_file` 接口。

## 4. OCR 服务

- **百度智能云"智能财务票据识别"（multiple_invoice）**，内置 `HttpWebRequest` + `JavaScriptSerializer`，未引第三方 HTTP/JSON 库。
- 实现见 `Core/Ocr/Baidu/*` 与 [BAIDU_OCR_INTEGRATION_DESIGN.md](BAIDU_OCR_INTEGRATION_DESIGN.md)。
- 在线 OCR 由**编译期开关**门控（内网版禁用），见 [BUILD_VARIANTS.md](BUILD_VARIANTS.md)。

## 5. 配置、日志、安全

| 方面 | 实现 |
|------|------|
| 配置 | `My.Settings`（User 作用域，归档目录/命名模板/OCR 参数）+ 加密配置文件 `%AppData%\iWorkHelper\baidu-ocr.config.xml` |
| 密钥保护 | Secret Key 经 **DPAPI（CurrentUser）** 加密（`DPAPI:` 前缀），见 [../security/SECRET_STORAGE_DESIGN.md](../security/SECRET_STORAGE_DESIGN.md) |
| 日志 | `Core/Logging/AppLogger` 文件日志，写入归档目录`\logs` 或 `%AppData%\iWorkHelper\logs`；**不含 AK/SK/token 明文** |
| 异常处理 | 统一错误分类（`AppErrorCode`/`AppError`/`UserFriendlyMessageProvider`）+ 归档前预检查 + 批量失败隔离，见 [ERROR_HANDLING_DESIGN.md](ERROR_HANDLING_DESIGN.md) |

## 6. 构建 / 调试

| 环节 | 方式 |
|------|------|
| 构建 | VS2022 或 MSBuild；四套配置 Debug/Release/Release-Intranet/Release-Internet（见 [BUILD_VARIANTS.md](BUILD_VARIANTS.md)） |
| 调试 | F5 启动 Outlook（`DebugInfoExeName` 指向 outlook.exe） |
| 离线测试 | `tools/OfflineTester`（不依赖 Outlook），见 [../testing/OFFLINE_TESTER_GUIDE.md](../testing/OFFLINE_TESTER_GUIDE.md) |
| 签名 | 清单签名临时证书 `iWorkhelper_TemporaryKey.pfx`（私钥不入库，发布需替换正式证书） |

> 构建环境需：VS2022 + "Office/SharePoint 开发（VSTO）"工作负载 + .NET Framework 4.8 开发包 + 已安装 Outlook + NuGet 还原 `packages\`。

## 7. 一句话总结

> **VB.NET + VSTO 4.0 Outlook 加载项，目标 .NET Framework 4.8，Ribbon UI + My.Settings；PdfPig 文本抽取/合并 + 百度在线 OCR（编译期门控）+ DPAPI 密钥保护 + 统一日志/异常/预检查体系。**
