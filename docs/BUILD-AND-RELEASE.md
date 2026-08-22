# oWorkHelper 构建与发布指南

## 1. 环境要求

- **Visual Studio 2022**，需安装"Office/SharePoint 开发 (VSTO)"工作负载
- **.NET Framework 4.8 开发包**（目标框架）
- **Outlook 桌面版**（调试时需要，作为 VSTO 宿主进程）
- **VSTO Runtime 4.0**（Visual Studio Tools for Office 运行时）
- **NuGet 包还原**：项目通过 `packages.config` 管理依赖，构建前需还原 NuGet 包

## 2. 项目结构

| 路径 | 说明 |
|------|------|
| `oWorkhelper.sln` | 主解决方案文件 |
| `oWorkhelper.vbproj` | VSTO Outlook 外接程序项目（VB.NET） |
| `tools/OfflineTester/OfflineTester.vbproj` | 离线测试控制台应用，用于脱离 Outlook 验证核心逻辑 |
| `packages.config` | NuGet 依赖清单（PdfPig 0.1.14 及其传递依赖） |

## 3. 编译配置

解决方案 `oWorkhelper.sln` 包含四种编译配置：

| 配置名称 | 编译常量 | 输出路径 | 在线 OCR |
|----------|----------|----------|----------|
| Debug | （无） | `bin\Debug\` | 禁用 |
| Release | （无） | `bin\Release\` | 禁用 |
| Release-Intranet | `INTRANET_BUILD` | `bin\Release-Intranet\` | 禁用 |
| Release-Internet | `INTERNET_BUILD` | `bin\Release-Internet\` | 启用 |

**默认安全策略**：只有定义了 `INTERNET_BUILD` 编译常量时，在线 OCR 功能才会启用（由 `BuildFeatures.vb` 控制）。未定义该常量的配置均默认禁用在线 OCR，防止内网环境意外调用外部服务。

## 4. 构建命令

### 构建主项目

内网版本：

```batch
MSBuild oWorkhelper.sln /p:Configuration=Release-Intranet /p:Platform=AnyCPU /restore
```

外网版本（启用在线 OCR）：

```batch
MSBuild oWorkhelper.sln /p:Configuration=Release-Internet /p:Platform=AnyCPU /restore
```

### 构建离线测试工具

```batch
MSBuild tools\OfflineTester\OfflineTester.vbproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU
```

## 5. 版本号规则

版本号格式为 `a.b.yyMMdd.d`，例如 `1.2.260713.1`。

各段含义：

| 段 | 含义 | 示例 |
|----|------|------|
| `a` | 主版本号，重大功能变更时递增 | `1` |
| `b` | 次版本号，功能增强时递增 | `2` |
| `c` | 发布日期，格式为 `yyMMdd` | `260713`（2026 年 7 月 13 日） |
| `d` | 修订号，同日多次发布时递增 | `1` |

### 版本号同步位置

以下三处版本号必须保持一致：

1. **`My Project/AssemblyInfo.vb`**：`AssemblyVersion`、`AssemblyFileVersion`、`AssemblyInformationalVersion` 三个特性
2. **`oWorkhelper.vbproj`**：`ApplicationVersion` 属性（ClickOnce 发布版本号）
3. **Git 标签**：格式为 `v{版本号}`，例如 `v1.2.260713.1`

## 6. 签名

- 清单签名使用临时证书 `iWorkhelper_TemporaryKey.pfx`（历史兼容文件名，不随品牌展示调整）
- 私钥**不包含**在代码仓库中，生产部署时需使用正式的代码签名证书
- 项目文件 `oWorkhelper.vbproj` 中的 `ManifestCertificateThumbprint` 属性指向当前使用的证书指纹

## 7. 发布流程

1. **更新版本号**：修改 `My Project/AssemblyInfo.vb` 中的三个版本特性，以及 `oWorkhelper.vbproj` 中的 `ApplicationVersion`，确保一致
2. **编译构建**：分别构建 `Release-Intranet` 和 `Release-Internet` 配置
3. **运行自测**：执行 `OfflineTester --selftest`，验证核心识别与归档逻辑正常工作
4. **提交与标签**：`git commit` 提交变更，`git tag v{版本号}` 创建标签，推送至远程仓库
5. **创建发布**：在 GitHub 上创建 Release，附上发布说明

## 8. 安装与部署 (VSTO)

- VSTO 外接程序通过注册表 `LoadBehavior=3` 配置为随 Outlook 自动加载
- 安装方式通常为 ClickOnce 发布或手动注册表配置
- Outlook 可能会禁用启动缓慢的外接程序（参见弹性机制说明）
- `tools/OutlookResiliency/` 目录包含诊断脚本，用于检查和修复 Outlook 弹性机制导致的外接程序被禁用问题

### Outlook 弹性机制

Outlook 会监控外接程序的加载时间。如果外接程序连续多次启动超时，Outlook 会自动禁用该外接程序。可使用 `tools/OutlookResiliency/` 中的脚本检查注册表中的禁用状态并进行恢复。

## 9. OfflineTester 构建与自测

OfflineTester 是一个控制台应用，可在不启动 Outlook 的情况下验证核心逻辑（发票识别、PDF 解析、归档命名等）。

### 构建

```batch
MSBuild tools\OfflineTester\OfflineTester.vbproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU
```

### 运行自测

```batch
tools\OfflineTester\bin\Debug\OfflineTester.exe --selftest
```

自测会执行内置的验证用例，覆盖以下核心功能：

- PDF 文本提取与布局分析
- 发票字段识别与分类
- 归档命名规则生成
- 编译常量与功能开关验证

自测通过后即可确认当前构建的核心逻辑工作正常，无需依赖 Outlook 环境。
