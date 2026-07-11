# 版本显示缺陷记录 VERSION_DISPLAY_BUG_REPORT

> 日期：2026-07-10（记录）／2026-07-11（据源码校正）
> 问题：设置页版本号一直显示 `1.0.0`，不跟随 Visual Studio 发布版本。状态：已修复。

## 问题现象

在 VS 发布界面填写应用版本（如 `1.2.711.x`），但设置页一直显示 `1.0.0`。

## 根本原因

修复前 SettingsForm 只读取 `Assembly.GetName().Version`（即 `AssemblyVersion`）。而 `AssemblyVersion` 在 `My Project/AssemblyInfo.vb` 中硬编码为 `1.0.0.0`，与 ClickOnce/VSTO 发布版本（`iWorkhelper.vbproj` 的 `<ApplicationVersion>`）无关。

版本来源与优先级：

| 来源 | 用途 | 优先级 |
|------|------|--------|
| ClickOnce 部署版本（`ApplicationDeployment.CurrentDeployment.CurrentVersion`） | 正式部署运行时实际版本 | 最高（仅正式部署有效，F5 调试为 None） |
| `AssemblyInformationalVersion` | 字符串版本，支持完整四段 | 次高 |
| `AssemblyFileVersion` | 文件版本 | 次 |
| `AssemblyVersion` | 程序集标识（硬编码 `1.0.0.0`） | 最低（fallback） |

> 说明：`AssemblyVersion`/`AssemblyFileVersion` 每段限 0–65535，若版本含发布日期编码（如 `260710`）会超限，故不宜写入这两处，应放 `AssemblyInformationalVersion` 或用 ClickOnce 版本。

## 修复

在 `SettingsForm.vb` 中实现私有方法 **`GetApplicationVersion()`**，按上表优先级读取：ClickOnce 发布版本 > `AssemblyInformationalVersion` > `AssemblyFileVersion` > `AssemblyVersion`；`SettingsForm_Load` 中 `lblVersion.Text = "版本：" & GetApplicationVersion()`（内网版再追加 `BuildFeatures.EditionDisplay` "（内网版）"）。

- **F5 调试**：`IsNetworkDeployed=False` → 回退到 `AssemblyInformationalVersion`（若配置）否则 `AssemblyVersion`（`1.0.0.0`）。
- **正式部署**：`IsNetworkDeployed=True` → 显示 ClickOnce 部署版本（完整四段）。

> 实现位置为 `SettingsForm.vb` 的 `GetApplicationVersion()` 方法（**非独立 `AppVersionProvider.vb` 文件**——早期记录曾误述为新增该文件，实际未创建）。

## 待开发者确认（源码一致性）

`My Project/AssemblyInfo.vb` 当前**未配置** `AssemblyInformationalVersion`，且 `AssemblyVersion`/`AssemblyFileVersion` 仍为 `1.0.0.0`。因此 **F5 调试下版本仍显示 `1.0.0.0`**（正式 ClickOnce 部署才显示发布版本）。若希望调试期也显示完整版本，需在 `AssemblyInfo.vb` 增加 `<Assembly: AssemblyInformationalVersion("1.2.xxxxxx.x")>`（本文档整理未改动业务代码，仅记录）。

## 部署后版本仍异常时的排查

- Manifest（`.vsto`）中 `appVersion` 是否正确。
- ClickOnce 缓存 `%LOCALAPPDATA%\Apps\2.0\` 是否缓存旧版本（提高部署版本号或清缓存）。
- `iWorkhelper.vbproj` 的 `<ApplicationVersion>` 是否为预期值。
