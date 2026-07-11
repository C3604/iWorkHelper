# Outlook 启动性能与慢启动诊断 STARTUP_PERFORMANCE

> 更新日期：2026-07-11
> 本文整合原启动性能优化、根因排查、证据收集与按需加载评估等多份文档。
> 面向：开发者与维护者。用户侧恢复操作见 [OUTLOOK_ADDIN_RESILIENCY_GUIDE.md](OUTLOOK_ADDIN_RESILIENCY_GUIDE.md)。

## 1. 问题

Outlook 启动时反复提示"此加载项导致 Outlook 启动缓慢，已被禁用"，即使选"30 天内不提醒"仍循环出现，Ribbon 中"工作助手"按钮消失。根因是 Outlook Resiliency 检测到加载项启动（或关闭）耗时超过阈值（通常 ~1000ms）。

## 2. 已实施的代码优化（阶段 10）

针对 Startup 中的磁盘 I/O 瓶颈：

- **移除启动时的 `My.Settings.Save()`**：原 `ThisAddIn_Startup` 每次都调用 `ProtectedSettingsProvider.MigratePlaintextIfNeeded()`，其中 `My.Settings.Save()` 写 `user.config`（在 OneDrive 等云盘上可能 50–500ms 甚至更久）。
- **密钥迁移延后**：迁移改到 `SettingsForm.Load` 与 `ArchivePreflightChecker.Check`（用户主动操作时），Startup 不再触发。
- **启动耗时诊断**：新增 `Core/Diagnostics/StartupPerformanceTracker.vb`，用 `Stopwatch` 记录各阶段耗时到日志（`Startup 总耗时：XXms`）。
- **Shutdown 轻量化**：仅记日志、不做耗时清理，避免关闭慢导致下次禁用。

**已确认无需改动**：Ribbon `Load` 为空、业务对象均按钮点击时才创建、无模块级重依赖提前加载（PdfPig 等仅在使用时加载）。

**健康指标**：Startup 总耗时 < 500ms 健康，500–1000ms 偏慢，> 1000ms 易被禁用；Shutdown < 200ms。

## 3. 若优化后仍被禁用：根因诊断

代码优化无效时，说明大部分耗时发生在 VSTO 外层（Manifest 验证、依赖加载、ClickOnce 更新检查、部署/版本问题、关闭慢）。需收集 **Outlook 真实数据**再对症修复，不要继续盲目优化代码。

### 3.1 关键诊断指标：差值

```
差值 = Outlook Event ID 45 的 TimeTaken − iWorkHelper 日志的 Startup 总耗时
```

| 差值 | 结论 | 方向 |
|------|------|------|
| < 50ms | 问题在 Startup 代码 | 继续优化代码（收益递减） |
| 50–200ms | 中等外层开销 | Ribbon/依赖 |
| > 200ms | 问题在 VSTO 外层 | 部署/版本/**按需加载** |

### 3.2 证据收集

1. **Windows 事件查看器**（`eventvwr.msc` → Windows 日志 → 应用程序，来源 Outlook，筛选事件 ID **45**/**59**）：
   - Event ID 45：`TimeThreshold` 与 `TimeTaken`（启动或关闭耗时超阈值）。
   - Event ID 59：启动/关闭异常堆栈（立即禁用）。
2. **iWorkHelper 日志**：`{归档目录}\logs\yyyy-MM-dd.log` 中的 `Startup 总耗时` / `Shutdown 总耗时`。
3. **注册表 / 部署**（诊断脚本 `tools/OutlookResiliency/check_iworkhelper_addin_registration.ps1` 自动扫描）：
   - `LoadBehavior` 值：3=正常加载、9=被禁用、16=按需加载、2/8=加载失败。
   - `...\Outlook\Resiliency\DisabledItems` / `CrashingAddinList` 是否含 iWorkHelper。
   - Outlook 实际加载的 DLL 路径与版本，是否有旧版本残留、是否从 OneDrive 加载。

先关闭 Outlook（`Taskkill /F /IM OUTLOOK.EXE`）再收集，避免残留状态。

### 3.3 问题分类与修复方向

| 类型 | 特征 | 修复 |
|------|------|------|
| A 代码慢 | 差值 < 50ms | 继续优化 Startup（已做） |
| B Ribbon 慢 | 差值 100–300ms | 优化 Ribbon Load |
| C VSTO/依赖慢 | 差值 > 300ms | **按需加载** / 依赖优化 / 关闭自动更新检查 |
| D 版本/部署 | 多版本、加载旧版本、OneDrive 同步延迟 | 清理旧版本、重新部署 |
| E 关闭慢 | Shutdown > 500ms | 优化 Shutdown、及时释放 COM |

## 4. 备选方案：按需加载（LoadBehavior=16）

若诊断为类型 C（问题在 VSTO 外层），最有效的方案是把 `LoadBehavior` 由 3（启动加载）改为 16（按需加载）：Outlook 启动时不加载业务逻辑，用户点击按钮时才加载，从而不被禁用。

- **Manifest**：在 `.vsto` 中设置 `loadBehavior="16"`（或 post-build 脚本设置）。
- **注册表（开发/测试）**：`HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\iWorkhelper.ThisAddIn\LoadBehavior = 16`。
- **代码**：保持 Startup 极轻量（仅日志+Ribbon UI），业务逻辑已在按钮点击事件中创建，改动小。
- **权衡**：Outlook 启动更快、不被禁用；首次点击按钮有轻微延迟（< 500ms 可接受）；Ribbon 仍显示。
- **回退**：改回 `loadBehavior="3"` 即可。

## 5. 为什么不推荐"注册表强制启用"绕过

`DoNotDisableAddinList` / 删除 `DisabledItems` / 强改 `LoadBehavior=3` 只能临时生效——若启动仍慢，Outlook 每次启动重新测量后会再次禁用，形成循环，且注册表操作有风险。根治应让 Startup 足够快或改按需加载。用户侧详细说明见 [OUTLOOK_ADDIN_RESILIENCY_GUIDE.md](OUTLOOK_ADDIN_RESILIENCY_GUIDE.md)。

## 6. 状态

- 代码优化（阶段 10）已完成，**待真实 Outlook 环境验证**（连续启动 5 次对比 Event ID 45 的 TimeTaken）。
- 诊断框架与脚本已就绪；若问题复现，按第 3 节收集数据、按分类修复（最可能为按需加载或版本清理）。
