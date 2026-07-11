# 归档运行状态缺陷分析与修复报告 ARCHIVE_RUNNING_STATE_BUG_REPORT

> 日期：2026-07-10
> 主题：点击「归档」立即弹出「已有归档任务正在运行」的阻断错误——根因、修复与验证。
> 范围：仅归档前预检查 / 运行状态锁 / 启动顺序 / 状态释放 / 防重复点击。不涉及本地识别、在线 OCR、PDF 合并、邮件分类、命名模板业务逻辑。

---

## 1. 现象

用户在 Outlook 中选中邮件、点击工作助手「归档」按钮，**每一次**都立即弹窗：

```
无法开始归档，请先解决以下问题：

1. 已有归档任务正在运行。（请等待当前任务完成后再试。）
```

日志中只有一行、无法定位：

```
2026-07-10 11:53:11.027 [INFO ] 归档预检查：问题数=1，阻断=True
```

用户已尝试：清理缓存、正式部署（Release）、非 F5 调试，问题依旧。

---

## 2. 根因（确认）

这是一个**确定性自我阻断**缺陷，与运行环境无关，因此清缓存/正式部署/换调试方式都不可能解决。

修复前的调用链（`MainRibbon.ButtonArchive_Click`）：

1. `ArchivePreflightChecker.TryBeginRun()` —— 先把全局 `Shared _running` 置为 `True`（表示“归档进行中”）。
2. 紧接着 `checker.Check(...)` 执行归档前预检查。
3. 预检查内部有一段：

   ```vb
   ' 已有任务在运行
   If IsRunning Then
       r.AddCode(AppErrorCode.ArchiveAlreadyRunning, blocking:=True)
   End If
   ```

   此时 `IsRunning` 读到的正是**上一步自己刚刚置位的 `_running=True`**，于是把“本次点击”误判为“已有任务在运行”，生成唯一的阻断问题（正好对应日志“问题数=1，阻断=True”）。

**结论：预检查前先设置了运行中状态，预检查又检查运行中状态 → 每次点击都自我阻断。** 即用户高概率原因判断中的第 1 条 + 第 2 条。

### 弹窗文案来源

- 文案「已有归档任务正在运行」由 `Core/Common/UserFriendlyMessageProvider.Describe(AppErrorCode.ArchiveAlreadyRunning)` 唯一提供。
- 作为“预检查问题”被 `ArchivePreflightChecker.Check` 通过 `r.AddCode(AppErrorCode.ArchiveAlreadyRunning, blocking:=True)` 加入，再由 `ArchivePreflightResult.BuildUserText()` 拼成弹窗正文，`MainRibbon` 弹出。

### 排查结论（对应用户提问）

1. 弹窗“已有归档任务正在运行”由 **`ArchivePreflightChecker.Check`** 生成问题项，文案由 **`UserFriendlyMessageProvider.Describe`** 提供。
2. 归档前预检查由 **`MainRibbon.ButtonArchive_Click`** 调用。
3. **是**，预检查前 `TryBeginRun()` 已把 `_running` 置位——这是自我阻断的直接原因。
4. **否**，全项目只有一处运行标志：`ArchivePreflightChecker._running`（`Shared`）。`BatchArchiveWorkflow` / `ProgressForm` / `UiArchiveProgressReporter` / `ThisAddIn` 均**不**维护运行状态（已逐一核对）。
5. **是**，`_running` 是 `Private Shared`（进程级静态）。
6. **无**锁文件、无配置项、无注册表持久化运行状态——纯内存。
7. 原实现用 `Try/Finally` + `EndRun()` 释放，异常路径本身能释放；但因“预检查即阻断”，`EndRun` 每次都在 `Finally` 正常执行，`_running` 不会残留——**这不是残留问题，而是顺序问题**。
8. 存在重复点击并发风险（原 `SyncLock` 判断可防并发），但本缺陷与重复点击无关：第一次点击就被误判。
9. `ProgressForm` 与运行状态无关（它从不设置/持有运行标志），也不是本缺陷来源。

> 补充：因为运行状态是**进程内存**且 Ribbon 加载时还额外 `ResetState()`，Outlook 重启后绝不会残留旧状态。所以本问题**不是**跨重启残留，而是单次点击内的逻辑顺序错误。

---

## 3. 为什么清缓存 / 正式部署 / 非 F5 都没用

因为这是**源码逻辑顺序错误**，在任何构建配置、任何部署方式、任何缓存状态下都会 100% 复现——它不依赖任何外部状态或残留。唯一的解决办法是改代码顺序 / 拆分状态职责。

---

## 4. 修复方案（采用 Plan A：运行锁单一来源）

将“归档是否正在运行”的判断与设置**统一收敛到一个新的单一来源** `ArchiveRunGuard`，预检查**不再**检查运行状态，从根本上消除自我阻断。

### 新增

| 文件 | 职责 |
|------|------|
| `Core/Workflow/ArchiveRunGuard.vb` | 唯一运行锁。`Interlocked.CompareExchange` 原子获取（0→1）；成功返回 `ArchiveRunToken`；`IsRunning` 原子读；`DescribeHolder()` 输出批次/线程/持有时长（无敏感信息）；`Release()` 仅供 token 调用。纯内存、进程级。 |
| `Core/Workflow/ArchiveRunToken.vb` | 运行锁令牌，`IDisposable`。`Dispose` 幂等，释放全局运行锁。必须以 `Using`/`Try...Finally` 包裹。 |

### 改动

| 文件 | 改动 |
|------|------|
| `Core/Workflow/ArchivePreflightChecker.vb` | 删除 `_running` / `TryBeginRun` / `EndRun` / `ResetState` / `IsRunning` 及 `If IsRunning Then AddCode(ArchiveAlreadyRunning)` 的自我阻断分支；预检查不再管理运行状态。补充**逐条 issue 明细日志**（Code/Severity/Blocking/Message）。将依赖 `My.Settings` 的明文密钥迁移移出预检查（改到 Ribbon 侧执行），恢复“预检查纯配置入参、可离线测试”的设计。 |
| `MainRibbon.vb` | 重排启动顺序：**先** `ArchiveRunGuard.TryAcquire` → 失败则提示并返回 → 成功进入 `Using runToken` → 密钥迁移 → 预检查 → 阻断则显示真实原因并返回（不建进度窗口）→ 通过才建进度窗口 → 批量归档 → 汇总。运行锁在 `Using`/`Finally` 释放。补充点击、线程、获取锁、预检查开始/通过/失败、归档开始/结束、释放锁的流程日志。删除 Ribbon 加载时的 `ResetState()`（内存锁不需要）。 |
| `tools/OfflineTester/Program.vb` | 新增 `[7] ArchiveRunGuard` 自测（11 项）。 |
| `iWorkhelper.vbproj` / `OfflineTester.vbproj` | 注册两个新文件。 |

---

## 5. 修复前 / 修复后流程对比

**修复前（错误）：**

```
点击 → TryBeginRun(_running=True) → 预检查(检查 IsRunning=True → 阻断!) → 弹“已有任务运行” → Finally EndRun
                                     ▲ 每次点击都在这里被自己挡住
```

**修复后（正确，Plan A）：**

```
点击
 └─ ArchiveRunGuard.TryAcquire()  ── 失败 ─→ 提示“已有归档任务正在运行”，返回（真正并发时才触发）
        │成功(token)
        └─ Using token
             ├─ 密钥迁移（不阻断）
             ├─ 预检查（只查配置/目录/选中邮件/权限，不查运行状态）
             │     └─ 有阻断 → 显示真实原因，返回（不建进度窗口）
             ├─ 建进度窗口 → BatchArchiveWorkflow.Run → 汇总
             └─ End Using / Finally → 释放运行锁（预检查失败/业务异常/进度窗口异常均释放）
```

---

## 6. 运行锁设计与异常释放策略

- **单一来源**：全项目仅 `ArchiveRunGuard` 一处判断/设置运行状态。
- **原子获取**：`Interlocked.CompareExchange(_state, 1, 0)`，重复点击/并发只有第一次成功。
- **token 释放**：获取成功返回 `ArchiveRunToken(IDisposable)`；`Using` 保证预检查失败早退、业务异常、进度窗口异常关闭等**任何路径**都释放；`Dispose` 幂等。
- **无持久化**：不用锁文件/配置/注册表；Outlook 进程重启后静态字段自动为“未运行”，不存在跨重启残留。
- **防重复点击**：`Using` 打开期间第二次点击 `TryAcquire` 返回 `Nothing`，友好提示且**不进入**预检查或业务流程；第一次点击不再被误判。

---

## 7. 如何验证

- 离线（已执行）：
  - `OfflineTester --selftest` → 65/65 通过，含 `[7] ArchiveRunGuard` 11 项（首次可获取、二次失败、释放后可再获取、幂等 Dispose、预检查失败释放、业务异常释放、无残留）。
  - `OfflineTester --preflight <有效目录>` → “问题数=0，是否阻断=False”（不再出现自我阻断）。
  - 主工程 Debug/Release 均 Rebuild 通过。
- Outlook 内（待用户验证，见 [TROUBLESHOOTING](../deployment/TROUBLESHOOTING.md) 与 [TEST_PLAN](../testing/TEST_PLAN.md)）：
  1. 选中邮件点归档 → 不再立即提示“已有任务运行”。
  2. 人为制造预检查失败（如清空归档目录设置）→ 提示真实原因，再次点击可用（锁已释放）。
  3. 快速连点归档 → 只允许一个任务；第二次友好提示。

---

## 8. 若再次出现，应查看哪些日志

日志：`{归档目录}\logs\yyyy-MM-dd.log` 或 `%AppData%\iWorkHelper\logs\yyyy-MM-dd.log`。

关注新增流程日志：

```
用户点击『归档』按钮。线程=NN
已获取归档运行锁。                              ← 若无此行而直接“获取失败”，说明确有并发任务
获取归档运行锁失败：已有归档任务正在运行。持有者：批次=..., 持有线程=..., 已持有=..ms
归档前预检查开始。选中邮件数=N
归档预检查完成：问题数=X，阻断=Y
  预检查[阻断]：Code=..., Severity=..., Blocking=True, Message=...   ← 定位到底是哪一项
归档前预检查通过，开始批量归档。
批量归档流程结束。批次=...
归档运行锁已释放（线程=NN）。
```

- 若看到“获取归档运行锁失败”，用“持有者”信息判断是否真有并发（正常单人使用不应出现）。
- 若看到阻断，直接看 `预检查[阻断]` 那行的 `Code/Message` 即真实原因，**不会**再是被自我阻断的 `ArchiveAlreadyRunning`。
