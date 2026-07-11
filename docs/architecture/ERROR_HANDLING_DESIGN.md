# 错误处理与容错设计 ERROR_HANDLING_DESIGN

> 日期：2026-07-07
> 目标：统一错误分类、面向用户的友好提示、归档前预检查、批量失败隔离、日志与报告。

## 1. 统一错误分类

| 组件 | 职责 |
|------|------|
| `Core/Common/ErrorSeverity.vb` | 级别：Info / Warning / Error / Critical |
| `Core/Common/AppErrorCode.vb` | 错误码枚举（覆盖 27+ 类，见下） |
| `Core/Common/AppError.vb` | 结构化错误：码 + 级别 + 用户说明 + 建议 + 仅日志详情 |
| `Core/Common/UserFriendlyMessageProvider.vb` | 码 → 友好文案/建议的**单一来源**；OCR 失败归类 |

**原则**：内部日志保留详细错误（含堆栈，经 `ExceptionFormatter`）；用户弹窗只显示 `UserMessage`+`Suggestion`；**不显示英文堆栈/原始异常**；**绝不含 AK/SK/token**。

### 错误码清单（AppErrorCode）
未选择邮件 / 选中项不是邮件 / 邮件无 PDF / 附件保存失败 / PDF 合并失败 / PDF 文本抽取失败 /
本地字段不足 / OCR 未启用 / OCR 配置缺失 / OCR 密钥错误 / OCR 未授权或额度不足 / OCR 网络异常 /
OCR 超时 / OCR 空字段 / OCR 返回错误码 / 归档目录未配置 / 归档目录不存在 / 归档目录无写权限 /
归档路径非法 / 命名模板为空 / 命名未知变量 / 文件名生成失败 / 同名冲突 / 文件复制失败 /
DPAPI 解密失败 / 配置保存失败 / 日志写入失败 / 临时目录不可写 / 日志目录不可写 / 已有任务运行 / 未知异常。

### OCR 失败归类
`ClassifyOcrFailure(message, httpStatus, baiduErrorCode)`：401/认证→密钥错误；403/quota→未授权或额度不足；timeout→超时；网络关键字→网络异常；否则→返回错误码。**不含敏感信息**。

## 2. 配置校验（SettingsForm）

**归档目录**：路径非法→拦截；不存在→询问是否创建；无写权限→提示更换/检查。
**在线 OCR**：启用但缺 AK/SK→拦截（友好文案）；接口/Token 地址非法 http(s)→拦截；超时须 3000–120000ms；最大页数须正整数；AK/SK 自动 Trim；Secret Key 保存前 DPAPI 加密；保存失败→友好原因；成功→简洁提示。
**命名模板**：空→用默认；含未知变量→允许保存但提示（不替换）；变量说明单一来源 `ArchiveNamingRule.SupportedPlaceholders()`。
**测试 OCR 配置** 按钮：仅取 Token（轻量连通性），不上传 PDF、不显示 token；成功“配置可用”，失败按 `ClassifyOcrFailure` 给具体原因。

## 3. 归档运行锁（ArchiveRunGuard）与归档前预检查（ArchivePreflightChecker）

**运行状态与预检查分离（重要，修复 2026-07-10 自我阻断缺陷，见 [../development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md](../development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md)）：**

- **是否已有归档任务在运行**由 `Core/Workflow/ArchiveRunGuard.vb` **单一来源**管理：`TryAcquire()` 用 `Interlocked.CompareExchange` 原子获取，成功返回 `ArchiveRunToken(IDisposable)`；以 `Using`/`Try...Finally` 保证任何路径都释放；纯内存、进程级、无持久化，Outlook 重启不残留。
- **预检查 `ArchivePreflightChecker.Check` 不再检查运行状态**，只做静态检查，避免与外层运行锁自我阻断。

点击“归档”后由 `MainRibbon` 先 `ArchiveRunGuard.TryAcquire`，成功后在 `Using` 内执行预检查，**一次性列出所有问题**：
1. 是否选中邮件。2. 归档目录（未配置/路径非法/不存在→尝试创建/无写权限）。3. 临时目录可写。4. 日志目录可写（不阻断）。5. 命名模板为空（不阻断）。6. 启用 OCR 时配置完整（不阻断，本地可能足够）。

- **有阻断项**（Critical/未选中/临时目录不可写）→ 不进入批量处理（不建进度窗口），展示合并问题清单 + 日志位置，`Using` 释放运行锁。
- **获取运行锁失败**（真正并发）→ 提示“已有归档任务正在运行”，不进入预检查/业务流程。
- 结果写日志，且**逐条**记录每个 issue 的 `Code/Severity/Blocking/Message`。预检查可离线测试（入参为配置值，不依赖 Outlook / My.Settings）。

## 4. 批量失败隔离（BatchArchiveWorkflow，分流后）

- **分流**：滴滴邮件走合并；常规发票逐 PDF 单独归档；未识别 PDF 按 `未识别_原名` 归档；无 PDF 跳过（见 [../requirements/MAIL_PDF_CLASSIFICATION_SPEC.md](../requirements/MAIL_PDF_CLASSIFICATION_SPEC.md)）。
- **单个 PDF 失败不影响同邮件其它 PDF**（常规/未识别逐 PDF 各自 Try/Catch）；未识别 PDF 也能归档不失败。
- 每封邮件独立 `Try/Catch`；合并失败→标记失败、`Continue` 下一封；归档失败→保留临时路径。
- 分类失败/异常记日志；无法判断的 PDF 归 UnknownPdf（不强行归常规）。
- OCR 失败但本地有字段→部分成功；本地不足且 OCR 不可用→NeedsOcr（不生成明显错误文件名，走 fallback 命名）。
- 每封邮件状态：成功 / 部分成功 / 跳过 / 失败 + 用户可读原因（`BuildItemReason`）。
- 汇总（以邮件为单位）：总/成功/部分/需 OCR/跳过/失败/合并数 + 报告路径 + 日志位置。

## 5. 进度窗口（ProgressForm）

总邮件数 / 已处理 / 当前序号+主题 / 当前阶段（面向用户文案）/ 百分比；失败后显示“上一封邮件处理失败，已继续下一封”；完成显示“已完成”；不逐附件弹窗；`Application.DoEvents` 关键阶段刷新避免假死；**无取消按钮、禁止关闭**（`ControlBox=False`），关闭窗口不会崩溃。

## 6. 日志与报告

- 每次归档生成批次 ID（`B`+时间戳），日志记录批次 ID。
- 每封邮件日志：序号、主题、PDF 数、原始附件名、合并临时路径、最终归档路径、识别来源、命名规则、fallback、缺失字段、状态、原因。
- 日志**不含 AK/SK/token**（仅脱敏摘要）；日志写入失败不影响主流程。
- 归档报告：`{日志目录}\archive-report-yyyyMMdd-HHmmss.txt`，用户可读，不含敏感密钥；路径在汇总弹窗展示。
- 日志/报告位置：`{归档目录}\logs\` 或 `%AppData%\iWorkHelper\logs\`。

## 7. 已验证 / 待验证

- **已验证（离线，真实执行）**：错误码全覆盖有友好文案、OCR 失败归类、预检查阻断逻辑、运行锁 `ArchiveRunGuard`（原子获取/防重入/异常释放）、`--simulate-error`/`--preflight`、selftest 65/65、命名/OCR/合并回归、双工程构建。
- **待 Outlook 验证**：设置页各校验弹窗、测试 OCR 按钮、进度窗口显示、真实批量失败隔离、汇总弹窗与报告展示。
