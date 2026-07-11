# 风险登记 RISK_REGISTER

> 生成日期：2026-07-06
> 分级：影响 × 可能性 → 高/中/低。所有风险均给出应对措施。

## 一、技术风险

| ID | 风险 | 依据 | 等级 | 应对措施 |
|----|------|------|------|---------|
| R-TECH-1 | **PdfPig 引用的是 `net471` DLL，工程为 net48**，可能存在运行期加载/兼容问题 | `iWorkhelper.vbproj:142` HintPath 指向 `lib\net471`；`packages.config` 标 `net48` | 中 | 在 MVP 首环（NEXT_ACTIONS N4）先跑通实际读取；如异常，改用 PdfPig 的 net48/netstandard2.0 DLL 或升级 PdfPig 版本 |
| R-TECH-2 | **PdfPig 不是 OCR**，若发票为扫描/图片型 PDF 将抽不到文本，核心链路走不通 | PdfPig 仅抽取内嵌文本层；`Core\*OCR\` 为空 | 高 | 尽早关闭需求 Q1；若为图片型，按路线图阶段 3 引入真 OCR（本地或在线） |
| R-TECH-3 | **在线 OCR 方案完全未定**（服务/Key/费用/合规均空白） | `ParseMode=Online` 无任何实现或配置 | 中 | 关闭 Q2 后再评估；未定前 Online 分支保持禁用或提示“未配置” |
| R-TECH-4 | **`Option Strict Off`**，类型宽松，业务代码易埋运行期类型错误 | `iWorkhelper.vbproj:318` | 中 | 新增 `Core/` 业务文件逐文件加 `Option Strict On`；关键转换显式处理 |
| R-TECH-5 | **无日志/无全局异常处理**，出问题难定位、易在 Outlook 内静默失败或崩溃 | 全仓库无日志框架；仅零星 Try/Finally | 高 | 优先落地日志与异常封装（路线图阶段 1） |
| R-TECH-6 | **Outlook 对象模型陷阱**：选中多项、非邮件项、COM 对象未释放导致内存泄漏/Outlook 卡死 | 尚无附件读取代码，实现时易踩坑 | 中 | 实现 `MailAttachmentReader` 时严格判类型、及时 `Marshal.ReleaseComObject`、避免长链式 COM 调用 |
| R-TECH-7 | **耗时操作阻塞 UI**（大 PDF、在线识别）导致 Outlook 假死 | 归档为同步入口 | 中 | 处理异步化/后台线程；注意 Outlook 单线程套间(STA)约束 |

## 二、业务风险

| ID | 风险 | 依据 | 等级 | 应对措施 |
|----|------|------|------|---------|
| R-BIZ-1 | **提取字段、命名规范、分类规则均未确认**，先做会返工 | 需求 Q3/Q4/Q5 待确认 | 高 | 编码前关闭 Q3/Q4/Q5；先做与规则无关的取件/抽文本 |
| R-BIZ-2 | **缺少真实样例 PDF**，解析逻辑无法验证正确性 | 仓库无样例文件 | 高 | 向需求方索取 2~3 份脱敏样例 |
| R-BIZ-3 | **归档去重/覆盖策略未定**，可能重复归档或覆盖已有文件 | 需求 Q9 未确认；无写盘逻辑 | 中 | 关闭 Q9；实现时默认“同名加序号，不覆盖” |
| R-BIZ-4 | **仅面向滴滴**，其它发票来源无法处理，扩展性未规划 | 背景聚焦滴滴；无抽象 | 低 | 提取逻辑接口化，为未来来源预留 |

## 三、兼容性风险

| ID | 风险 | 依据 | 等级 | 应对措施 |
|----|------|------|------|---------|
| R-COMP-1 | **文档框架说明与实际不符**（“.NET Framework 8.0” vs 实际 4.8），易误导环境搭建 | `.trae/rules` vs `iWorkhelper.vbproj:29` | 中 | 关闭 Q8 后更正 `.trae` 说明；以工程 4.8 为准 |
| R-COMP-2 | **Office/Outlook 版本差异**：Interop 15.0 与目标 Office 365 2502 的功能/对象差异 | `iWorkhelper.vbproj:191-220`（Office 15.0）；背景 Office 2502 | 中 | 在目标 Outlook 版本实测；使用 `EmbedInteropTypes` 已缓解版本绑定 |
| R-COMP-3 | **AnyCPU + Office 位数**：32/64 位 Outlook 环境差异 | `Platform=AnyCPU` | 低 | AnyCPU 一般兼容；如遇原生依赖(OCR)再评估位数 |

## 四、维护风险

| ID | 风险 | 依据 | 等级 | 应对措施 |
|----|------|------|------|---------|
| R-MAINT-1 | **非 Git 仓库**，无版本历史、无法回溯、协作困难 | `git status` 报 not a repository | 高 | 立即 `git init`（NEXT_ACTIONS N1） |
| R-MAINT-2 | **临时签名证书入库风险**：`iWorkhelper_TemporaryKey.pfx` 含私钥 | 根目录存在该 pfx | 中 | `.gitignore` 忽略私钥；发布用正式证书并妥善保管 |
| R-MAINT-3 | **构建产物在库**（bin/obj/.vs）污染，若入 Git 更严重 | 根目录存在这些目录 | 低 | `.gitignore` 忽略 |
| R-MAINT-4 | **packages.config 旧式依赖模式**，NuGet 还原/升级不如 PackageReference 便捷 | `packages.config` 模式 | 低 | 暂维持；未来可评估迁移到 PackageReference |

## 五、文档缺失风险

| ID | 风险 | 依据 | 等级 | 应对措施 |
|----|------|------|------|---------|
| R-DOC-1 | **原本几乎无文档**（仅 4 行环境说明），需求/设计/命名规则无据可依 | 仓库仅 `.trae/rules/projectIntroduction.md` | 中 | 本次已补齐 `/docs`；需求方补充需求细节后持续更新 |
| R-DOC-2 | **无变更记录/无 README**，交接与协作靠口头 | 无相关文件 | 中 | 建 Git 后以提交信息 + 本 `/docs` 维护；补根级 README（可选） |
| R-DOC-3 | **关键决策（OCR 选型、命名规范）未落文档**，易反复讨论 | Q1~Q10 未确认 | 中 | 确认后写入 `REQUIREMENTS_BACKLOG.md`，形成决策记录 |

## 六、Top 风险速览（需优先处理）

1. **R-TECH-2 / R-BIZ-1 / R-BIZ-2**：是否需真 OCR + 提取规则 + 样例——决定 MVP 可行性，**必须先关闭需求 Q1/Q3/Q4 并取样例**。
2. **R-TECH-5**：无日志/异常——**阶段 1 立即建设**。
3. **R-MAINT-1**：非 Git——**立即 `git init`**。
4. **R-TECH-1**：PdfPig net471/net48 兼容——**MVP 首环即验证**。
