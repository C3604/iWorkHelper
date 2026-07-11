# 开发路线 DEVELOPMENT_ROADMAP

> 生成日期：2026-07-06
> 依据：当前实际进度（见 [../project/CURRENT_PROGRESS.md](../project/CURRENT_PROGRESS.md)）。

## 1. 当前阶段判断

**阶段 0：脚手架阶段（已完成 UI/设置层，核心业务为零）。**

- 已具备：可加载的 VSTO 工程、Ribbon UI、可用的设置界面与配置持久化。
- 缺失：附件读取、PDF/OCR 解析、字段提取、重命名、归档——**端到端链路完全没有**。
- 阻塞：多项关键需求待确认（Q1~Q8，见需求清单），其中 **Q1（是否需要真 OCR）、Q3（提取字段）、Q4（命名规范）** 最关键。

## 2. 下一阶段优先级

**先打通「最小可用链路」（MVP）：当前选中邮件 → 取 PDF → PdfPig 抽文本 → 提取字段 → 重命名 → 归档。**
优先 **Local + 文本型 PDF** 路径（不引入 OCR、不引入云服务），因为它无需外部依赖、可立即验证、能暴露大部分架构问题。Online/真 OCR 待 Q1/Q2 确认后再做。

优先级顺序：**先纳管版本 → 搭基础设施（日志/异常）→ 打通 Local MVP → 再扩展 OCR/Online → 再做批量/去重/回写。**

## 3. 阶段划分与任务拆分

> 标注：【必须做】立即价值高且无阻塞；【建议做】提升质量；【待确认后再做】依赖需求澄清。

### 阶段 1：工程治理与基础设施【必须做，无需求阻塞】
| 任务 | 目标 | 主要文件/模块 | 验收标准 | 风险 |
|------|------|--------------|---------|------|
| T1.1 `git init` 并首次提交 | 纳入版本管理 | 仓库根（新增 `.gitignore`） | 有初始提交；`bin/obj/.vs/packages` 被忽略 | 低（注意勿提交 pfx 私钥） |
| T1.2 建立日志模块 | 全局可记录处理过程 | `Core/Infrastructure/Logger.vb`（新增） | 能写入文件日志，含时间/级别/消息 | 低 |
| T1.3 建立异常/提示封装 | 统一错误处理与用户提示 | `Core/Infrastructure/`（新增） | 业务异常有统一 Catch 与提示 | 低 |

### 阶段 2：Local 最小链路 MVP【必须做；T2.3 起需 Q3/Q4】
| 任务 | 目标 | 主要文件/模块 | 验收标准 | 风险 |
|------|------|--------------|---------|------|
| T2.1 读取当前邮件 PDF 附件 | 从选中 MailItem 保存 PDF 到临时目录 | `Core/MailAttachmentReader.vb`（新增），`ThisAddIn`（取 Application/Explorer.Selection） | 选中含 PDF 的邮件能得到临时文件路径；无附件/非 PDF 有提示 | Outlook 对象模型/多选处理 |
| T2.2 PdfPig 抽取文本 | 读出文本型 PDF 文本 | `Core/Local OCR/PdfPigTextParser.vb`（新增） | 文本型 PDF 返回非空文本；加密/图片型 PDF 返回明确“无文本”信号 | net471 DLL 在 net48 下需验证；图片型无文本（触发 Q1） |
| T2.3 字段提取（滴滴） | 从文本解析约定字段 | `Core/Extraction/DidiInvoiceExtractor.vb`（新增） | 对样例 PDF 正确提取字段（依 Q3） | **依赖 Q3；需样例 PDF** |
| T2.4 文件命名 | 依字段生成规范文件名 | `Core/Naming/FileNamingService.vb`（新增） | 生成名符合规范、去非法字符、冲突加序号（依 Q4） | **依赖 Q4** |
| T2.5 归档落盘 | 复制/移动到 ArchiveFolderPath | `Core/Archiving/ArchiveService.vb`（新增） | 文件落到设置目录；目录缺失自动建；无权限有提示 | 权限/占用/覆盖策略 |
| T2.6 串联入口 | 用真实流程替换归档占位 | `MainRibbon.vb:9 ButtonArchive_Click` | 点击“归档”跑通 T2.1→T2.5，弹出结果汇总 | 整合与错误传播 |

### 阶段 3：解析能力扩展【待确认后再做】
| 任务 | 目标 | 依赖 |
|------|------|------|
| T3.1 本地 OCR（图片型 PDF） | 支持扫描件识别 | Q1=图片型 → 选型（Tesseract/Windows.Media.Ocr/第三方）后实现 `Core/Local OCR/` |
| T3.2 在线 OCR/发票识别 | 实现 Online 分支 | Q2（服务、Key、预算）确认后实现 `Core/Online OCR/` |
| T3.3 解析接口抽象 | Local/Online 可切换 | 定义 `IDocumentParser`，`ParseMode` 运行时选实现 |

### 阶段 4：体验与稳健性【建议做】
| 任务 | 目标 | 主要文件 |
|------|------|---------|
| T4.1 处理结果报告 UI | 替代 MsgBox，展示成功/失败/跳过 | 新增结果窗体或状态提示 |
| T4.2 异步/后台处理 | 避免冻结 Outlook | 处理流程异步化 |
| T4.3 批量处理 | 多封邮件/文件夹（依 Q6） | 遍历 Selection/Folder |
| T4.4 去重 | 避免重复归档（依 Q9） | 已处理记录清单 |

### 阶段 5：发布【发布前必须做】
| 任务 | 目标 |
|------|------|
| T5.1 正式代码签名证书 | 替换 `iWorkhelper_TemporaryKey.pfx` |
| T5.2 部署方式 | ClickOnce / 安装包 / 组策略分发（待定） |
| T5.3 更正框架说明 | 更新 `.trae` 规则中“.NET Framework 8.0” → 实际 4.8（Q8 确认后） |

## 4. 推荐开发顺序（关键路径）

```
T1.1 (git)  →  T1.2/T1.3 (基础设施)
      │
      ▼
T2.1 (附件读取)  →  T2.2 (PdfPig 抽文本)
      │                    │
      │        ┌───────────┴──── 【关闭 Q1】若图片型 → 阶段3 T3.1 提前
      ▼        ▼
【关闭 Q3/Q4】T2.3 (字段提取) → T2.4 (命名) → T2.5 (归档)
      │
      ▼
T2.6 (串联，替换占位)  ← MVP 完成，可演示端到端
      │
      ▼
阶段3（OCR/Online，依 Q1/Q2） → 阶段4（体验/批量） → 阶段5（发布）
```

**里程碑：**
- **M1**：阶段 1 完成——工程受控、有日志/异常骨架。
- **M2（关键）**：T2.6 完成——Local 文本型 PDF 端到端可用，可对真实滴滴邮件演示。
- **M3**：OCR/Online 就绪，双模式可用。
- **M4**：批量/去重/异步，具备生产可用性。
- **M5**：签名与分发，正式发布。

## 5. 开工前置条件（务必先关闭）

- **Q1**（是否图片型 PDF）、**Q3**（字段清单）、**Q4**（命名规范）——直接决定 T2.2~T2.4，**必须先确认**。
- 需求方提供 **2~3 个真实滴滴发票/行程单 PDF 样例**（脱敏），否则 T2.3 无法验证。
- 详见 [../requirements/REQUIREMENTS_BACKLOG.md](../requirements/REQUIREMENTS_BACKLOG.md) 与 [NEXT_ACTIONS.md](NEXT_ACTIONS.md)。
