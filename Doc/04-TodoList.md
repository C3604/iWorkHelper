# iWorkHelper 开发任务 TodoList（按优先级）

> 文档版本：v0.1.0  ·  更新日期：2025-11-06  ·  适用范围：Outlook VSTO 加载项（VB / .NET Framework 4.8）

## 0. 相关文档
- 开发框架与技术架构设计：[`Doc/01-Architecture.md`](01-Architecture.md)
- 技术栈与兼容性：[`Doc/02-TechStack.md`](02-TechStack.md)
- 开发流程与规范：[`Doc/03-DevelopmentProcess.md`](03-DevelopmentProcess.md)
- 开发注意事项与常见问题：[`Doc/05-Notes.md`](05-Notes.md)
- 变更记录：[`Doc/CHANGELOG.md`](CHANGELOG.md)

说明：本清单聚焦于首次版本（v0.1.x）交付所需任务。每条任务包含描述、预计工时、依赖关系与完成标准。优先级分为 P0（关键）、P1（重要）、P2（增强）。

| 优先级 | 任务名称 | 任务描述 | 预计工时 | 依赖关系 | 完成标准 |
|---|---|---|---|---|---|
| P0 | 初始化加载项宿主 | 完成 `ThisAddIn` 生命周期初始化与异常兜底 | 4h | 环境/模板 | Outlook 启动加载成功，无异常日志 |
| P0 | 功能区 UI 基线 | 创建 `MainRibbon`，添加「处理邮件」按钮与进度区 | 6h | 宿主 | Ribbon 可见且按钮可点击，显示进度占位 |
| P0 | 附件提取模块 | 从选中邮件提取 PDF 附件到临时目录 | 8h | 宿主 | 能提取本地 PDF 附件，输出 `AttachmentInfo` 列表 |
| P0 | PdfPig 文本提取 | 实现 `OcrService` 基于 PdfPig 的文本抽取 | 10h | 附件提取 | 文本提取稳定，返回 `DocumentInfo` 与关键字段 |
| P0 | 滴滴文档识别与合并 | 识别发票/行程单并按顺序合并（iTextSharp） | 12h | 文本提取 | 正确识别合并，输出合并后 PDF，错误时降级 |
| P0 | 智能重命名 | 基于日期/金额/姓名/类型生成规范化文件名 | 6h | 文本提取 | 文件名符合约定（示例规则），冲突有策略 |
| P0 | 归档与冲突处理 | 移动到 `ArchivePath`，创建结构化目录并避重名 | 8h | 重命名 | 归档成功，重名自动追加后缀，失败有日志 |
| P0 | 邮件状态标记 | 完成后自动标记分类或旗标，避免重复处理 | 4h | 宿主 | 标记持久化且可视化可见 |
| P1 | 进度可视化 | 实现进度窗体 `ProgressForm` 与 IProgressReporter | 6h | P0 流程 | 关键阶段更新百分比与消息，不卡 UI |
| P1 | 设置界面 | `SettingsForm` 管理 ArchivePath 与选项（MergeDidiFiles） | 8h | 宿主 | 设置可保存到 `My.Settings` 并校验路径 |
| P1 | 日志模块 | `LogManager` 文件日志，支持滚动与等级 | 6h | 宿主 | 关键步骤有日志，错误有堆栈与上下文 ID |
| P1 | 单元/集成测试 | 核心服务测试与端到端流程验证 | 12h | P0/P1 | 覆盖关键分支，端到端成功率 ≥ 95% |
| P2 | OCR 扩展（可选） | 扫描版引入 OCR（如 Tesseract）并抽象接口 | 16h | 文本提取 | 扫描版可识别，性能与准确率达基线 |
| P2 | ClickOnce 发布脚本 | 自动化打包与版本号更新脚本（CMD） | 6h | 发布流程 | 一键生成安装包与发布说明 |

### 文件命名示例规则
- 格式：`YYYYMMDD_滴滴_姓名_金额_类型.pdf`
- 例子：`20250103_滴滴_张三_128.50_发票.pdf`
- 冲突策略：若存在同名文件，追加 `_({n})` 后缀。