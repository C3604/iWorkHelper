# iWorkHelper 开发框架与技术架构设计

> 文档版本：v0.1.0  ·  更新日期：2025-11-06  ·  适用范围：Outlook VSTO 加载项（VB / .NET Framework 4.8）

## 1. 概述
- 目标：打造一款高效、易用的 Outlook 加载项，实现邮件附件中 PDF 的自动化处理，聚焦滴滴发票/行程单识别、重命名与归档。
- 一致性：与项目根目录 `Readme.md` 与 `.trae/rules/project_rules.md` 保持一致（IDE：VS 2022，Outlook：Office 365 2502，框架：.NET 4.8，语言：VB）。
- 核心能力：附件提取、PDF文本识别（PdfPig）、滴滴文档合并（iTextSharp）、智能重命名、自动归档、状态标记与进度显示。

## 0. 相关文档
- 技术栈与兼容性：[`Doc/02-TechStack.md`](02-TechStack.md)
- 开发流程与规范：[`Doc/03-DevelopmentProcess.md`](03-DevelopmentProcess.md)
- 开发任务 TodoList：[`Doc/04-TodoList.md`](04-TodoList.md)
- 开发注意事项与常见问题：[`Doc/05-Notes.md`](05-Notes.md)
- 变更记录：[`Doc/CHANGELOG.md`](CHANGELOG.md)

## 2. 技术架构
```
┌──────────────────────────────────────────────────────────────────────┐
│                   Outlook (Host) / Office 365 2502                  │
└───────────────▲──────────────────────────────────────────────────────┘
                │ VSTO 加载项生命周期事件（ThisAddIn_Startup/Shutdown）
┌───────────────┴──────────────────────────────────────────────────────┐
│                          VSTO Add-in（VB/.NET 4.8）                  │
│  • ThisAddIn：宿主入口，注册事件、初始化服务容器                       │
│  • MainRibbon：功能区 UI，触发命令                                   │
└───────────────▲──────────────────────────────────────────────────────┘
                │ 命令调用（按钮点击/上下文菜单）
┌───────────────┴──────────────────────────────────────────────────────┐
│                         业务服务层（可测试）                          │
│  • MailProcessor（编排）：附件提取→OCR→合并→重命名→归档→标记          │
│  • OcrService（PdfpigOcr）：PDF 文本提取（必要时扩展 OCR）             │
│  • DidiMergeService：滴滴发票/行程单识别与合并（iTextSharp）          │
│  • RenameService：基于内容的规范化文件命名                            │
│  • ArchiveService：目标目录移动、冲突处理、结构化归档                 │
└───────────────▲──────────────────────────────────────────────────────┘
                │ 依赖注入/接口调用
┌───────────────┴──────────────────────────────────────────────────────┐
│                         基础设施与集成层                             │
│  • SettingsManager（My.Settings 封装）                               │
│  • LogManager（文件日志/事件日志）                                   │
│  • PdfPig（UglyToad.PdfPig）                                         │
│  • iTextSharp（PDF 合并）                                            │
└──────────────────────────────────────────────────────────────────────┘
```

### 2.1 模块划分与职责边界
- ThisAddIn（宿主层）
  - 职责：生命周期管理（启动/关闭）、加载 UI、初始化服务与配置、异常兜底。
  - 边界：不实施业务处理，不操作 UI 控件之外的业务逻辑。
- MainRibbon（UI 层）
  - 职责：提供按钮与进度展示，收集用户参数，调用编排服务。
  - 边界：不直接处理文件与 PDF；仅通过接口调用服务层。
- MailProcessor（编排层）
  - 职责：串联完整流程：附件→文本→识别→合并→重命名→归档→标记。
  - 边界：不实现具体 OCR/合并/归档细节；依赖对应服务接口。
- OcrService（PdfpigOcr）
  - 职责：基于 PdfPig 提取文本；必要时可扩展为对扫描件的 OCR。
  - 边界：不保存文件；仅返回结构化文本与键值对。
- DidiMergeService
  - 职责：识别滴滴发票与行程单配对逻辑，按顺序合并为一个 PDF。
  - 边界：不负责重命名和归档；只输出合并结果路径或字节流。
- RenameService
  - 职责：根据提取到的日期/金额/姓名/类型生成规范化文件名。
  - 边界：不移动文件；仅返回建议文件名与冲突解决策略。
- ArchiveService
  - 职责：将文件移动到归档路径，建立分目录（年/月/类型），处理重名冲突。
  - 边界：不修改文件内容；只进行文件系统操作。
- SettingsManager
  - 职责：统一读取/写入 `My.Settings`，提供默认值与校验。
  - 边界：不含 UI；对外暴露只读/可写属性与验证方法。
- LogManager
  - 职责：按等级记录日志（Info/Warn/Error），保留操作上下文 ID。
  - 边界：不抛出异常；失败时应尽量降级。

### 2.2 模块间交互规范
- 依赖方向：UI → 编排（MailProcessor） → 具体服务（Ocr/Merge/Rename/Archive）。服务间不互相调用，通过编排统一协调。
- 通信约定：
  - 输入：基于 `AttachmentInfo` 与 `DocumentInfo` 的数据契约。
  - 输出：统一使用 `OperationResult`（含成功/失败、消息、异常、上下文 ID）。
- 线程与异步：
  - UI 线程仅触发命令与更新进度；业务处理在后台任务中执行。
  - 使用 `Task`/`Async` 模式；UI 更新通过 `IProgressReporter` 回调到主线程。
- 错误处理：
  - 服务层自行捕获异常并返回失败结果；编排层统一汇总与告警。
  - 必须保证偶发失败不影响其他附件处理（隔离策略）。

## 3. 数据契约（示例）
```vb
' 附件元信息
Public Class AttachmentInfo
    Public Property MailEntryId As String
    Public Property FileName As String
    Public Property TempPath As String
    Public Property SizeInBytes As Long
    Public Property CreatedAt As DateTime
End Class

' 文档抽取结果
Public Class DocumentInfo
    Public Property SourcePath As String
    Public Property Text As String
    Public Property KeyValues As Dictionary(Of String, String) ' 如: 日期、金额、订单号
    Public Property Vendor As String ' 如: 滴滴出行
    Public Property DocType As String ' 发票/行程单
End Class

' 统一返回结果
Public Class OperationResult
    Public Property Success As Boolean
    Public Property Message As String
    Public Property Exception As Exception
    Public Property ContextId As Guid
End Class
```

## 4. 接口定义（示例）
```vb
Public Interface IMailProcessor
    Function ProcessSelectedItems(progress As IProgressReporter) As Task(Of OperationResult)
End Interface

Public Interface IOcrService
    Function ExtractText(pdfPath As String) As Task(Of DocumentInfo)
End Interface

Public Interface IDidiMergeService
    Function TryMerge(invoice As DocumentInfo, itinerary As DocumentInfo) As Task(Of String) ' 返回合并后文件路径
End Interface

Public Interface IRenameService
    Function SuggestFileName(info As DocumentInfo) As String
End Interface

Public Interface IArchiveService
    Function MoveToArchive(sourcePath As String, suggestedName As String) As OperationResult
End Interface

Public Interface IProgressReporter
    Sub Report(stepName As String, percent As Integer, message As String)
End Interface
```

## 5. 业务流程（用户交互 → 自动处理）
1. 用户在 Outlook 选择邮件，点击功能区「处理邮件」。
2. MailProcessor 读取并保存附件到临时目录（逐个处理）。
3. OcrService（PdfPig）提取文本，输出 `DocumentInfo` 与关键字段。
4. DidiMergeService 识别发票/行程单并匹配合并；无配对时跳过合并。
5. RenameService 根据内容生成文件名（如：`20250103_滴滴_张三_128.50_发票.pdf`）。
6. ArchiveService 将文件移动到 `ArchivePath` 下的结构化目录，并处理重名冲突。
7. Outlook 邮件标记状态（如分类/旗标），记录日志并更新进度 UI。

## 6. 配置约定
- 配置源：`My.Settings`
- 关键项：
  - `ArchivePath`（必填）：归档根目录（示例：`D:\WorkArchive\Invoices`）。
  - `MergeDidiFiles`（布尔）：是否合并滴滴发票与行程单。
  - `EnableProgressUI`（布尔）：是否显示进度窗体。

> 说明：命令行终端建议与兼容性请参阅《技术栈与兼容性》。

## 7. 日志规范
- 等级：`Info`、`Warning`、`Error`。
- 结构：`时间戳 | 上下文ID | 模块 | 等级 | 消息`。
- 存储：默认写入 `My Documents/iWorkHelper/logs`，支持滚动与最大大小限制。

## 8. 安全与权限
- Outlook 加载项运行在用户上下文；文件归档需具备写权限。
- 避免在主线程执行耗时操作；所有文件系统与 PDF 操作在后台执行。

## 9. 示例：Ribbon 触发代码（伪示例）
```vb
Private Async Sub btnProcess_Click(sender As Object, e As RibbonControlEventArgs) Handles btnProcess.Click
    Dim progress = New ProgressReporter(AddressOf UpdateProgressUi)
    Dim result = Await _mailProcessor.ProcessSelectedItems(progress)
    If result.Success Then
        MessageBox.Show("处理完成")
    Else
        MessageBox.Show("处理失败: " & result.Message)
    End If
End Sub
```

---

附录 A：术语表
- PdfPig：用于从 PDF 中提取文本的 .NET 库。
- iTextSharp：用于 PDF 合并与操作的 .NET 库。
- 滴滴文档：包含发票与行程单的两类 PDF，需识别并合并。