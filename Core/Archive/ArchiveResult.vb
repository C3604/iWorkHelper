Imports System.Collections.Generic

''' <summary>
''' 单封邮件的归档处理结果（归档单位为“邮件”，产出一份合并 PDF）。
''' 保留可追踪信息：原始邮件、原始 PDF、合并临时 PDF、最终归档 PDF、识别来源等。
''' </summary>
Public Class ArchiveItemResult

    Public Sub New()
        MissingFields = New List(Of String)()
        SourcePdfNames = New List(Of String)()
        SourcePdfPaths = New List(Of String)()
    End Sub

    ' —— 邮件维度 ——
    Public Property MailSubject As String
    Public Property SenderName As String
    Public Property MailIndex As Integer

    ' —— 原始 PDF（合并前）——
    Public Property PdfCount As Integer
    Public Property SourcePdfNames As List(Of String)
    Public Property SourcePdfPaths As List(Of String)

    ' —— 合并与归档 ——
    ''' <summary>合并后的临时 PDF 路径（单 PDF 时为该 PDF 的临时副本）。</summary>
    Public Property MergedTempPath As String
    Public Property TargetPath As String
    Public Property FinalFileName As String

    ' 兼容旧字段（首个/原始附件名）
    Public Property OriginalFileName As String
    Public Property SourceTempPath As String

    Public Property Status As ProcessStatus
    Public Property DocumentType As InvoiceDocumentType
    Public Property Message As String

    ''' <summary>识别来源（LocalText/BaiduOcr/Mixed/Failed）。</summary>
    Public Property RecognitionSource As RecognitionSource

    ''' <summary>所用命名规则名称。</summary>
    Public Property NamingRule As String

    ''' <summary>命名时缺失的字段。</summary>
    Public Property MissingFields As List(Of String)

    ''' <summary>行程明细数量（网约车行程单）。</summary>
    Public Property TripCount As Integer

    ''' <summary>处理类型：滴滴合并 / 常规发票 / 未识别PDF / 跳过。用于汇总分类统计。</summary>
    Public Property ProcessingKind As String

    ''' <summary>命名是否触发了未识别 fallback。</summary>
    Public Property NamingFallbackTriggered As Boolean

    ''' <summary>命名模板中的未知占位符。</summary>
    Public Property UnknownPlaceholders As List(Of String)

    ''' <summary>是否发生了合并（PDF 数量 > 1）。</summary>
    Public ReadOnly Property Merged As Boolean
        Get
            Return PdfCount > 1
        End Get
    End Property

End Class

''' <summary>
''' 批量归档汇总。含成功/失败/跳过/需 OCR 计数与逐条明细。
''' </summary>
Public Class ArchiveBatchResult

    Public Sub New()
        Items = New List(Of ArchiveItemResult)()
        Messages = New List(Of String)()
    End Sub

    Public Property Items As List(Of ArchiveItemResult)
    Public Property Messages As List(Of String)

    ''' <summary>整体状态（配置缺失/成功/部分成功/失败）。</summary>
    Public Property OverallStatus As ProcessStatus

    ''' <summary>本次归档批次 ID（用于日志/报告关联）。</summary>
    Public Property BatchId As String

    ''' <summary>本次归档结果报告文件路径（可为空）。</summary>
    Public Property ReportPath As String

    Public ReadOnly Property TotalCount As Integer
        Get
            Return Items.Count
        End Get
    End Property

    Public ReadOnly Property SuccessCount As Integer
        Get
            Return CountByStatus(ProcessStatus.Success)
        End Get
    End Property

    Public ReadOnly Property FailureCount As Integer
        Get
            Return CountByStatus(ProcessStatus.Failure)
        End Get
    End Property

    Public ReadOnly Property SkippedCount As Integer
        Get
            Return CountByStatus(ProcessStatus.Skipped)
        End Get
    End Property

    Public ReadOnly Property NeedsOcrCount As Integer
        Get
            Return CountByStatus(ProcessStatus.NeedsOcr)
        End Get
    End Property

    ''' <summary>部分成功（识别不完整但已归档）。</summary>
    Public ReadOnly Property PartialCount As Integer
        Get
            Return CountByStatus(ProcessStatus.PartialSuccess)
        End Get
    End Property

    Private Function CountByStatus(status As ProcessStatus) As Integer
        Dim n As Integer = 0
        For Each it As ArchiveItemResult In Items
            If it.Status = status Then
                n += 1
            End If
        Next
        Return n
    End Function

    Public Sub AddMessage(text As String)
        If Not String.IsNullOrEmpty(text) Then
            Messages.Add(text)
        End If
    End Sub

    ''' <summary>合并归档的邮件数（PDF 数 > 1）。</summary>
    Public ReadOnly Property MergedCount As Integer
        Get
            Dim n As Integer = 0
            For Each it As ArchiveItemResult In Items
                If it.Merged Then n += 1
            Next
            Return n
        End Get
    End Property

    ''' <summary>去重后的邮件数（一封邮件可能产出多个归档项）。</summary>
    Public ReadOnly Property MailCount As Integer
        Get
            Dim seen As New HashSet(Of Integer)()
            For Each it As ArchiveItemResult In Items
                seen.Add(it.MailIndex)
            Next
            Return seen.Count
        End Get
    End Property

    Public ReadOnly Property DidiCount As Integer
        Get
            Return CountByKind("滴滴合并")
        End Get
    End Property
    Public ReadOnly Property GeneralInvoiceCount As Integer
        Get
            Return CountByKind("常规发票")
        End Get
    End Property
    Public ReadOnly Property UnknownPdfCount As Integer
        Get
            Return CountByKind("未识别PDF")
        End Get
    End Property

    Private Function CountByKind(kind As String) As Integer
        Dim n As Integer = 0
        For Each it As ArchiveItemResult In Items
            If String.Equals(it.ProcessingKind, kind, StringComparison.Ordinal) Then n += 1
        Next
        Return n
    End Function

    ''' <summary>生成简明汇总文本（以“邮件”为单位 + 分类统计），供 UI 展示。</summary>
    Public Function BuildSummaryText() As String
        Return String.Format("共 {0} 封邮件（滴滴 {7}，常规发票 {8}，未识别PDF {9}）：成功 {1}，部分成功 {2}，需要 OCR {3}，跳过 {4}，失败 {5}（合并归档 {6}）。",
                             MailCount, SuccessCount, PartialCount, NeedsOcrCount, SkippedCount, FailureCount, MergedCount,
                             DidiCount, GeneralInvoiceCount, UnknownPdfCount)
    End Function

End Class
