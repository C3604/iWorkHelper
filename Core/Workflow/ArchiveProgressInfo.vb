''' <summary>
''' 归档处理阶段。用于进度窗口的“当前阶段描述”。
''' </summary>
Public Enum ArchiveStage
    Reading = 0        ' 正在读取邮件附件
    SavingPdf = 1      ' 正在保存 PDF 附件
    Merging = 2        ' 正在合并发票和行程单
    ExtractingText = 3 ' 正在提取 PDF 文本
    CallingOcr = 4     ' 正在调用百度 OCR
    ParsingFields = 5  ' 正在解析票据字段
    Naming = 6         ' 正在生成归档文件名
    Archiving = 7      ' 正在归档文件
    WritingResult = 8  ' 正在写入处理结果
    Completed = 9      ' 已完成
    Classifying = 10   ' 正在判断邮件类型
    ProcessingGeneral = 11 ' 正在处理常规发票
    ProcessingUnknown = 12 ' 正在处理未识别 PDF
    NoPdfSkipped = 13  ' 当前邮件无 PDF，已跳过
End Enum

''' <summary>
''' 进度信息（以“邮件”为统计单位）。
''' </summary>
Public Class ArchiveProgressInfo

    Public Property TotalEmails As Integer
    Public Property ProcessedEmails As Integer
    Public Property CurrentEmailIndex As Integer
    Public Property CurrentEmailSubject As String
    Public Property Stage As ArchiveStage

    ''' <summary>备注（如“上一封处理失败，继续下一封”），可空。</summary>
    Public Property Note As String

    ''' <summary>阶段中文描述。</summary>
    Public ReadOnly Property StageText As String
        Get
            Select Case Stage
                Case ArchiveStage.Reading : Return "正在读取邮件附件"
                Case ArchiveStage.SavingPdf : Return "正在保存 PDF 附件"
                Case ArchiveStage.Merging : Return "正在合并发票和行程单"
                Case ArchiveStage.ExtractingText : Return "正在提取 PDF 文本"
                Case ArchiveStage.CallingOcr : Return "正在调用百度 OCR"
                Case ArchiveStage.ParsingFields : Return "正在解析票据字段"
                Case ArchiveStage.Naming : Return "正在生成归档文件名"
                Case ArchiveStage.Archiving : Return "正在归档文件"
                Case ArchiveStage.WritingResult : Return "正在写入处理结果"
                Case ArchiveStage.Completed : Return "已完成"
                Case ArchiveStage.Classifying : Return "正在判断邮件类型"
                Case ArchiveStage.ProcessingGeneral : Return "正在处理常规发票"
                Case ArchiveStage.ProcessingUnknown : Return "正在处理未识别 PDF"
                Case ArchiveStage.NoPdfSkipped : Return "当前邮件无 PDF，已跳过"
                Case Else : Return ""
            End Select
        End Get
    End Property

    ''' <summary>百分比（0-100）。以邮件为单位。</summary>
    Public ReadOnly Property Percent As Integer
        Get
            If TotalEmails <= 0 Then Return 0
            Dim p As Integer = CInt(Math.Floor(ProcessedEmails * 100.0 / TotalEmails))
            If p < 0 Then Return 0
            If p > 100 Then Return 100
            Return p
        End Get
    End Property

End Class
