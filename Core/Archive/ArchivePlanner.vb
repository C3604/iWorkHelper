Imports System.IO

''' <summary>
''' 归档规划器：校验归档目录、依命名规则计算不冲突的目标路径。
''' 不做实际文件操作（交给 ArchiveExecutor），便于单元测试与职责分离。
''' </summary>
Public Class ArchivePlanner

    Private ReadOnly _namingRule As ArchiveNamingRule

    Public Sub New()
        _namingRule = New ArchiveNamingRule()
    End Sub

    Public Sub New(templates As NamingTemplates)
        _namingRule = New ArchiveNamingRule(templates)
    End Sub

    ''' <summary>
    ''' 校验归档目录：为空或无法创建则返回失败（明确错误，不静默）。
    ''' </summary>
    Public Function ValidateArchiveFolder(archiveFolder As String) As Result
        If String.IsNullOrWhiteSpace(archiveFolder) Then
            Return Result.ConfigMissing("尚未设置归档目录，请先在【设置】中选择归档文件夹。")
        End If

        Try
            If Not Directory.Exists(archiveFolder) Then
                ' 尝试创建（用户设置的目录可能尚未建立）。
                Directory.CreateDirectory(archiveFolder)
            End If
        Catch ex As Exception
            Return Result.Fail("归档目录不可用：" & archiveFolder & "，原因：" & ExceptionFormatter.ToUserMessage(ex))
        End Try

        Return Result.Ok()
    End Function

    ''' <summary>
    ''' 计算某附件的目标完整路径（已解决同名冲突）。
    ''' 需在实际拷贝前紧邻调用，以保证冲突检测反映已归档文件。
    ''' </summary>
    Public Function PlanTargetPath(archiveFolder As String, invoice As InvoiceInfo, originalFileName As String, timestampToken As String) As String
        Dim fileName As String = _namingRule.BuildFileName(invoice, originalFileName, timestampToken)
        Return PathHelper.GetNonConflictingPath(archiveFolder, fileName)
    End Function

    ''' <summary>
    ''' 计算目标（含命名方案与不冲突完整路径）。
    ''' </summary>
    Public Function PlanTarget(archiveFolder As String, invoice As InvoiceInfo, docType As InvoiceDocumentType, originalFileName As String, timestampToken As String,
                               Optional mailSubject As String = Nothing, Optional pdfCount As Integer = 0,
                               Optional recognitionSource As String = Nothing, Optional templateOverride As String = Nothing) As ArchiveTargetPlan
        Dim namePlan As ArchiveNamePlan = _namingRule.BuildPlan(invoice, docType, originalFileName, timestampToken, mailSubject, pdfCount, recognitionSource, templateOverride)
        Dim fullPath As String = PathHelper.GetNonConflictingPath(archiveFolder, namePlan.FileName)
        Return New ArchiveTargetPlan With {.FullPath = fullPath, .NamePlan = namePlan}
    End Function

    ''' <summary>未识别 PDF 命名：未识别_{原始文件名}.pdf，清理非法字符 + 不冲突路径。</summary>
    Public Function PlanUnknownTarget(archiveFolder As String, originalFileName As String) As ArchiveTargetPlan
        Dim fileName As String = UnknownPdfNamingRule.BuildFileName(originalFileName)
        Dim fullPath As String = PathHelper.GetNonConflictingPath(archiveFolder, fileName)
        Dim plan As New ArchiveNamePlan With {.FileName = fileName, .RuleName = "未识别PDF:未识别_原始文件名"}
        Return New ArchiveTargetPlan With {.FullPath = fullPath, .NamePlan = plan}
    End Function

    ''' <summary>常规发票命名：使用常规发票模板（templateOverride）。</summary>
    Public Function PlanGeneralInvoiceTarget(archiveFolder As String, invoice As InvoiceInfo, originalFileName As String, timestampToken As String,
                                             mailSubject As String, recognitionSource As String, generalTemplate As String) As ArchiveTargetPlan
        Return PlanTarget(archiveFolder, invoice, InvoiceDocumentType.VatInvoice, originalFileName, timestampToken,
                          mailSubject, 1, recognitionSource, generalTemplate)
    End Function

End Class

''' <summary>归档目标规划结果：完整路径 + 命名方案。</summary>
Public Class ArchiveTargetPlan
    Public Property FullPath As String
    Public Property NamePlan As ArchiveNamePlan
End Class
