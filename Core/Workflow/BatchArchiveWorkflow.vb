Imports System.IO
Imports System.Collections.Generic
Imports Outlook = Microsoft.Office.Interop.Outlook

''' <summary>
''' 批量归档工作流编排器（**按邮件/PDF 类型分流**）：
'''  - 滴滴发票邮件：发票 PDF + 行程单 PDF 合并为一个 PDF，统一命名归档；
'''  - 常规发票邮件：每个发票 PDF 单独识别、命名（常规发票模板）、归档；
'''  - 未识别 PDF：每个按 未识别_{原始文件名} 归档；
'''  - 无 PDF 邮件：跳过（不报错，计入跳过）。
''' 进度以邮件为主维度；单封邮件/单个 PDF 失败不影响其它；全程日志 + 报告。
''' </summary>
Public Class BatchArchiveWorkflow

    Private _carryNote As String = Nothing

    ' 共享服务/上下文（Run 内初始化，供分流方法复用）
    Private _archiveFolder As String
    Private _tempDir As String
    Private _runToken As String
    Private _generalTemplate As String
    Private _pipeline As RecognitionPipeline
    Private _merger As InvoiceRecognitionMerger
    Private _mergeSvc As PdfMergeService
    Private _executor As ArchiveExecutor
    Private _planner As ArchivePlanner
    Private _classifier As MailProcessingClassifier
    Private _ocrEnabled As Boolean
    Private _total As Integer

    ''' <summary>无进度上报的重载（离线/兼容）。</summary>
    Public Function Run(application As Outlook.Application) As ArchiveBatchResult
        Return Run(application, New NullArchiveProgressReporter())
    End Function

    ''' <summary>执行批量归档（分流），并上报进度。</summary>
    Public Function Run(application As Outlook.Application, reporter As IArchiveProgressReporter) As ArchiveBatchResult
        If reporter Is Nothing Then reporter = New NullArchiveProgressReporter()
        Dim batch As New ArchiveBatchResult()

        _archiveFolder = SafeSetting(Function() My.Settings.ArchiveFolderPath)
        Dim ocrOptions As BaiduOcrOptions = OcrConfigProvider.LoadBaiduOptions()
        AppLogger.Initialize(_archiveFolder)
        AppLogger.Info("=== 批量归档开始（分流）=== " & ocrOptions.ToSafeSummary())

        Dim templates As New NamingTemplates()
        templates.UnifiedTemplate = SafeSetting(Function() My.Settings.UnifiedNameTemplate)
        templates.GeneralInvoiceTemplate = SafeSetting(Function() My.Settings.GeneralInvoiceNameTemplate)
        _generalTemplate = If(String.IsNullOrWhiteSpace(templates.GeneralInvoiceTemplate), NamingTemplates.DefaultGeneralInvoiceTemplate, templates.GeneralInvoiceTemplate)

        _planner = New ArchivePlanner(templates)
        Dim validate As Result = _planner.ValidateArchiveFolder(_archiveFolder)
        If Not validate.IsSuccess Then
            batch.OverallStatus = validate.Status
            batch.AddMessage(validate.Message)
            AppLogger.Warn("归档目录校验失败：" & validate.Message)
            Return batch
        End If

        SafeReport(reporter, New ArchiveProgressInfo With {.Stage = ArchiveStage.Reading})
        Dim reader As New MailAttachmentReader()
        Dim grouping As MailPdfGroupingResult = reader.ReadSelectedPdfGroups(application)
        For Each m As String In grouping.Messages
            batch.AddMessage(m)
        Next
        If grouping.Groups.Count = 0 Then
            batch.OverallStatus = ProcessStatus.Skipped
            batch.AddMessage("没有可处理的邮件（无 PDF 附件）。")
            AppLogger.Info("无可处理邮件。")
            Return batch
        End If

        _pipeline = New RecognitionPipeline(ocrOptions)
        _merger = New InvoiceRecognitionMerger()
        _mergeSvc = New PdfMergeService()
        _executor = New ArchiveExecutor()
        _classifier = New MailProcessingClassifier()
        _tempDir = PathHelper.GetTempWorkDirectory()
        _runToken = DateTime.Now.ToString("yyyyMMddHHmmss")
        batch.BatchId = "B" & _runToken
        Dim reportStamp As String = DateTime.Now.ToString("yyyyMMdd-HHmmss")
        _ocrEnabled = ocrOptions.IsConfigured()
        _total = grouping.Groups.Count
        AppLogger.Info("批次 ID=" & batch.BatchId & "，待处理邮件=" & _total)

        Dim processed As Integer = 0
        For Each group As MailPdfGroup In grouping.Groups
            Try
                ' 逐 PDF 识别 + 分类
                ReportStage(reporter, processed, group, ArchiveStage.ExtractingText)
                Dim perPdf As New List(Of PdfAttachmentClassificationResult)()
                For Each pdf As MailAttachmentItem In group.Pdfs
                    If _ocrEnabled Then ReportStage(reporter, processed, group, ArchiveStage.CallingOcr)
                    Dim recog As InvoiceRecognitionResult = _pipeline.Recognize(pdf.TempFilePath, pdf.OriginalFileName)
                    Dim cls As PdfAttachmentClassification = _classifier.ClassifyPdf(recog, recog.RawText, pdf.OriginalFileName, group.MailSubject)
                    perPdf.Add(New PdfAttachmentClassificationResult With {.Attachment = pdf, .Classification = cls, .Recognition = recog})
                Next

                ' 邮件类型
                ReportStage(reporter, processed, group, ArchiveStage.Classifying)
                Dim classes As New List(Of PdfAttachmentClassification)()
                For Each pc As PdfAttachmentClassificationResult In perPdf
                    classes.Add(pc.Classification)
                Next
                Dim mailType As MailProcessingType = _classifier.ClassifyMail(classes)
                AppLogger.Info(String.Format("邮件[{0}]《{1}》类型={2}，PDF数={3}",
                                             group.Index, PrivacySafeFormatter.MaskSubject(group.MailSubject), mailType, group.PdfCount))

                Select Case mailType
                    Case MailProcessingType.DidiInvoiceMail
                        FinishItem(batch, ProcessDidiMail(group, perPdf, reporter, processed), group, reporter, processed)
                    Case MailProcessingType.GeneralInvoiceMail, MailProcessingType.MixedPdfMail, MailProcessingType.UnknownPdfOnlyMail
                        Dim idx As Integer = 0
                        For Each pc As PdfAttachmentClassificationResult In perPdf
                            idx += 1
                            Dim it As ArchiveItemResult
                            If pc.Classification = PdfAttachmentClassification.GeneralInvoicePdf Then
                                ReportStage(reporter, processed, group, ArchiveStage.ProcessingGeneral)
                                it = ProcessGeneralPdf(group, pc, idx)
                            Else
                                ReportStage(reporter, processed, group, ArchiveStage.ProcessingUnknown)
                                it = ProcessUnknownPdf(group, pc, idx)
                            End If
                            FinishItem(batch, it, group, reporter, processed)
                        Next
                    Case Else ' NoPdfMail（分组已过滤，一般不会到此）
                        ReportStage(reporter, processed, group, ArchiveStage.NoPdfSkipped)
                        FinishItem(batch, BuildSkippedItem(group), group, reporter, processed)
                End Select

            Catch ex As Exception
                Dim unk As New AppError(AppErrorCode.Unknown, ExceptionFormatter.ToUserMessage(ex))
                AppLogger.Error("处理邮件异常：" & PrivacySafeFormatter.MaskSubject(group.MailSubject), ex)
                FinishItem(batch, BuildFailedItem(group, unk.UserMessage), group, reporter, processed)
            End Try

            processed += 1
        Next

        batch.OverallStatus = ComputeOverall(batch)
        SafeReport(reporter, New ArchiveProgressInfo With {.TotalEmails = _total, .ProcessedEmails = _total, .Stage = ArchiveStage.Completed})
        batch.ReportPath = ArchiveReportWriter.Write(batch, _archiveFolder, reportStamp)
        AppLogger.Info("=== 批量归档结束（批次 " & batch.BatchId & "）：" & batch.BuildSummaryText() & " ===")
        Return batch
    End Function

    ' ===== 滴滴合并归档（保持原有链路） =====
    Private Function ProcessDidiMail(group As MailPdfGroup, perPdf As List(Of PdfAttachmentClassificationResult),
                                     reporter As IArchiveProgressReporter, processed As Integer) As ArchiveItemResult
        Dim item As New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .SenderName = group.SenderName, .MailIndex = group.Index,
            .PdfCount = group.PdfCount, .ProcessingKind = "滴滴合并",
            .OriginalFileName = If(group.Pdfs.Count > 0, group.Pdfs(0).OriginalFileName, Nothing)}
        For Each pc As PdfAttachmentClassificationResult In perPdf
            item.SourcePdfNames.Add(pc.Attachment.OriginalFileName)
            item.SourcePdfPaths.Add(pc.Attachment.TempFilePath)
        Next

        Dim recognized As New List(Of InvoiceRecognitionResult)()
        For Each pc As PdfAttachmentClassificationResult In perPdf
            recognized.Add(pc.Recognition)
        Next
        ReportStage(reporter, processed, group, ArchiveStage.ParsingFields)
        Dim mergedRecog As InvoiceRecognitionResult = _merger.Merge(recognized)
        item.DocumentType = mergedRecog.DocumentType
        item.RecognitionSource = mergedRecog.Source
        If mergedRecog.Invoice IsNot Nothing AndAlso mergedRecog.Invoice.Trips IsNot Nothing Then item.TripCount = mergedRecog.Invoice.Trips.Count

        ' 合并 PDF（发票在前、行程单在后）
        ReportStage(reporter, processed, group, ArchiveStage.Merging)
        Dim orderedPaths As List(Of String) = OrderDidiPaths(perPdf)
        Dim mergedTemp As String = PathHelper.GetNonConflictingPath(_tempDir, "merged_" & _runToken & "_" & group.Index & ".pdf")
        Dim mergeRes As Result = _mergeSvc.Merge(orderedPaths, mergedTemp)
        If Not mergeRes.IsSuccess Then
            item.Status = ProcessStatus.Failure
            Dim mergeErr As New AppError(AppErrorCode.PdfMergeFailed, mergeRes.Message) : mergeErr.LogSelf()
            item.Message = mergeErr.UserMessage
            Return item
        End If
        item.MergedTempPath = mergedTemp

        ReportStage(reporter, processed, group, ArchiveStage.Naming)
        Dim token As String = _runToken & "_" & group.Index.ToString()
        Dim plan As ArchiveTargetPlan = _planner.PlanTarget(_archiveFolder, mergedRecog.Invoice, mergedRecog.DocumentType,
                                                            item.OriginalFileName, token, group.MailSubject, group.PdfCount, mergedRecog.Source.ToString())
        ApplyNamePlan(item, plan)

        ReportStage(reporter, processed, group, ArchiveStage.Archiving)
        Dim exec As Result = _executor.Execute(mergedTemp, plan.FullPath)
        If exec.IsSuccess Then
            item.Status = MapRecognitionToProcess(mergedRecog.Status)
            item.Message = BuildItemReason(item.Status, mergedRecog)
            For Each pc As PdfAttachmentClassificationResult In perPdf
                SafeDeleteTemp(pc.Attachment.TempFilePath)
            Next
            SafeDeleteTemp(mergedTemp)
        Else
            item.Status = ProcessStatus.Failure
            Dim copyErr As New AppError(AppErrorCode.FileCopyFailed, exec.Message) : copyErr.LogSelf()
            item.Message = copyErr.UserMessage & "（临时文件已保留，可在本机临时目录排查）"
        End If
        Return item
    End Function

    ' ===== 常规发票（单 PDF 归档） =====
    Private Function ProcessGeneralPdf(group As MailPdfGroup, pc As PdfAttachmentClassificationResult, idx As Integer) As ArchiveItemResult
        Dim pdf As MailAttachmentItem = pc.Attachment
        Dim recog As InvoiceRecognitionResult = pc.Recognition
        Dim item As New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .SenderName = group.SenderName, .MailIndex = group.Index,
            .PdfCount = 1, .ProcessingKind = "常规发票", .OriginalFileName = pdf.OriginalFileName,
            .DocumentType = recog.DocumentType, .RecognitionSource = recog.Source}
        item.SourcePdfNames.Add(pdf.OriginalFileName)
        item.SourcePdfPaths.Add(pdf.TempFilePath)

        Try
            Dim token As String = _runToken & "_" & group.Index & "_" & idx
            Dim plan As ArchiveTargetPlan = _planner.PlanGeneralInvoiceTarget(_archiveFolder, recog.Invoice, pdf.OriginalFileName, token,
                                                                             group.MailSubject, recog.Source.ToString(), _generalTemplate)
            ApplyNamePlan(item, plan)
            Dim exec As Result = _executor.Execute(pdf.TempFilePath, plan.FullPath)
            If exec.IsSuccess Then
                item.Status = MapRecognitionToProcess(recog.Status)
                item.Message = "常规发票：" & BuildItemReason(item.Status, recog)
                SafeDeleteTemp(pdf.TempFilePath)
            Else
                item.Status = ProcessStatus.Failure
                Dim copyErr As New AppError(AppErrorCode.FileCopyFailed, exec.Message) : copyErr.LogSelf()
                item.Message = copyErr.UserMessage & "（临时文件已保留，可在本机临时目录排查）"
            End If
        Catch ex As Exception
            item.Status = ProcessStatus.Failure
            item.Message = New AppError(AppErrorCode.Unknown, ExceptionFormatter.ToUserMessage(ex)).UserMessage
            AppLogger.Error("处理常规发票异常：" & PrivacySafeFormatter.MaskFileName(pdf.OriginalFileName), ex)
        End Try
        Return item
    End Function

    ' ===== 未识别 PDF（未识别_原始文件名） =====
    Private Function ProcessUnknownPdf(group As MailPdfGroup, pc As PdfAttachmentClassificationResult, idx As Integer) As ArchiveItemResult
        Dim pdf As MailAttachmentItem = pc.Attachment
        Dim item As New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .SenderName = group.SenderName, .MailIndex = group.Index,
            .PdfCount = 1, .ProcessingKind = "未识别PDF", .OriginalFileName = pdf.OriginalFileName,
            .DocumentType = InvoiceDocumentType.Unknown, .RecognitionSource = RecognitionSource.None}
        item.SourcePdfNames.Add(pdf.OriginalFileName)
        item.SourcePdfPaths.Add(pdf.TempFilePath)

        Try
            Dim plan As ArchiveTargetPlan = _planner.PlanUnknownTarget(_archiveFolder, pdf.OriginalFileName)
            ApplyNamePlan(item, plan)
            Dim exec As Result = _executor.Execute(pdf.TempFilePath, plan.FullPath)
            If exec.IsSuccess Then
                item.Status = ProcessStatus.Success
                item.Message = "未识别 PDF，已按原文件名归档。"
                SafeDeleteTemp(pdf.TempFilePath)
            Else
                item.Status = ProcessStatus.Failure
                Dim copyErr As New AppError(AppErrorCode.FileCopyFailed, exec.Message) : copyErr.LogSelf()
                item.Message = copyErr.UserMessage & "（临时文件已保留，可在本机临时目录排查）"
            End If
        Catch ex As Exception
            item.Status = ProcessStatus.Failure
            item.Message = New AppError(AppErrorCode.Unknown, ExceptionFormatter.ToUserMessage(ex)).UserMessage
            AppLogger.Error("处理未识别 PDF 异常：" & PrivacySafeFormatter.MaskFileName(pdf.OriginalFileName), ex)
        End Try
        Return item
    End Function

    Private Function BuildSkippedItem(group As MailPdfGroup) As ArchiveItemResult
        Return New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .MailIndex = group.Index, .PdfCount = group.PdfCount,
            .ProcessingKind = "跳过", .Status = ProcessStatus.Skipped, .Message = "无 PDF 附件，已跳过。"}
    End Function

    Private Function BuildFailedItem(group As MailPdfGroup, reason As String) As ArchiveItemResult
        Return New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .MailIndex = group.Index, .PdfCount = group.PdfCount,
            .ProcessingKind = "失败", .Status = ProcessStatus.Failure, .Message = reason}
    End Function

    Private Sub ApplyNamePlan(item As ArchiveItemResult, plan As ArchiveTargetPlan)
        item.TargetPath = plan.FullPath
        item.FinalFileName = Path.GetFileName(plan.FullPath)
        item.NamingRule = plan.NamePlan.RuleName
        item.MissingFields = plan.NamePlan.MissingFields
        item.NamingFallbackTriggered = plan.NamePlan.FallbackTriggered
        item.UnknownPlaceholders = plan.NamePlan.UnknownPlaceholders
    End Sub

    Private Function BuildItemReason(status As ProcessStatus, recog As InvoiceRecognitionResult) As String
        Select Case status
            Case ProcessStatus.Success : Return "识别成功并归档。"
            Case ProcessStatus.PartialSuccess : Return If(String.IsNullOrEmpty(recog.Message), "部分成功，已归档（字段不完整）。", recog.Message)
            Case ProcessStatus.NeedsOcr : Return If(String.IsNullOrEmpty(recog.Message), UserFriendlyMessageProvider.Describe(AppErrorCode.LocalFieldsInsufficient).UserMessage, recog.Message)
            Case Else : Return If(String.IsNullOrEmpty(recog.Message), "已处理。", recog.Message)
        End Select
    End Function

    Private Sub FinishItem(batch As ArchiveBatchResult, item As ArchiveItemResult, group As MailPdfGroup, reporter As IArchiveProgressReporter, processed As Integer)
        ReportStage(reporter, processed, group, ArchiveStage.WritingResult)
        _carryNote = If(item.Status = ProcessStatus.Failure, "上一封邮件处理失败，已继续下一封。", Nothing)
        batch.Items.Add(item)
        AppLogger.Info(String.Format("邮件[{0}]【{1}】类型={2}, PDF={3}, 来源={4}, 状态={5}, 归档={6}, fallback={7}",
                                     item.MailIndex, PrivacySafeFormatter.MaskSubject(item.MailSubject), item.ProcessingKind,
                                     PrivacySafeFormatter.MaskFileList(item.SourcePdfNames),
                                     item.RecognitionSource, item.Status,
                                     If(String.IsNullOrEmpty(item.FinalFileName), "(未归档)", PrivacySafeFormatter.MaskFileName(item.FinalFileName)),
                                     item.NamingFallbackTriggered))
    End Sub

    ''' <summary>滴滴合并顺序：滴滴发票在前、滴滴行程单在后、其它最后。</summary>
    Private Function OrderDidiPaths(perPdf As List(Of PdfAttachmentClassificationResult)) As List(Of String)
        Dim paths As New List(Of String)()
        Dim order As PdfAttachmentClassification() = New PdfAttachmentClassification() {
            PdfAttachmentClassification.DidiInvoicePdf, PdfAttachmentClassification.DidiTripPdf,
            PdfAttachmentClassification.GeneralInvoicePdf, PdfAttachmentClassification.UnknownPdf}
        For Each k As PdfAttachmentClassification In order
            For Each pc As PdfAttachmentClassificationResult In perPdf
                If pc.Classification = k Then paths.Add(pc.Attachment.TempFilePath)
            Next
        Next
        Return paths
    End Function

    Private Sub ReportStage(reporter As IArchiveProgressReporter, processed As Integer, group As MailPdfGroup, stage As ArchiveStage)
        SafeReport(reporter, New ArchiveProgressInfo With {
            .TotalEmails = _total, .ProcessedEmails = processed,
            .CurrentEmailIndex = group.Index, .CurrentEmailSubject = group.MailSubject,
            .Stage = stage, .Note = _carryNote})
    End Sub

    Private Sub SafeReport(reporter As IArchiveProgressReporter, info As ArchiveProgressInfo)
        Try
            reporter.Report(info)
        Catch
        End Try
    End Sub

    Private Function MapRecognitionToProcess(status As RecognitionStatus) As ProcessStatus
        Select Case status
            Case RecognitionStatus.Success : Return ProcessStatus.Success
            Case RecognitionStatus.PartialSuccess : Return ProcessStatus.PartialSuccess
            Case RecognitionStatus.NeedsOcr : Return ProcessStatus.NeedsOcr
            Case RecognitionStatus.ConfigurationMissing : Return ProcessStatus.NeedsOcr
            Case Else : Return ProcessStatus.PartialSuccess
        End Select
    End Function

    Private Function ComputeOverall(batch As ArchiveBatchResult) As ProcessStatus
        If batch.TotalCount = 0 Then Return ProcessStatus.Skipped
        If batch.FailureCount = 0 AndAlso batch.NeedsOcrCount = 0 AndAlso batch.PartialCount = 0 Then Return ProcessStatus.Success
        If batch.SuccessCount = 0 AndAlso batch.PartialCount = 0 Then
            If batch.FailureCount = batch.TotalCount Then Return ProcessStatus.Failure
            Return ProcessStatus.PartialSuccess
        End If
        Return ProcessStatus.PartialSuccess
    End Function

    Private Function SafeSetting(getter As Func(Of String)) As String
        Try
            Return getter()
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub SafeDeleteTemp(path As String)
        Try
            If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then File.Delete(path)
        Catch
        End Try
    End Sub

End Class
