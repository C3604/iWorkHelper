Imports System.Collections.Generic

''' <summary>
''' 识别结果合并器：把同一封邮件内多个 PDF（发票 / 行程单 / 其它）的识别结果合并为一个
''' InvoiceRecognitionResult，绑定到合并后的归档 PDF。
''' 策略：
'''  - 发票字段以“含发票号码”的结果为准；
'''  - 行程明细做并集；扩展字段/诊断字段做并集；
'''  - 命名优先发票字段（见 ArchiveNamingRule）；
'''  - 识别来源综合：全本地=LocalText；含在线=Mixed/BaiduOcr；全失败=Failed。
''' </summary>
Public Class InvoiceRecognitionMerger

    ''' <summary>PDF 分类。</summary>
    Public Enum PdfKind
        Invoice = 0
        Trip = 1
        Unknown = 2
    End Enum

    ''' <summary>依识别结果判断 PDF 类型：有发票号码=发票；有行程/行程单标记且无发票号码=行程单；否则未知。</summary>
    Public Function Classify(result As InvoiceRecognitionResult) As PdfKind
        If result Is Nothing OrElse result.Invoice Is Nothing Then
            Return PdfKind.Unknown
        End If
        If Not String.IsNullOrWhiteSpace(result.Invoice.InvoiceNumber) Then
            Return PdfKind.Invoice
        End If
        If (result.Invoice.Trips IsNot Nothing AndAlso result.Invoice.Trips.Count > 0) _
           OrElse result.DocumentType = InvoiceDocumentType.RideTripStatement Then
            Return PdfKind.Trip
        End If
        Return PdfKind.Unknown
    End Function

    ''' <summary>
    ''' 合并多个识别结果为一个。results 顺序不限。
    ''' </summary>
    Public Function Merge(results As List(Of InvoiceRecognitionResult)) As InvoiceRecognitionResult
        If results Is Nothing OrElse results.Count = 0 Then
            Return InvoiceRecognitionResult.Failure("无识别结果可合并。")
        End If
        If results.Count = 1 Then
            Return results(0)
        End If

        ' 选基准：优先“含发票号码”的结果。
        Dim baseResult As InvoiceRecognitionResult = Nothing
        For Each r As InvoiceRecognitionResult In results
            If r IsNot Nothing AndAlso r.Invoice IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(r.Invoice.InvoiceNumber) Then
                baseResult = r
                Exit For
            End If
        Next
        If baseResult Is Nothing Then baseResult = results(0)

        Dim merged As New InvoiceRecognitionResult()
        merged.RecognizerName = "InvoiceRecognitionMerger"
        Dim inv As InvoiceInfo = merged.Invoice

        ' 逐个叠加（基准优先，其它填补空缺 + 并集）。
        Dim ordered As New List(Of InvoiceRecognitionResult)()
        ordered.Add(baseResult)
        For Each r As InvoiceRecognitionResult In results
            If Not ReferenceEquals(r, baseResult) Then ordered.Add(r)
        Next

        Dim anyUsable As Boolean = False
        For Each r As InvoiceRecognitionResult In ordered
            If r Is Nothing Then Continue For
            If r.IsUsable Then anyUsable = True
            OverlayInvoice(inv, r.Invoice)
            If r.Fields IsNot Nothing Then merged.Fields.AddRange(r.Fields)
            If r.Messages IsNot Nothing Then merged.Messages.AddRange(r.Messages)
            If String.IsNullOrEmpty(merged.RawText) AndAlso Not String.IsNullOrEmpty(r.RawText) Then merged.RawText = r.RawText
        Next

        ' 文档类型：有发票号码 → 发票；否则行程单/未知。
        If Not String.IsNullOrWhiteSpace(inv.InvoiceNumber) Then
            merged.DocumentType = InvoiceDocumentType.VatInvoice
        ElseIf inv.Trips.Count > 0 Then
            merged.DocumentType = InvoiceDocumentType.RideTripStatement
        Else
            merged.DocumentType = InvoiceDocumentType.Unknown
        End If

        merged.Source = CombineSource(ordered)
        merged.MissingKeyFields = KeyFieldEvaluator.GetMissingKeyFields(inv, merged.DocumentType)

        If inv.HasAnyCoreField() Then
            merged.Status = If(merged.MissingKeyFields.Count = 0, RecognitionStatus.Success, RecognitionStatus.PartialSuccess)
            merged.Message = "已合并 " & results.Count & " 份 PDF 的识别结果。"
        ElseIf anyUsable Then
            merged.Status = RecognitionStatus.PartialSuccess
            merged.Message = "合并结果字段不完整。"
        Else
            merged.Status = RecognitionStatus.NeedsOcr
            merged.Message = "合并后仍未提取到关键字段。"
        End If

        AppLogger.Info(String.Format("识别结果合并：来源={0}, 类型={1}, 行程数={2}, 状态={3}",
                                     merged.Source, merged.DocumentType, inv.Trips.Count, merged.Status))
        Return merged
    End Function

    ''' <summary>把 src 的字段填补到 dst（dst 空缺才填；行程与扩展字段做并集）。</summary>
    Private Sub OverlayInvoice(dst As InvoiceInfo, src As InvoiceInfo)
        If src Is Nothing Then Return
        FillIfEmpty(Function() dst.InvoiceNumber, Sub(x) dst.InvoiceNumber = x, src.InvoiceNumber)
        FillIfEmpty(Function() dst.InvoiceCode, Sub(x) dst.InvoiceCode = x, src.InvoiceCode)
        FillIfEmpty(Function() dst.InvoiceDate, Sub(x) dst.InvoiceDate = x, src.InvoiceDate)
        FillIfEmpty(Function() dst.SellerName, Sub(x) dst.SellerName = x, src.SellerName)
        FillIfEmpty(Function() dst.SellerTaxId, Sub(x) dst.SellerTaxId = x, src.SellerTaxId)
        FillIfEmpty(Function() dst.BuyerName, Sub(x) dst.BuyerName = x, src.BuyerName)
        FillIfEmpty(Function() dst.BuyerTaxId, Sub(x) dst.BuyerTaxId = x, src.BuyerTaxId)
        FillIfEmpty(Function() dst.ItemName, Sub(x) dst.ItemName = x, src.ItemName)
        FillIfEmpty(Function() dst.Specification, Sub(x) dst.Specification = x, src.Specification)
        FillIfEmpty(Function() dst.Unit, Sub(x) dst.Unit = x, src.Unit)
        FillIfEmpty(Function() dst.Quantity, Sub(x) dst.Quantity = x, src.Quantity)
        FillIfEmpty(Function() dst.UnitPrice, Sub(x) dst.UnitPrice = x, src.UnitPrice)
        FillIfEmpty(Function() dst.Amount, Sub(x) dst.Amount = x, src.Amount)
        FillIfEmpty(Function() dst.TaxRate, Sub(x) dst.TaxRate = x, src.TaxRate)
        FillIfEmpty(Function() dst.TaxAmount, Sub(x) dst.TaxAmount = x, src.TaxAmount)
        FillIfEmpty(Function() dst.TotalWithTax, Sub(x) dst.TotalWithTax = x, src.TotalWithTax)
        FillIfEmpty(Function() dst.Remark, Sub(x) dst.Remark = x, src.Remark)
        FillIfEmpty(Function() dst.CheckCode, Sub(x) dst.CheckCode = x, src.CheckCode)
        FillIfEmpty(Function() dst.Payee, Sub(x) dst.Payee = x, src.Payee)
        FillIfEmpty(Function() dst.Reviewer, Sub(x) dst.Reviewer = x, src.Reviewer)
        FillIfEmpty(Function() dst.Drawer, Sub(x) dst.Drawer = x, src.Drawer)

        If src.StatedTripCount > dst.StatedTripCount Then dst.StatedTripCount = src.StatedTripCount
        If src.Trips IsNot Nothing Then
            For Each t As InvoiceTripInfo In src.Trips
                dst.Trips.Add(t)
            Next
        End If
        If src.ExtendedFields IsNot Nothing Then
            For Each kv In src.ExtendedFields
                If Not dst.ExtendedFields.ContainsKey(kv.Key) Then
                    dst.ExtendedFields(kv.Key) = kv.Value
                End If
            Next
        End If
    End Sub

    Private Sub FillIfEmpty(getter As Func(Of String), setter As Action(Of String), value As String)
        If String.IsNullOrWhiteSpace(getter()) AndAlso Not String.IsNullOrWhiteSpace(value) Then
            setter(value)
        End If
    End Sub

    ''' <summary>综合识别来源。</summary>
    Private Function CombineSource(results As List(Of InvoiceRecognitionResult)) As RecognitionSource
        Dim hasLocal As Boolean = False, hasOnline As Boolean = False, anyUsable As Boolean = False
        For Each r As InvoiceRecognitionResult In results
            If r Is Nothing Then Continue For
            If r.IsUsable Then anyUsable = True
            Select Case r.Source
                Case RecognitionSource.LocalText : hasLocal = True
                Case RecognitionSource.BaiduOcr : hasOnline = True
                Case RecognitionSource.Mixed : hasLocal = True : hasOnline = True
            End Select
        Next
        If Not anyUsable Then Return RecognitionSource.Failed
        If hasLocal AndAlso hasOnline Then Return RecognitionSource.Mixed
        If hasOnline Then Return RecognitionSource.BaiduOcr
        If hasLocal Then Return RecognitionSource.LocalText
        Return RecognitionSource.None
    End Function

End Class
