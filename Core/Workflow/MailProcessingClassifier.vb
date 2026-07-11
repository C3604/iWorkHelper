Imports System.Collections.Generic

''' <summary>
''' 邮件 / PDF 分流分类器。分类依据优先级（稳健，不只看邮件主题）：
'''  1) 已识别字段与票据类型（recog.Invoice / DocumentType / Trips）；
'''  2) PDF 文本内容关键词；
'''  3) OCR 结构化结果（体现在 recog 字段中）；
'''  4) 邮件主题；
'''  5) 附件文件名。
''' 滴滴同时具备发票+行程特征时优先归滴滴；普通发票无行程特征归常规；无法判断归未识别。
''' </summary>
Public Class MailProcessingClassifier

    Private Shared ReadOnly DidiKeywords As String() = {
        "滴滴出行科技有限公司", "滴滴出行", "滴滴", "行程单", "网约车", "上车时间", "订单号", "DIDI TRAVEL", "TRIP TABLE", "共1笔行程", "笔行程"}

    Private Shared ReadOnly GeneralInvoiceKeywords As String() = {
        "发票代码", "发票号码", "开票日期", "购买方", "销售方", "纳税人识别号", "价税合计", "校验码",
        "增值税电子普通发票", "增值税电子发票", "数电", "电子发票"}

    ''' <summary>对单个 PDF 分类。</summary>
    Public Function ClassifyPdf(recog As InvoiceRecognitionResult, extractedText As String, fileName As String, mailSubject As String) As PdfAttachmentClassification
        Dim text As String = If(extractedText, "")
        Dim subject As String = If(mailSubject, "")
        Dim name As String = If(fileName, "")

        Dim hasInvoiceNo As Boolean = recog IsNot Nothing AndAlso recog.Invoice IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(recog.Invoice.InvoiceNumber)
        Dim hasTripFields As Boolean = recog IsNot Nothing AndAlso recog.Invoice IsNot Nothing AndAlso recog.Invoice.Trips IsNot Nothing AndAlso recog.Invoice.Trips.Count > 0
        Dim sellerDidi As Boolean = recog IsNot Nothing AndAlso recog.Invoice IsNot Nothing AndAlso Not String.IsNullOrEmpty(recog.Invoice.SellerName) AndAlso recog.Invoice.SellerName.Contains("滴滴")

        Dim isDidi As Boolean = hasTripFields OrElse sellerDidi OrElse
                                recog IsNot Nothing AndAlso recog.DocumentType = InvoiceDocumentType.RideTripStatement OrElse
                                ContainsAny(text, DidiKeywords) OrElse ContainsAny(subject, New String() {"滴滴", "行程单", "出行发票", "网约车"}) OrElse
                                ContainsAny(name, New String() {"滴滴", "行程单"})

        Dim hasTripText As Boolean = hasTripFields OrElse text.Contains("行程单") OrElse text.Contains("TRIP TABLE") OrElse text.Contains("上车时间") OrElse text.Contains("笔行程")

        If isDidi Then
            If hasInvoiceNo Then Return PdfAttachmentClassification.DidiInvoicePdf
            If hasTripText Then Return PdfAttachmentClassification.DidiTripPdf
            Return PdfAttachmentClassification.DidiInvoicePdf ' 滴滴但含糊 → 归滴滴发票
        End If

        ' 非滴滴
        If hasInvoiceNo OrElse ContainsAny(text, GeneralInvoiceKeywords) Then
            Return PdfAttachmentClassification.GeneralInvoicePdf
        End If

        Return PdfAttachmentClassification.UnknownPdf
    End Function

    ''' <summary>依据组内各 PDF 分类得出邮件类型。</summary>
    Public Function ClassifyMail(pdfClasses As List(Of PdfAttachmentClassification)) As MailProcessingType
        If pdfClasses Is Nothing OrElse pdfClasses.Count = 0 Then
            Return MailProcessingType.NoPdfMail
        End If

        Dim hasDidi As Boolean = False, hasGeneral As Boolean = False, hasUnknown As Boolean = False
        For Each c As PdfAttachmentClassification In pdfClasses
            Select Case c
                Case PdfAttachmentClassification.DidiInvoicePdf, PdfAttachmentClassification.DidiTripPdf : hasDidi = True
                Case PdfAttachmentClassification.GeneralInvoicePdf : hasGeneral = True
                Case PdfAttachmentClassification.UnknownPdf : hasUnknown = True
            End Select
        Next

        If hasDidi Then Return MailProcessingType.DidiInvoiceMail
        If hasGeneral AndAlso hasUnknown Then Return MailProcessingType.MixedPdfMail
        If hasGeneral Then Return MailProcessingType.GeneralInvoiceMail
        If hasUnknown Then Return MailProcessingType.UnknownPdfOnlyMail
        Return MailProcessingType.NoPdfMail
    End Function

    Private Function ContainsAny(text As String, keywords As String()) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each k As String In keywords
            If text.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Next
        Return False
    End Function

End Class
