''' <summary>
''' 单个 PDF 附件的分类结果：附件 + 分类 + 识别结果，供工作流分流。
''' </summary>
Public Class PdfAttachmentClassificationResult
    Public Property Attachment As MailAttachmentItem
    Public Property Classification As PdfAttachmentClassification
    Public Property Recognition As InvoiceRecognitionResult

    Public ReadOnly Property IsDidi As Boolean
        Get
            Return Classification = PdfAttachmentClassification.DidiInvoicePdf OrElse Classification = PdfAttachmentClassification.DidiTripPdf
        End Get
    End Property
End Class
