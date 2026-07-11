''' <summary>邮件级处理类型（决定走哪条归档链路）。</summary>
Public Enum MailProcessingType
    ''' <summary>滴滴发票邮件：走现有“发票+行程单合并”归档。</summary>
    DidiInvoiceMail = 0
    ''' <summary>常规发票邮件：每个发票 PDF 单独识别、命名、归档。</summary>
    GeneralInvoiceMail = 1
    ''' <summary>常规发票 + 未识别 PDF 混合邮件：分别处理。</summary>
    MixedPdfMail = 2
    ''' <summary>无 PDF 附件：直接跳过。</summary>
    NoPdfMail = 3
    ''' <summary>仅含未识别 PDF：每个 PDF 按“未识别_原名”归档。</summary>
    UnknownPdfOnlyMail = 4
End Enum

''' <summary>PDF 附件级分类。</summary>
Public Enum PdfAttachmentClassification
    ''' <summary>滴滴发票 PDF（含发票信息 + 滴滴特征）。</summary>
    DidiInvoicePdf = 0
    ''' <summary>滴滴行程单 PDF（含行程信息、无发票号码）。</summary>
    DidiTripPdf = 1
    ''' <summary>常规发票 PDF。</summary>
    GeneralInvoicePdf = 2
    ''' <summary>无法识别的 PDF。</summary>
    UnknownPdf = 3
End Enum
