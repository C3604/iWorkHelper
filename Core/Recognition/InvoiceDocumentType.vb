''' <summary>
''' 票据文档类型。用于区分处理分支（增值税发票 / 网约车行程单等）。
''' </summary>
Public Enum InvoiceDocumentType
    ''' <summary>未识别/未知类型。</summary>
    Unknown = 0
    ''' <summary>增值税电子普通/专用发票。</summary>
    VatInvoice = 1
    ''' <summary>网约车（滴滴等）行程单。</summary>
    RideTripStatement = 2
    ''' <summary>其它票据。</summary>
    Other = 99
End Enum
