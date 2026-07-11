''' <summary>
''' 百度返回的票据 type 字段 → 内部 InvoiceDocumentType 映射。
''' 未知类型不报错，归为 Other/Unknown。
''' </summary>
Public Module BaiduInvoiceTypeMapper

    ''' <summary>
    ''' 映射类型字符串。
    ''' 已知：vat_invoice=增值税发票；taxi_online_ticket=网约车行程单；
    ''' taxi_receipt=出租车票；others/其它=Other。
    ''' </summary>
    Public Function Map(typeRaw As String) As InvoiceDocumentType
        If String.IsNullOrWhiteSpace(typeRaw) Then
            Return InvoiceDocumentType.Unknown
        End If

        Select Case typeRaw.Trim().ToLowerInvariant()
            Case "vat_invoice", "vat_special_invoice", "vat_electronic_invoice", "vat_electronic_normal_invoice", "vat_normal_invoice"
                Return InvoiceDocumentType.VatInvoice
            Case "taxi_online_ticket", "online_taxi_itinerary", "web_car_itinerary"
                Return InvoiceDocumentType.RideTripStatement
            Case "taxi_receipt"
                ' 出租车票：归入网约车/出行类，命名走通用出行分支。
                Return InvoiceDocumentType.RideTripStatement
            Case "others", "other"
                Return InvoiceDocumentType.Other
            Case Else
                Return InvoiceDocumentType.Other
        End Select
    End Function

    ''' <summary>是否为出行/行程类（用于命名分支）。</summary>
    Public Function IsRideType(typeRaw As String) As Boolean
        Return Map(typeRaw) = InvoiceDocumentType.RideTripStatement
    End Function

End Module
