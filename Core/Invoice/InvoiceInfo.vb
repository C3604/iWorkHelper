Imports System.Collections.Generic

''' <summary>
''' 发票信息模型。覆盖常见增值税电子发票字段，并预留行程单字段与可扩展字段字典，
''' 以便承载百度 OCR 返回的、模型未显式定义的额外字段（避免字段丢失）。
''' 说明：所有金额/数量类字段均以字符串保存原始文本，不在模型层强制数值转换，
''' 以免因票据格式差异导致解析失败或精度问题。
''' </summary>
Public Class InvoiceInfo

    Public Sub New()
        ExtendedFields = New Dictionary(Of String, String)()
        Trips = New List(Of InvoiceTripInfo)()
        LineItems = New List(Of InvoiceLineItem)()
    End Sub

    ' —— 增值税电子发票通用字段 ——
    Public Property InvoiceCode As String            ' 发票代码
    Public Property InvoiceNumber As String          ' 发票号码
    Public Property InvoiceDate As String            ' 开票日期（原始文本）
    Public Property BuyerName As String              ' 购买方名称
    Public Property BuyerTaxId As String             ' 购买方纳税人识别号
    Public Property SellerName As String             ' 销售方名称
    Public Property SellerTaxId As String            ' 销售方纳税人识别号
    Public Property ItemName As String               ' 项目名称
    Public Property Specification As String          ' 规格型号
    Public Property Unit As String                   ' 单位
    Public Property Quantity As String               ' 数量
    Public Property UnitPrice As String              ' 单价
    Public Property Amount As String                 ' 金额（不含税）
    Public Property TaxRate As String                ' 税率
    Public Property TaxAmount As String              ' 税额
    Public Property TotalWithTax As String           ' 价税合计
    Public Property Remark As String                 ' 备注
    Public Property CheckCode As String              ' 校验码
    Public Property Payee As String                  ' 收款人
    Public Property Reviewer As String               ' 复核人
    Public Property Drawer As String                 ' 开票人

    ''' <summary>行程明细（滴滴/网约车行程单，可含多条）。</summary>
    Public Property Trips As List(Of InvoiceTripInfo)

    ''' <summary>常规发票商品明细（可含多条；解析失败或无明细时为空列表，不影响头部字段）。</summary>
    Public Property LineItems As List(Of InvoiceLineItem)

    ''' <summary>行程单声明的行程笔数（来自"共N笔行程"，可能与实际解析出的 Trips.Count 不同）。</summary>
    Public Property StatedTripCount As Integer

    ''' <summary>
    ''' 扩展/未知字段：保存模型未显式定义的字段，key 为字段名，value 为原始文本。
    ''' 用于承载百度 OCR 返回的额外字段，避免信息丢失。
    ''' </summary>
    Public Property ExtendedFields As Dictionary(Of String, String)

    ''' <summary>
    ''' 写入一个扩展字段（key 为空则忽略）。
    ''' </summary>
    Public Sub SetExtendedField(name As String, value As String)
        If String.IsNullOrEmpty(name) Then
            Return
        End If
        ExtendedFields(name) = If(value, String.Empty)
    End Sub

    ''' <summary>
    ''' 判断是否已提取到任何关键字段，用于评估解析是否“部分成功”。
    ''' </summary>
    Public Function HasAnyCoreField() As Boolean
        Return Not String.IsNullOrWhiteSpace(InvoiceNumber) _
            OrElse Not String.IsNullOrWhiteSpace(InvoiceCode) _
            OrElse Not String.IsNullOrWhiteSpace(TotalWithTax) _
            OrElse Not String.IsNullOrWhiteSpace(SellerName) _
            OrElse Not String.IsNullOrWhiteSpace(InvoiceDate) _
            OrElse (Trips IsNot Nothing AndAlso Trips.Count > 0)
    End Function

End Class
