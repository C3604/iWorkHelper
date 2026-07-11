''' <summary>
''' 常规发票商品明细的单行（增值税发票「货物或应税劳务、服务名称」表格中的一行）。
''' 所有字段以字符串保存原始文本，允许为空——不同发票列结构差异大，缺列不应导致解析失败。
''' </summary>
Public Class InvoiceLineItem
    ''' <summary>项目名称 / 货物或应税劳务名称（可能已去除 *类别* 前缀，也可能含换行拼接）。</summary>
    Public Property ItemName As String
    ''' <summary>规格型号（可空）。</summary>
    Public Property Specification As String
    ''' <summary>单位（可空）。</summary>
    Public Property Unit As String
    ''' <summary>数量（可空，原始文本）。</summary>
    Public Property Quantity As String
    ''' <summary>单价（可空，原始文本）。</summary>
    Public Property UnitPrice As String
    ''' <summary>金额（不含税，原始文本）。</summary>
    Public Property Amount As String
    ''' <summary>税率（如 3% / 6% / 13% / 免税，原始文本）。</summary>
    Public Property TaxRate As String
    ''' <summary>税额（可空，原始文本）。</summary>
    Public Property TaxAmount As String
    ''' <summary>该明细来源的原始行文本（诊断用）。</summary>
    Public Property RawLine As String
    ''' <summary>行号（从 1 开始，诊断用）。</summary>
    Public Property LineIndex As Integer
    ''' <summary>页码（从 1 开始，诊断用）。</summary>
    Public Property PageNumber As Integer
End Class
