''' <summary>
''' 单个识别字段。除字段名与值外，携带置信度与来源，便于诊断本地解析与在线 OCR 的差异。
''' </summary>
Public Class InvoiceField

    Public Sub New()
    End Sub

    Public Sub New(name As String, value As String, Optional source As String = "", Optional confidence As Double = 1.0)
        Me.Name = name
        Me.Value = value
        Me.Source = source
        Me.Confidence = confidence
    End Sub

    ''' <summary>字段名，建议使用 InvoiceFieldNames 中的常量。</summary>
    Public Property Name As String

    ''' <summary>字段值（原始文本）。</summary>
    Public Property Value As String

    ''' <summary>来源，如 "LocalText" 或 "BaiduOcr"。</summary>
    Public Property Source As String

    ''' <summary>置信度 0~1。本地正则解析默认 1.0；OCR 可回填真实置信度。</summary>
    Public Property Confidence As Double

End Class
