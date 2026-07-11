''' <summary>
''' 常规发票字段候选：同一字段可能在票面出现多次（如多个金额、多处名称），
''' 先收集全部候选并各自打分，再择优，避免「第一个匹配就采用」造成误取。
''' </summary>
Public Class GeneralInvoiceFieldCandidate
    ''' <summary>字段名（用 InvoiceFieldNames 常量，如 发票号码/销售方名称/价税合计）。</summary>
    Public Property FieldName As String
    ''' <summary>清洗后的字段值。</summary>
    Public Property Value As String
    ''' <summary>候选来源的原始行文本（诊断用）。</summary>
    Public Property RawText As String
    ''' <summary>页码（从 1 开始）。</summary>
    Public Property PageNumber As Integer
    ''' <summary>逻辑行号（从 0 开始）。</summary>
    Public Property LineIndex As Integer
    ''' <summary>X 坐标（有坐标时；否则 0）。</summary>
    Public Property X As Double
    ''' <summary>Y 坐标（有坐标时；否则 0）。</summary>
    Public Property Y As Double
    ''' <summary>综合评分（越高越可信）。</summary>
    Public Property ConfidenceScore As Double
    ''' <summary>产生该候选的规则标识（诊断用）。</summary>
    Public Property SourceRule As String
    ''' <summary>评分理由（诊断用，人可读）。</summary>
    Public Property Reason As String

    Public Overrides Function ToString() As String
        Return String.Format("{0}='{1}' 分={2:F1} [{3}] p{4}L{5} {6}",
                             FieldName, Value, ConfidenceScore, SourceRule, PageNumber, LineIndex, Reason)
    End Function
End Class
