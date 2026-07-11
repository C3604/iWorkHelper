Imports System.Collections.Generic
Imports System.Text.RegularExpressions

''' <summary>
''' 常规发票商品明细解析器。在检测出的明细区内逐行尽力解析，允许字段缺失：
'''  - 支持一行/多行明细；
'''  - 支持项目名称换行（无 * 前缀、且几乎无数字的行视为上一条名称的续行）；
'''  - 支持规格/单位/数量/单价为空；
'''  - 支持仅有「名称+金额+税率+税额」的简化明细。
''' 任何单行异常都被吞掉，不影响其它明细与发票头部字段。
''' </summary>
Public Class GeneralInvoiceLineItemParser

    ''' <summary>
    ''' 解析逻辑行序列中的商品明细。
    ''' 模型：每条明细以「*类别*名称」开头；不含 * 的行视为**上一条的续行**——
    ''' 既支持「名称换行」，也支持「名称在上一行、金额在续行」的换行版式。
    ''' </summary>
    Public Function Parse(lines As List(Of PdfTextLine)) As List(Of InvoiceLineItem)
        Dim items As New List(Of InvoiceLineItem)()
        Try
            Dim region As PdfTextBlock = PdfTableRegionDetector.DetectLineItemRegion(lines)
            If region Is Nothing OrElse region.IsEmpty Then Return items

            Dim current As InvoiceLineItem = Nothing
            Dim lineNo As Integer = 0
            For Each ln As PdfTextLine In region.Lines
                lineNo += 1
                Try
                    Dim raw As String = If(ln.Text, "").Trim()
                    If raw.Length = 0 OrElse LocalTextNormalizer.IsLikelyHeaderFooter(raw) Then Continue For
                    Dim hasStar As Boolean = raw.Contains("*")

                    If hasStar OrElse current Is Nothing Then
                        ' 新明细行（*类别*名称开头，或区首行）
                        current = New InvoiceLineItem With {.RawLine = raw, .LineIndex = lineNo, .PageNumber = ln.PageIndex}
                        current.ItemName = ExtractName(raw)
                        FillNumbers(current, raw, overwrite:=True)
                        If Not String.IsNullOrWhiteSpace(current.ItemName) OrElse Not String.IsNullOrWhiteSpace(current.Amount) Then
                            items.Add(current)
                        End If
                    Else
                        ' 续行：拼接名称，并补齐上一条尚缺的金额/税率/税额
                        Dim lead As String = LeadingName(raw)
                        If Not String.IsNullOrWhiteSpace(lead) Then
                            current.ItemName = (If(current.ItemName, "") & lead).Trim()
                        End If
                        current.RawLine = (If(current.RawLine, "") & " " & raw).Trim()
                        FillNumbers(current, raw, overwrite:=False)
                    End If
                Catch
                    ' 单行失败跳过
                End Try
            Next
        Catch ex As Exception
            AppLogger.Warn("常规发票明细解析异常（忽略，不影响头部字段）：" & ex.Message)
        End Try
        Return items
    End Function

    ''' <summary>取名称：优先 *类别*名称，否则取首个数字前的非数字串。</summary>
    Private Function ExtractName(raw As String) As String
        Dim mName As Match = Regex.Match(raw, "\*[^*]*\*\s*([^\d\r\n]{1,40})")
        If mName.Success Then Return mName.Groups(1).Value.Trim()
        Return LeadingName(raw)
    End Function

    ''' <summary>取行首到首个数字前的非数字串（用于续行名称拼接），去除尾随括号/标点。</summary>
    Private Function LeadingName(raw As String) As String
        Dim m As Match = Regex.Match(raw, "^\s*([^\d\r\n]*?)\s*\d")
        Dim s As String
        If m.Success Then
            s = m.Groups(1).Value
        Else
            s = Regex.Replace(raw, "[\d\.\,\s¥￥%]+$", "") ' 无数字：整行作名称
        End If
        Return s.Trim().TrimEnd("("c, "（"c, "-"c, "/"c, " "c)
    End Function

    ''' <summary>从一行补金额/税额/税率：末位金额=税额（有≥2个金额时），其前=金额；仅一个=金额。overwrite=False 时只补空缺。</summary>
    Private Sub FillNumbers(item As InvoiceLineItem, raw As String, overwrite As Boolean)
        Dim amounts As MatchCollection = Regex.Matches(raw, "\d+\.\d{2}")
        If amounts.Count >= 2 Then
            If overwrite OrElse String.IsNullOrEmpty(item.Amount) Then item.Amount = amounts(amounts.Count - 2).Value
            If overwrite OrElse String.IsNullOrEmpty(item.TaxAmount) Then item.TaxAmount = amounts(amounts.Count - 1).Value
        ElseIf amounts.Count = 1 Then
            If overwrite OrElse String.IsNullOrEmpty(item.Amount) Then item.Amount = amounts(0).Value
        End If
        Dim rate As String = LocalTextNormalizer.NormalizeTaxRate(raw)
        If Not String.IsNullOrEmpty(rate) AndAlso (overwrite OrElse String.IsNullOrEmpty(item.TaxRate)) Then item.TaxRate = rate
    End Sub

End Class
