Imports System.Collections.Generic
Imports System.Text.RegularExpressions

''' <summary>
''' 常规发票字段候选评分与择优。评分原则（越高越可信）：
'''  - 标签就近：候选值与其字段标签（发票号码/开票日期/价税合计/销售方名称…）越近越高；
'''  - 区域合理：位于发票头部/购销区/合计区等预期区域加分；
'''  - 格式合理：日期完整、金额两位小数、号码长度合理加分；
'''  - 排除干扰：税额/税率不作金额、订单号/税号/校验码不作发票号码、购方名称不作销方名称。
''' 具体信号由识别器在建候选时提供，这里给出可复用的打分块与「每字段择优」。
''' </summary>
Public Module GeneralInvoiceCandidateScorer

    ''' <summary>组织机构名称特征词（用于销售方/购买方名称打分）。</summary>
    Private ReadOnly OrgKeywords As String() = {
        "有限公司", "股份", "公司", "有限责任", "经营部", "中心", "事务所", "合伙",
        "厂", "店", "超市", "商行", "商店", "工作室", "分公司", "集团", "医院", "学校", "银行", "酒店"}

    ''' <summary>号码长度合理性打分：数电票 20 位 / 传统 8 位最高，10~12 位（疑似发票代码）降分。</summary>
    Public Function InvoiceNumberPlausibility(value As String) As Double
        If String.IsNullOrEmpty(value) Then Return 0
        Dim n As Integer = value.Length
        If n = 20 Then Return 6 ' 数电票发票号码
        If n = 8 Then Return 5 ' 传统发票号码
        If n >= 16 AndAlso n <= 20 Then Return 3
        If n >= 10 AndAlso n <= 12 Then Return -2 ' 疑似发票代码
        If n >= 6 AndAlso n <= 20 Then Return 1
        Return -1
    End Function

    ''' <summary>
    ''' 长数字上下文惩罚：若所在行含「纳税人识别号/统一社会信用代码/校验码/订单号」等标签，
    ''' 说明该数字更可能是税号/校验码/订单号而非发票号码，据此降分。
    ''' </summary>
    Public Function LongNumberContextPenalty(lineText As String) As Double
        Dim t As String = If(lineText, "")
        Dim penalty As Double = 0
        If t.Contains("纳税人识别号") OrElse t.Contains("统一社会信用代码") Then penalty -= 8
        If t.Contains("校验码") Then penalty -= 8
        If t.Contains("订单号") OrElse t.Contains("流水号") Then penalty -= 6
        Return penalty
    End Function

    ''' <summary>发票代码长度合理性：10 或 12 位最高。</summary>
    Public Function InvoiceCodePlausibility(value As String) As Double
        If String.IsNullOrEmpty(value) Then Return 0
        Dim n As Integer = value.Length
        If n = 12 OrElse n = 10 Then Return 5
        If n >= 10 AndAlso n <= 12 Then Return 3
        Return -1
    End Function

    ''' <summary>日期格式完整度打分：含年月日/分隔符且年份合理加分。</summary>
    Public Function DateFormatBonus(value As String) As Double
        If String.IsNullOrEmpty(value) Then Return 0
        Dim ymd As String = LocalTextNormalizer.NormalizeDateToYmd(value)
        If String.IsNullOrEmpty(ymd) OrElse ymd.Length < 8 Then Return 0
        Dim year As Integer
        If Integer.TryParse(ymd.Substring(0, 4), year) AndAlso year >= 2000 AndAlso year <= 2099 Then Return 4
        Return 1
    End Function

    ''' <summary>组织名称特征打分：含「有限公司/公司/经营部/店…」加分；过短或纯数字降分。</summary>
    Public Function OrgNameBonus(value As String) As Double
        If String.IsNullOrWhiteSpace(value) Then Return -5
        Dim v As String = value.Trim()
        If v.Length < 3 Then Return -3
        If Regex.IsMatch(v, "^[\d\.\,\s¥￥%]+$") Then Return -5 ' 纯数字/金额，不是名称
        Dim bonus As Double = 0
        For Each k As String In OrgKeywords
            If v.Contains(k) Then bonus += 4 : Exit For
        Next
        If v.Length >= 4 AndAlso v.Length <= 40 Then bonus += 1
        Return bonus
    End Function

    ''' <summary>金额是否位于「价税合计/小写/¥」上下文：强加分；位于「税额/税率/不含税/合计(税前)」：降分。</summary>
    Public Function AmountContextBonus(lineText As String, isTotalWithTaxLabel As Boolean) As Double
        Dim t As String = If(lineText, "")
        Dim bonus As Double = 0
        If isTotalWithTaxLabel Then bonus += 6
        If t.Contains("价税合计") Then bonus += 5
        If t.Contains("小写") Then bonus += 4
        If t.Contains("¥") OrElse t.Contains("￥") Then bonus += 2
        If t.Contains("税额") AndAlso Not t.Contains("价税合计") Then bonus -= 4
        If t.Contains("税率") Then bonus -= 3
        Return bonus
    End Function

    ''' <summary>把「标签在同一行/相邻行」的距离转成就近加分（距离越小越高）。</summary>
    Public Function LabelProximityBonus(lineDistance As Integer) As Double
        If lineDistance <= 0 Then Return 6 ' 同一逻辑行
        If lineDistance = 1 Then Return 4
        If lineDistance = 2 Then Return 2
        If lineDistance <= 4 Then Return 1
        Return 0
    End Function

    ''' <summary>
    ''' 每字段择优：同名字段取评分最高者；分数相同取先出现者。
    ''' 返回 字段名 → 最优候选。
    ''' </summary>
    Public Function ChooseBest(candidates As List(Of GeneralInvoiceFieldCandidate)) As Dictionary(Of String, GeneralInvoiceFieldCandidate)
        Dim best As New Dictionary(Of String, GeneralInvoiceFieldCandidate)()
        If candidates Is Nothing Then Return best
        For Each c As GeneralInvoiceFieldCandidate In candidates
            If c Is Nothing OrElse String.IsNullOrWhiteSpace(c.FieldName) OrElse String.IsNullOrWhiteSpace(c.Value) Then Continue For
            Dim cur As GeneralInvoiceFieldCandidate = Nothing
            If Not best.TryGetValue(c.FieldName, cur) OrElse c.ConfidenceScore > cur.ConfidenceScore Then
                best(c.FieldName) = c
            End If
        Next
        Return best
    End Function

End Module
