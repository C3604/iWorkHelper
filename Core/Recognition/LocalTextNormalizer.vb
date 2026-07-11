Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' 本地解析前的文本归一化。目标：降低因全/半角、冒号、空格、换行、金额符号差异导致的正则失配，
''' 同时**保留中文地址、括号、数字、金额等必要信息**（不过度清洗）。
''' </summary>
Public Module LocalTextNormalizer

    ''' <summary>
    ''' 归一化文本：全角转半角（仅 ASCII 区）、冒号统一为半角、空格折叠、换行统一、金额符号保留但统一。
    ''' </summary>
    Public Function Normalize(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text

        Dim sb As New StringBuilder(text.Length)
        For Each c As Char In text
            Dim code As Integer = AscW(c)
            ' 全角 ASCII（FF01-FF5E）转半角
            If code >= &HFF01 AndAlso code <= &HFF5E Then
                sb.Append(ChrW(code - &HFEE0))
            ElseIf code = &H3000 Then ' 全角空格 → 半角空格
                sb.Append(" "c)
            Else
                sb.Append(c)
            End If
        Next

        Dim s As String = sb.ToString()

        ' 冒号统一为半角（中文冒号、全角冒号已在上面转过，这里兜底）
        s = s.Replace("："c, ":"c)
        ' 换行统一
        s = s.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        ' 金额符号统一：全角￥(FFE5) 与半角¥ 统一为 ¥（保留，便于金额正则）
        s = s.Replace("￥", "¥")
        ' 多个空格/制表符折叠为单个空格（不动换行）
        s = Regex.Replace(s, "[ \t]{2,}", " ")

        Return s
    End Function

    ''' <summary>
    ''' 把各种日期格式统一规整为 8 位数字 YYYYMMDD；无法识别返回 Nothing。
    ''' 支持 yyyy-MM-dd / yyyy/MM/dd / yyyy.MM.dd / yyyy年MM月dd日。
    ''' </summary>
    Public Function NormalizeDateToYmd(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return Nothing
        Dim m As Match = Regex.Match(raw, "(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})")
        If m.Success Then
            Return m.Groups(1).Value & Pad2(m.Groups(2).Value) & Pad2(m.Groups(3).Value)
        End If
        ' 退化：抓前 8 位数字
        Dim digits As String = Regex.Replace(raw, "\D", "")
        If digits.Length >= 8 Then Return digits.Substring(0, 8)
        Return Nothing
    End Function

    Private Function Pad2(s As String) As String
        If s.Length = 1 Then Return "0" & s
        Return s
    End Function

    ''' <summary>
    ''' 从原始文本中提取金额候选（两位小数），兼容 ¥/￥/元/带千分位逗号。
    ''' 返回规范化的「数字.两位小数」；无法识别返回 Nothing。不做汇率/大写转换。
    ''' 示例：¥123.45→123.45，1,234.50元→1234.50，123.4→Nothing（要求两位小数）。
    ''' </summary>
    Public Function NormalizeAmount(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return Nothing
        ' 去掉货币符号与千分位逗号后再匹配两位小数金额
        Dim cleaned As String = raw.Replace("¥", "").Replace("￥", "").Replace(",", "").Replace("，", "")
        Dim m As Match = Regex.Match(cleaned, "(\d+\.\d{2})(?!\d)")
        If m.Success Then Return m.Groups(1).Value
        Return Nothing
    End Function

    ''' <summary>
    ''' 归一化税率：13%/9%/6%/3%/1%/0%，以及「免税/不征税/免/***」等特殊值原样保留。
    ''' 无法识别返回 Nothing。
    ''' </summary>
    Public Function NormalizeTaxRate(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return Nothing
        If raw.Contains("免税") Then Return "免税"
        If raw.Contains("不征税") Then Return "不征税"
        Dim m As Match = Regex.Match(raw, "(\d{1,2})\s*%")
        If m.Success Then Return m.Groups(1).Value & "%"
        Return Nothing
    End Function

    ''' <summary>
    ''' 是否疑似页眉/页脚等干扰行（如「第 X 页 共 Y 页」「打印日期」）。
    ''' 仅用于常规发票逐行解析时的过滤，保守判断，避免误删业务行。
    ''' </summary>
    Public Function IsLikelyHeaderFooter(line As String) As Boolean
        If String.IsNullOrWhiteSpace(line) Then Return True
        Dim t As String = line.Trim()
        If Regex.IsMatch(t, "^第?\s*\d+\s*页(\s*/\s*|\s*共\s*)\d+\s*页?$") Then Return True
        If Regex.IsMatch(t, "^共\s*\d+\s*页") Then Return True
        Return False
    End Function

End Module
