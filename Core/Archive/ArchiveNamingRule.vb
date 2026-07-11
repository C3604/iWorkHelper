Imports System.Text
Imports System.Collections.Generic

''' <summary>
''' 归档命名方案（含所用规则名、缺失字段、未知占位符、是否触发 fallback），供归档结果标注。
''' </summary>
Public Class ArchiveNamePlan
    Public Property FileName As String
    Public Property RuleName As String
    Public Property FallbackTriggered As Boolean
    Public Sub New()
        MissingFields = New List(Of String)()
        UnknownPlaceholders = New List(Of String)()
    End Sub
    Public Property MissingFields As List(Of String)
    Public Property UnknownPlaceholders As List(Of String)
End Class

''' <summary>
''' 归档命名规则。第八阶段起：**只使用一套“最终归档文件命名规则”（UnifiedTemplate）**。
''' 默认模板：{乘车日期}_{金额}_{出发地点}_{到达地点}
''' 命名以合并后完整识别结果为准，业务变量取值优先级：
'''  - {乘车日期}：行程出发日期 &gt; 开票日期
'''  - {金额}：行程金额 &gt; 价税合计 &gt; 金额
'''  - {出发地点}=行程起点，{到达地点}=行程终点
''' 规则：空占位符跳过（不产生连续下划线）；未知占位符提示；清理非法字符；同名追加序号；
''' 渲染为空/字段严重不足 → 内置 fallback（未识别票据_邮件主题/原附件名_时间戳），不在设置页面暴露。
''' </summary>
Public Class ArchiveNamingRule

    Private ReadOnly _templates As NamingTemplates

    Public Sub New()
        _templates = NamingTemplates.Default()
    End Sub

    Public Sub New(templates As NamingTemplates)
        _templates = If(templates, NamingTemplates.Default())
    End Sub

    ''' <summary>
    ''' 生成命名方案（统一模板 + fallback）。
    ''' </summary>
    Public Function BuildPlan(invoice As InvoiceInfo, docType As InvoiceDocumentType, originalFileName As String, timestampToken As String,
                              Optional mailSubject As String = Nothing, Optional pdfCount As Integer = 0,
                              Optional recognitionSource As String = Nothing, Optional templateOverride As String = Nothing) As ArchiveNamePlan
        Dim values As Dictionary(Of String, String) = BuildValues(invoice, docType, originalFileName, timestampToken, mailSubject, pdfCount, recognitionSource)
        Dim originalBase As String = System.IO.Path.GetFileNameWithoutExtension(If(originalFileName, "未识别票据"))
        Dim fallbackBase As String = If(String.IsNullOrWhiteSpace(originalBase), "未识别票据", originalBase)

        ' 模板：显式覆盖 > 统一模板 > 默认模板（用于滴滴/常规发票不同命名）
        Dim tpl As String = templateOverride
        If String.IsNullOrWhiteSpace(tpl) Then tpl = _templates.UnifiedTemplate
        If String.IsNullOrWhiteSpace(tpl) Then tpl = NamingTemplates.DefaultUnifiedTemplate
        Dim render As NamingRenderResult = NamingTemplateEngine.Render(tpl, values, fallbackBase)

        Dim plan As New ArchiveNamePlan With {.RuleName = "统一模板"}
        plan.MissingFields.AddRange(render.MissingPlaceholders)
        plan.UnknownPlaceholders.AddRange(render.UnknownPlaceholders)
        If render.UnknownPlaceholders.Count > 0 Then
            AppLogger.Warn("命名模板含未知占位符：" & String.Join("、", render.UnknownPlaceholders.ToArray()))
        End If

        ' 渲染有效（非空、字段非严重不足）→ 采用
        If render.FileName IsNot Nothing AndAlso Not render.WasEmpty AndAlso Not String.IsNullOrWhiteSpace(render.FileName) Then
            plan.FileName = render.FileName
            Return plan
        End If

        ' fallback：未识别票据_{邮件主题}_{时间戳}；邮件主题不可用则用 {原附件名}
        plan.FallbackTriggered = True
        plan.RuleName = "未识别Fallback"
        plan.FileName = BuildUnrecognizedFallback(mailSubject, originalBase, timestampToken)
        AppLogger.Info("命名触发 fallback（字段严重不足）：" & PrivacySafeFormatter.MaskFileName(plan.FileName))
        Return plan
    End Function

    ''' <summary>兼容旧签名：仅返回文件名。</summary>
    Public Function BuildFileName(invoice As InvoiceInfo, originalFileName As String, timestampToken As String) As String
        Dim docType As InvoiceDocumentType = If(invoice IsNot Nothing AndAlso invoice.Trips IsNot Nothing AndAlso invoice.Trips.Count > 0,
                                                InvoiceDocumentType.RideTripStatement, InvoiceDocumentType.VatInvoice)
        Return BuildPlan(invoice, docType, originalFileName, timestampToken).FileName
    End Function

    ''' <summary>构建全部已知占位符的取值表（值可为空）。含业务统一变量。</summary>
    Private Function BuildValues(invoice As InvoiceInfo, docType As InvoiceDocumentType, originalFileName As String, timestampToken As String,
                                 Optional mailSubject As String = Nothing, Optional pdfCount As Integer = 0,
                                 Optional recognitionSource As String = Nothing) As Dictionary(Of String, String)
        Dim v As New Dictionary(Of String, String)()
        Dim trip As InvoiceTripInfo = Nothing
        If invoice IsNot Nothing AndAlso invoice.Trips IsNot Nothing AndAlso invoice.Trips.Count > 0 Then
            trip = invoice.Trips(0) ' 命名默认使用首条行程
        End If

        Dim originalBase As String = System.IO.Path.GetFileNameWithoutExtension(If(originalFileName, String.Empty))

        ' —— 发票字段 ——
        v("开票日期") = NormalizeDate(GetProp(invoice, Function() invoice.InvoiceDate))
        v("销售方名称") = GetProp(invoice, Function() invoice.SellerName)
        v("购买方名称") = GetProp(invoice, Function() invoice.BuyerName)
        v("发票号码") = GetProp(invoice, Function() invoice.InvoiceNumber)
        v("发票代码") = GetProp(invoice, Function() invoice.InvoiceCode)
        v("价税合计") = GetProp(invoice, Function() invoice.TotalWithTax)
        v("税额") = GetProp(invoice, Function() invoice.TaxAmount)
        v("税率") = GetProp(invoice, Function() invoice.TaxRate)
        v("开票人") = GetProp(invoice, Function() invoice.Drawer)
        v("项目名称") = GetProp(invoice, Function() invoice.ItemName)
        v("票据类型") = DocTypeText(docType)

        ' —— 行程字段 ——
        Dim tripDate As String = Nothing
        If invoice IsNot Nothing AndAlso invoice.ExtendedFields IsNot Nothing AndAlso invoice.ExtendedFields.ContainsKey("行程起止日期") Then
            tripDate = invoice.ExtendedFields("行程起止日期")
        End If
        If String.IsNullOrWhiteSpace(tripDate) AndAlso trip IsNot Nothing Then tripDate = trip.DepartureTime

        Dim startLoc As String = If(trip IsNot Nothing, trip.StartLocation, Nothing)
        Dim endLoc As String = If(trip IsNot Nothing, trip.EndLocation, Nothing)
        Dim tripAmount As String = If(trip IsNot Nothing, trip.TripAmount, Nothing)

        v("出发日期") = NormalizeDate(tripDate)
        v("出发时间") = If(trip IsNot Nothing, trip.DepartureTime, Nothing)
        v("到达时间") = If(trip IsNot Nothing, trip.ArrivalTime, Nothing)
        v("乘车人") = If(trip IsNot Nothing, trip.Passenger, Nothing)
        v("起点") = startLoc
        v("终点") = endLoc
        v("出发地点") = startLoc   ' {出发地点} 映射到行程起点
        v("到达地点") = endLoc     ' {到达地点} 映射到行程终点
        v("城市") = If(trip IsNot Nothing, trip.City, Nothing)
        v("车型") = If(trip IsNot Nothing, trip.ServiceType, Nothing)
        v("订单号") = If(trip IsNot Nothing, trip.OrderNumber, Nothing)
        v("里程") = If(trip IsNot Nothing, trip.Mileage, Nothing)
        v("行程金额") = If(Not String.IsNullOrWhiteSpace(tripAmount), tripAmount, GetProp(invoice, Function() invoice.TotalWithTax))

        ' —— 业务统一变量 ——
        ' {乘车日期}：行程出发日期 > 开票日期
        v("乘车日期") = FirstNonEmpty(NormalizeDate(tripDate), NormalizeDate(GetProp(invoice, Function() invoice.InvoiceDate)))
        ' {金额}：行程金额 > 价税合计 > 金额
        v("金额") = FirstNonEmpty(tripAmount, GetProp(invoice, Function() invoice.TotalWithTax), GetProp(invoice, Function() invoice.Amount))

        ' —— 通用变量 ——
        Dim mergedScenario As Boolean = pdfCount > 1
        v("原附件名") = If(mergedScenario AndAlso Not String.IsNullOrWhiteSpace(mailSubject), mailSubject.Trim(), originalBase)
        v("时间戳") = If(timestampToken, String.Empty)
        v("邮件主题") = If(mailSubject, String.Empty)
        v("附件数量") = If(pdfCount > 0, pdfCount.ToString(), String.Empty)
        v("合并文件名") = originalBase
        v("识别来源") = If(recognitionSource, String.Empty)
        ' {原始文件名}：当前 PDF 附件原始名（不含扩展名），用于常规发票/未识别命名
        v("原始文件名") = originalBase

        Return v
    End Function

    ''' <summary>未识别 fallback：优先 未识别票据_邮件主题_时间戳；无邮件主题则用原附件名。</summary>
    Private Function BuildUnrecognizedFallback(mailSubject As String, originalBase As String, timestampToken As String) As String
        Dim parts As New List(Of String)()
        parts.Add("未识别票据")
        If Not String.IsNullOrWhiteSpace(mailSubject) Then
            parts.Add(mailSubject.Trim())
        ElseIf Not String.IsNullOrWhiteSpace(originalBase) Then
            parts.Add(originalBase)
        End If
        If Not String.IsNullOrWhiteSpace(timestampToken) Then parts.Add(timestampToken)
        Return JoinToFileName(parts, "未识别票据")
    End Function

    Private Function DocTypeText(docType As InvoiceDocumentType) As String
        Select Case docType
            Case InvoiceDocumentType.VatInvoice : Return "增值税发票"
            Case InvoiceDocumentType.RideTripStatement : Return "行程单"
            Case InvoiceDocumentType.Other : Return "其它"
            Case Else : Return "未知"
        End Select
    End Function

    ''' <summary>命名模板支持的占位符清单（与设置页“变量说明”一致的单一来源）。业务变量在前。</summary>
    Public Shared Function SupportedPlaceholders() As String()
        Return New String() {
            "乘车日期", "金额", "出发地点", "到达地点",
            "邮件主题", "原附件名", "原始文件名", "时间戳", "票据类型", "识别来源", "附件数量", "合并文件名",
            "出发日期", "出发时间", "到达时间", "乘车人", "起点", "终点", "城市", "车型", "订单号", "里程", "行程金额",
            "开票日期", "销售方名称", "购买方名称", "发票号码", "发票代码", "价税合计", "税额", "税率", "开票人", "项目名称"}
    End Function

    Private Function FirstNonEmpty(ParamArray candidates As String()) As String
        For Each c As String In candidates
            If Not String.IsNullOrWhiteSpace(c) Then Return c
        Next
        Return Nothing
    End Function

    Private Function GetProp(invoice As InvoiceInfo, getter As Func(Of String)) As String
        If invoice Is Nothing Then Return Nothing
        Try
            Return getter()
        Catch
            Return Nothing
        End Try
    End Function

    Private Function JoinToFileName(parts As List(Of String), fallback As String) As String
        If parts Is Nothing OrElse parts.Count = 0 Then
            Return FileNameSanitizer.BuildFileName(fallback, ".pdf", fallback)
        End If
        Dim raw As String = String.Join("_", parts.ToArray())
        Return FileNameSanitizer.BuildFileName(raw, ".pdf", fallback)
    End Function

    Private Function NormalizeDate(rawDate As String) As String
        If String.IsNullOrWhiteSpace(rawDate) Then
            Return Nothing
        End If
        Dim sb As New StringBuilder()
        For Each c As Char In rawDate
            If Char.IsDigit(c) Then
                sb.Append(c)
            End If
        Next
        Dim digits As String = sb.ToString()
        If digits.Length >= 8 Then
            Return digits.Substring(0, 8)
        End If
        Return rawDate.Trim()
    End Function

End Class
