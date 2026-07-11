''' <summary>
''' 命名模板配置。
''' 第八阶段起：**只使用一套“最终归档文件命名规则”（UnifiedTemplate）**，因为每封邮件最终只生成一个合并 PDF。
''' 旧的三套模板（Invoice/Trip/Unknown）字段保留仅为兼容旧配置，业务逻辑不再使用。
''' 未识别 fallback 名称由 ArchiveNamingRule 内置，不作为用户可配置模板暴露。
''' </summary>
Public Class NamingTemplates

    ''' <summary>滴滴统一最终命名模板默认值。</summary>
    Public Const DefaultUnifiedTemplate As String = "{乘车日期}_{金额}_{出发地点}_{到达地点}"

    ''' <summary>常规发票命名模板默认值。</summary>
    Public Const DefaultGeneralInvoiceTemplate As String = "{开票日期}_{金额}_{销售方名称}"

    ''' <summary>未识别文件命名模板默认值（保留源文件名）。</summary>
    Public Const DefaultUnknownTemplate As String = "{原始文件名}"

    ''' <summary>未识别 PDF 固定命名前缀（内部规则，不作为设置项）。</summary>
    Public Const UnknownPdfPrefix As String = "未识别_"

    ''' <summary>内置 fallback（不在设置页面暴露）：优先邮件主题，其次原附件名。</summary>
    Public Const FallbackWithSubject As String = "未识别票据_{邮件主题}_{时间戳}"
    Public Const FallbackWithOriginal As String = "未识别票据_{原附件名}_{时间戳}"

    ''' <summary>用户配置的滴滴统一最终命名模板。</summary>
    Public Property UnifiedTemplate As String = DefaultUnifiedTemplate

    ''' <summary>用户配置的常规发票命名模板。</summary>
    Public Property GeneralInvoiceTemplate As String = DefaultGeneralInvoiceTemplate

    ' —— 以下为旧三套模板，仅兼容旧配置，业务逻辑不再使用 ——
    <Obsolete("已合并为 UnifiedTemplate，仅兼容旧配置")>
    Public Property InvoiceTemplate As String
    <Obsolete("已合并为 UnifiedTemplate，仅兼容旧配置")>
    Public Property TripTemplate As String
    <Obsolete("已合并为 UnifiedTemplate，仅兼容旧配置")>
    Public Property UnknownTemplate As String

    Public Shared Function [Default]() As NamingTemplates
        Return New NamingTemplates()
    End Function

End Class
