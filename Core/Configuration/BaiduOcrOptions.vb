''' <summary>
''' 百度智能云 OCR 配置项（纯数据对象，不依赖 My.Settings，便于离线测试复用）。
''' 严禁在代码中硬编码 API Key / Secret Key，实际值来自 SettingsForm 写入的 My.Settings
''' （经 OcrConfigProvider 读取）或兼容的外部 XML。
''' </summary>
Public Class BaiduOcrOptions

    ''' <summary>是否启用在线 OCR。默认 False（未配置即视为禁用）。</summary>
    Public Property Enabled As Boolean = False

    ''' <summary>百度 OCR API Key（应用的 AK）。</summary>
    Public Property ApiKey As String

    ''' <summary>百度 OCR Secret Key（应用的 SK）。</summary>
    Public Property SecretKey As String

    ''' <summary>获取 Access Token 的鉴权地址。</summary>
    Public Property TokenUrl As String = "https://aip.baidubce.com/oauth/2.0/token"

    ''' <summary>
    ''' OCR 识别接口地址。默认智能财务票据识别（多票据）接口。
    ''' </summary>
    Public Property OcrApiUrl As String = "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice"

    ''' <summary>请求超时时间（毫秒）。默认 30 秒。</summary>
    Public Property TimeoutMilliseconds As Integer = 30000

    ''' <summary>是否返回字段置信度 probability。默认 True。</summary>
    Public Property ReturnProbability As Boolean = True

    ''' <summary>是否返回字段位置 location。默认 False。</summary>
    Public Property ReturnLocation As Boolean = False

    ''' <summary>是否开启验真参数 verify_parameter。默认 False。</summary>
    Public Property VerifyParameter As Boolean = False

    ''' <summary>最大识别页数。默认 1（仅识别第 1 页）。</summary>
    Public Property MaxPages As Integer = 1

    ''' <summary>是否优先使用本地文本解析。默认 True。</summary>
    Public Property PreferLocalParse As Boolean = True

    ''' <summary>本地解析失败或字段不足时是否自动调用在线 OCR。默认 True。</summary>
    Public Property AutoFallbackToOcr As Boolean = True

    ''' <summary>
    ''' 判断必要配置是否齐全（启用且 AK/SK/接口地址均非空）。
    ''' </summary>
    Public Function IsConfigured() As Boolean
        Return Enabled _
            AndAlso Not String.IsNullOrWhiteSpace(ApiKey) _
            AndAlso Not String.IsNullOrWhiteSpace(SecretKey) _
            AndAlso Not String.IsNullOrWhiteSpace(OcrApiUrl) _
            AndAlso Not String.IsNullOrWhiteSpace(TokenUrl)
    End Function

    ''' <summary>
    ''' 生成脱敏摘要（供日志），绝不输出完整 AK/SK。
    ''' </summary>
    Public Function ToSafeSummary() As String
        Return String.Format("Enabled={0}, ApiKey={1}, SecretKey={2}, OcrApiUrl={3}, MaxPages={4}, Probability={5}, Location={6}, Verify={7}, PreferLocal={8}, AutoFallback={9}",
                             Enabled, MaskSecret(ApiKey), MaskSecret(SecretKey), OcrApiUrl, MaxPages,
                             ReturnProbability, ReturnLocation, VerifyParameter, PreferLocalParse, AutoFallbackToOcr)
    End Function

    ''' <summary>
    ''' 脱敏：仅保留前 2、后 2 字符，中间用 * 代替。用于日志/展示，绝不输出完整密钥。
    ''' </summary>
    Public Shared Function MaskSecret(value As String) As String
        If String.IsNullOrEmpty(value) Then
            Return "(空)"
        End If
        If value.Length <= 4 Then
            Return New String("*"c, value.Length)
        End If
        Return value.Substring(0, 2) & New String("*"c, Math.Max(1, value.Length - 4)) & value.Substring(value.Length - 2)
    End Function

End Class
