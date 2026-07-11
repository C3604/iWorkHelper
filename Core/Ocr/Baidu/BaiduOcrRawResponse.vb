''' <summary>
''' 百度 OCR 接口的原始响应封装。保留 HTTP 状态码、原始 JSON 与百度错误码/信息，
''' 供解析与诊断。敏感信息（token）不在此保存。
''' </summary>
Public Class BaiduOcrRawResponse

    ''' <summary>网络/HTTP 是否成功拿到响应体（不代表业务成功）。</summary>
    Public Property Success As Boolean

    ''' <summary>HTTP 状态码。</summary>
    Public Property HttpStatusCode As Integer

    ''' <summary>原始响应 JSON 文本。</summary>
    Public Property RawJson As String

    ''' <summary>百度错误码 error_code（如有）。</summary>
    Public Property ErrorCode As String

    ''' <summary>百度错误信息 error_msg（如有）。</summary>
    Public Property ErrorMsg As String

    ''' <summary>本次请求对应的页码（多页循环时使用）。</summary>
    Public Property PageIndex As Integer

    ''' <summary>网络层错误信息（如超时）。</summary>
    Public Property NetworkError As String

    Public ReadOnly Property HasBaiduError As Boolean
        Get
            Return Not String.IsNullOrEmpty(ErrorCode)
        End Get
    End Property

End Class
