''' <summary>
''' 百度 Access Token 获取结果（结构化，不直接抛到 UI）。
''' 注意：Token 属敏感信息，不得写入日志明文。
''' </summary>
Public Class BaiduAccessTokenResult

    ''' <summary>是否成功获取。</summary>
    Public Property Success As Boolean

    ''' <summary>access_token（仅内存使用，勿写日志）。</summary>
    Public Property AccessToken As String

    ''' <summary>有效期（秒），来自 expires_in。</summary>
    Public Property ExpiresInSeconds As Integer

    ''' <summary>错误信息（脱敏后可记录）。</summary>
    Public Property ErrorMessage As String

    ''' <summary>HTTP 状态码（如有）。</summary>
    Public Property HttpStatusCode As Integer

    ''' <summary>百度错误码（error / error_code 字段，如有）。</summary>
    Public Property BaiduError As String

    Public Shared Function Ok(token As String, expiresIn As Integer) As BaiduAccessTokenResult
        Return New BaiduAccessTokenResult With {.Success = True, .AccessToken = token, .ExpiresInSeconds = expiresIn}
    End Function

    Public Shared Function Fail(message As String, Optional httpStatus As Integer = 0, Optional baiduError As String = Nothing) As BaiduAccessTokenResult
        Return New BaiduAccessTokenResult With {.Success = False, .ErrorMessage = message, .HttpStatusCode = httpStatus, .BaiduError = baiduError}
    End Function

End Class
