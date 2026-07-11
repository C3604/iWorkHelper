Imports System.IO
Imports System.Xml.Linq

''' <summary>
''' OCR 配置读取器。第三阶段起 **以 SettingsForm 写入的 My.Settings 为主要来源**。
''' 配置来源优先级：
'''  1) My.Settings（SettingsForm 设置界面写入）——主入口；
'''  2) 若 My.Settings 中未填 AK/SK，则尝试兼容旧的外部 XML 文件
'''     %AppData%\iWorkHelper\baidu-ocr.config.xml（历史遗留，可选）。
''' 任何异常都回退为禁用配置，不抛出。密钥不写入日志。
''' </summary>
Public Module OcrConfigProvider

    Public Const ConfigFileName As String = "baidu-ocr.config.xml"

    ''' <summary>
    ''' 读取百度 OCR 配置（主入口）。
    ''' </summary>
    Public Function LoadBaiduOptions() As BaiduOcrOptions
        Dim options As BaiduOcrOptions
        Try
            options = LoadFromSettings()
        Catch ex As Exception
            AppLogger.Error("从 My.Settings 读取 OCR 配置失败，尝试回退 XML。", ex)
            options = New BaiduOcrOptions()
        End Try

        ' 若设置中未提供 AK/SK，则以外部加密配置文件（BaiduXmlConfigStore）作为完整来源。
        ' 该文件的 Secret Key 为 DPAPI 加密，读取时解密。用于"开发期配置落地"。
        If String.IsNullOrWhiteSpace(options.ApiKey) OrElse String.IsNullOrWhiteSpace(options.SecretKey) Then
            Try
                Dim fromXml As BaiduOcrOptions = BaiduXmlConfigStore.Load()
                If fromXml IsNot Nothing Then
                    options = fromXml ' 完整采用外部配置（含 Enabled/URL/开关等）
                End If
            Catch ex As Exception
                AppLogger.Error("读取外部 OCR 配置失败（忽略）。", ex)
            End Try
        End If

        AppLogger.Info("OCR 配置已加载：" & options.ToSafeSummary())
        Return options
    End Function

    ''' <summary>
    ''' 从 My.Settings 构建配置对象。
    ''' </summary>
    Private Function LoadFromSettings() As BaiduOcrOptions
        Dim s = My.Settings
        Dim o As New BaiduOcrOptions()
        o.Enabled = s.OcrEnabled
        o.ApiKey = If(s.BaiduApiKey, String.Empty).Trim()
        ' Secret Key 经 DPAPI 加密存储，此处解密使用（解密失败返回空，不阻断）。
        o.SecretKey = If(ProtectedSettingsProvider.GetSecretKey(), String.Empty).Trim()
        o.OcrApiUrl = If(String.IsNullOrWhiteSpace(s.BaiduOcrApiUrl), "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice", s.BaiduOcrApiUrl.Trim())
        o.TokenUrl = If(String.IsNullOrWhiteSpace(s.BaiduTokenUrl), "https://aip.baidubce.com/oauth/2.0/token", s.BaiduTokenUrl.Trim())
        o.TimeoutMilliseconds = If(s.OcrTimeoutMs > 0, s.OcrTimeoutMs, 30000)
        o.ReturnProbability = s.OcrReturnProbability
        o.ReturnLocation = s.OcrReturnLocation
        o.VerifyParameter = s.OcrVerifyParameter
        o.MaxPages = If(s.OcrMaxPages > 0, s.OcrMaxPages, 1)
        o.PreferLocalParse = s.PreferLocalParse
        o.AutoFallbackToOcr = s.AutoFallbackToOcr
        Return o
    End Function

    Public Function GetConfigFilePath() As String
        Return BaiduXmlConfigStore.GetPath()
    End Function

End Module
