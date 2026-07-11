Imports System.IO
Imports System.Xml.Linq

''' <summary>
''' 百度 OCR 外部配置文件存储（%AppData%\iWorkHelper\baidu-ocr.config.xml）。
''' 特性：
'''  - **Secret Key 以 DPAPI（CurrentUser）加密保存**（带 DPAPI: 前缀），读时解密；
'''  - API Key/接口地址/开关等普通配置明文保存（非高度敏感）；
'''  - 不依赖 My.Settings，主程序与 OfflineTester 共用同一实现，避免两套规则；
'''  - 该文件属本机用户配置，已被 .gitignore 忽略，不得入库。
''' 用途：开发期把 OCR 配置安全落地到本机；主程序 OcrConfigProvider 在 My.Settings 未填时读取。
''' </summary>
Public Module BaiduXmlConfigStore

    Public Const ConfigFileName As String = "baidu-ocr.config.xml"

    Public Function GetPath() As String
        Return Path.Combine(PathHelper.GetAppDataRoot(), ConfigFileName)
    End Function

    Public Function Exists() As Boolean
        Try
            Return File.Exists(GetPath())
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 保存配置。plainSecret 为明文 Secret Key，写入前 DPAPI 加密。
    ''' </summary>
    Public Sub Save(options As BaiduOcrOptions, plainSecret As String)
        PathHelper.EnsureDirectory(PathHelper.GetAppDataRoot())
        Dim encSecret As String = SecretProtector.Protect(If(plainSecret, String.Empty))

        Dim doc As New XDocument(
            New XElement("BaiduOcr",
                New XElement("Enabled", options.Enabled.ToString()),
                New XElement("ApiKey", If(options.ApiKey, String.Empty)),
                New XElement("SecretKey", encSecret),
                New XElement("TokenUrl", If(options.TokenUrl, String.Empty)),
                New XElement("OcrApiUrl", If(options.OcrApiUrl, String.Empty)),
                New XElement("TimeoutMilliseconds", options.TimeoutMilliseconds.ToString()),
                New XElement("ReturnProbability", options.ReturnProbability.ToString()),
                New XElement("ReturnLocation", options.ReturnLocation.ToString()),
                New XElement("VerifyParameter", options.VerifyParameter.ToString()),
                New XElement("MaxPages", options.MaxPages.ToString()),
                New XElement("PreferLocalParse", options.PreferLocalParse.ToString()),
                New XElement("AutoFallbackToOcr", options.AutoFallbackToOcr.ToString())))

        doc.Save(GetPath())
        AppLogger.Info("OCR 外部配置已保存（Secret Key 已 DPAPI 加密）。")
    End Sub

    ''' <summary>
    ''' 读取配置。SecretKey 若为 DPAPI 密文则解密；文件不存在返回 Nothing。异常返回 Nothing。
    ''' </summary>
    Public Function Load() As BaiduOcrOptions
        Try
            Dim p As String = GetPath()
            If Not File.Exists(p) Then
                Return Nothing
            End If
            Dim root As XElement = XDocument.Load(p).Root
            If root Is Nothing Then
                Return Nothing
            End If

            Dim o As New BaiduOcrOptions()
            o.Enabled = ParseBool(GetValue(root, "Enabled"), False)
            o.ApiKey = GetValue(root, "ApiKey")

            Dim storedSecret As String = GetValue(root, "SecretKey")
            o.SecretKey = If(SecretProtector.Unprotect(storedSecret), String.Empty)

            Dim apiUrl As String = GetValue(root, "OcrApiUrl")
            If Not String.IsNullOrWhiteSpace(apiUrl) Then o.OcrApiUrl = apiUrl
            Dim tokenUrl As String = GetValue(root, "TokenUrl")
            If Not String.IsNullOrWhiteSpace(tokenUrl) Then o.TokenUrl = tokenUrl

            o.TimeoutMilliseconds = ParseInt(GetValue(root, "TimeoutMilliseconds"), 30000)
            o.ReturnProbability = ParseBool(GetValue(root, "ReturnProbability"), True)
            o.ReturnLocation = ParseBool(GetValue(root, "ReturnLocation"), False)
            o.VerifyParameter = ParseBool(GetValue(root, "VerifyParameter"), False)
            o.MaxPages = ParseInt(GetValue(root, "MaxPages"), 1)
            o.PreferLocalParse = ParseBool(GetValue(root, "PreferLocalParse"), True)
            o.AutoFallbackToOcr = ParseBool(GetValue(root, "AutoFallbackToOcr"), True)
            Return o
        Catch ex As Exception
            AppLogger.Error("读取 OCR 外部配置失败（忽略）。", ex)
            Return Nothing
        End Try
    End Function

    Private Function GetValue(root As XElement, name As String) As String
        Dim el As XElement = root.Element(name)
        If el Is Nothing Then Return Nothing
        Return el.Value
    End Function

    Private Function ParseBool(v As String, fallback As Boolean) As Boolean
        Dim b As Boolean
        If Boolean.TryParse(v, b) Then Return b
        Return fallback
    End Function

    Private Function ParseInt(v As String, fallback As Integer) As Integer
        Dim i As Integer
        If Integer.TryParse(v, i) Then Return i
        Return fallback
    End Function

End Module
