Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web.Script.Serialization

''' <summary>
''' 百度 OCR HTTP 客户端。使用 HttpWebRequest（不引第三方库）调用智能财务票据识别接口。
''' 请求规范：
'''  - POST，Content-Type: application/x-www-form-urlencoded；
'''  - URL 参数 access_token；
'''  - Body 参数 pdf_file（PDF 二进制 Base64 后再 UrlEncode）、pdf_file_num（页码）；
'''  - 可选 probability / location / verify_parameter 按配置传入；
'''  - 不使用 image / url，不处理图片附件。
''' </summary>
Public Class BaiduOcrHttpClient

    ''' <summary>
    ''' 在 .NET Framework 4.8 下显式启用 TLS 1.2（叠加，不覆盖已有协议）。
    ''' </summary>
    Public Shared Sub EnsureTls12()
        Try
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12
        Catch
            ' 某些环境不支持设置时忽略。
        End Try
    End Sub

    ''' <summary>
    ''' 调用智能财务票据识别接口，识别指定页码。
    ''' </summary>
    ''' <param name="options">OCR 配置。</param>
    ''' <param name="accessToken">已获取的 access_token（不写日志）。</param>
    ''' <param name="pdfBytes">PDF 文件字节。</param>
    ''' <param name="pageNum">pdf_file_num 页码（从 1 开始）。</param>
    Public Function RecognizeMultipleInvoice(options As BaiduOcrOptions, accessToken As String, pdfBytes As Byte(), pageNum As Integer) As BaiduOcrRawResponse
        Dim resp As New BaiduOcrRawResponse With {.PageIndex = pageNum}

        Try
            EnsureTls12()

            Dim url As String = options.OcrApiUrl & "?access_token=" & Uri.EscapeDataString(accessToken)

            ' 组装 body：pdf_file 为 Base64 后再 UrlEncode。
            ' 注意：Uri.EscapeDataString 在 .NET Framework 下对超长字符串（约 65520 字符）会抛
            ' "URI 字符串太长"；PDF 的 Base64 远超该上限，故改用 WebUtility.UrlEncode（无长度限制）。
            Dim base64 As String = Convert.ToBase64String(pdfBytes)
            Dim sb As New StringBuilder()
            sb.Append("pdf_file=").Append(WebUtility.UrlEncode(base64))
            sb.Append("&pdf_file_num=").Append(pageNum.ToString())
            sb.Append("&probability=").Append(If(options.ReturnProbability, "true", "false"))
            sb.Append("&location=").Append(If(options.ReturnLocation, "true", "false"))
            sb.Append("&verify_parameter=").Append(If(options.VerifyParameter, "true", "false"))

            Dim bodyBytes As Byte() = Encoding.UTF8.GetBytes(sb.ToString())

            Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            request.Timeout = Math.Max(5000, options.TimeoutMilliseconds)
            request.ContentLength = bodyBytes.Length

            Using reqStream As Stream = request.GetRequestStream()
                reqStream.Write(bodyBytes, 0, bodyBytes.Length)
            End Using

            Try
                Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                    resp.HttpStatusCode = CInt(response.StatusCode)
                    resp.RawJson = ReadStream(response.GetResponseStream())
                End Using
            Catch wex As WebException
                resp.HttpStatusCode = ExtractStatus(wex)
                resp.RawJson = ReadWebExceptionBody(wex)
                If String.IsNullOrEmpty(resp.RawJson) Then
                    resp.Success = False
                    resp.NetworkError = wex.Message
                    AppLogger.Warn("OCR 请求网络错误（页 " & pageNum & "）：" & wex.Message & "，HTTP=" & resp.HttpStatusCode)
                    Return resp
                End If
            End Try

            resp.Success = True
            ExtractBaiduError(resp)
            AppLogger.Info("OCR 请求完成（页 " & pageNum & "）：HTTP=" & resp.HttpStatusCode &
                           If(resp.HasBaiduError, "，error_code=" & resp.ErrorCode, ""))
            Return resp

        Catch ex As Exception
            resp.Success = False
            resp.NetworkError = ExceptionFormatter.ToUserMessage(ex)
            AppLogger.Error("OCR 请求异常（页 " & pageNum & "）。", ex)
            Return resp
        End Try
    End Function

    ''' <summary>从原始 JSON 中提取 error_code / error_msg / pdf_file_size（如有）。</summary>
    Private Sub ExtractBaiduError(resp As BaiduOcrRawResponse)
        Try
            If String.IsNullOrEmpty(resp.RawJson) Then Return
            Dim serializer As New JavaScriptSerializer()
            Dim map As Dictionary(Of String, Object) = TryCast(serializer.DeserializeObject(resp.RawJson), Dictionary(Of String, Object))
            If map Is Nothing Then Return
            If map.ContainsKey("error_code") AndAlso map("error_code") IsNot Nothing Then
                resp.ErrorCode = Convert.ToString(map("error_code"))
            End If
            If map.ContainsKey("error_msg") AndAlso map("error_msg") IsNot Nothing Then
                resp.ErrorMsg = Convert.ToString(map("error_msg"))
            End If
        Catch
            ' 提取失败不影响主流程，交由解析器进一步处理。
        End Try
    End Sub

    Private Function ReadStream(s As Stream) As String
        Using sr As New StreamReader(s, Encoding.UTF8)
            Return sr.ReadToEnd()
        End Using
    End Function

    Private Function ReadWebExceptionBody(wex As WebException) As String
        Try
            If wex.Response IsNot Nothing Then
                Using sr As New StreamReader(wex.Response.GetResponseStream(), Encoding.UTF8)
                    Return sr.ReadToEnd()
                End Using
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Function ExtractStatus(wex As WebException) As Integer
        Try
            Dim resp As HttpWebResponse = TryCast(wex.Response, HttpWebResponse)
            If resp IsNot Nothing Then
                Return CInt(resp.StatusCode)
            End If
        Catch
        End Try
        Return 0
    End Function

End Class
