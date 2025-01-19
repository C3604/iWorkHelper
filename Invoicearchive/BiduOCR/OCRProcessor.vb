Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Web
Imports Newtonsoft.Json

''' <summary>
''' OCR 处理模块，用于调用 OCR 服务识别文档内容。
''' </summary>
Public Module OCRProcessor
    ' OCR 服务相关配置
    Private ReadOnly ClientId As String = My.Settings.ApiKey
    Private ReadOnly ClientSecret As String = My.Settings.SecretKey
    Private ReadOnly OCRUrl As String = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic"
    Private ReadOnly TokenUrl As String = "https://aip.baidubce.com/oauth/2.0/token"

    ''' <summary>
    ''' 处理指定文件夹中的 PDF 文件，执行 OCR 并生成结果。
    ''' </summary>
    ''' <param name="tempFolderPath">临时文件夹路径。</param>
    ''' <returns>返回处理状态：1（成功处理指定文件），2（成功处理其他文件），0（无可处理文件），-1（错误）。</returns>
    Public Function ProcessOCR(tempFolderPath As String ) As Integer
        Try
            Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", $"开始处理 OCR，文件夹路径: {tempFolderPath}")

            ' 获取 Access Token
            Dim accessToken As String = GetAccessToken()
            If String.IsNullOrEmpty(accessToken) Then
                Logger.WriteLog(Logger.LogLevel.Error, "OCRProcessor", "无法获取 Access Token，OCR 处理中止。")
                Return -1
            End If

            ' 优先处理特定文件名的 PDF
            Dim pdfFilePath As String = Path.Combine(tempFolderPath, "滴滴出行行程报销单.pdf")
            If File.Exists(pdfFilePath) Then
                Dim ocrResult As String = PerformOCR(pdfFilePath, accessToken)
                SaveOCRResult(tempFolderPath, ocrResult)
                Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", $"成功处理文件: {pdfFilePath}")
                Return 1
            End If

            ' 处理其他 PDF 文件
            For Each filePath As String In Directory.GetFiles(tempFolderPath, "*.pdf")
                Dim ocrResult As String = PerformOCR(filePath, accessToken)
                SaveOCRResult(tempFolderPath, ocrResult)
                Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", $"成功处理文件: {filePath}")
                Return 2
            Next

            Logger.WriteLog(Logger.LogLevel.Warning, "OCRProcessor", "文件夹中没有需要处理的 PDF 文件。")
            Return 0
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "OCRProcessor", $"处理 OCR 时发生错误: {ex.Message}")
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' 调用 OCR 服务处理 PDF 文件。
    ''' </summary>
    ''' <param name="pdfFilePath">PDF 文件路径。</param>
    ''' <param name="accessToken">OCR 服务的 Access Token。</param>
    ''' <returns>OCR 服务返回的结果字符串。</returns>
    Private Function PerformOCR(pdfFilePath As String, accessToken As String) As String
        Try
            Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", $"开始调用 OCR 服务处理文件: {pdfFilePath}")

            Dim client As New HttpClient()
            Dim pdfBytes As Byte() = File.ReadAllBytes(pdfFilePath)
            Dim pdfBase64 As String = Convert.ToBase64String(pdfBytes)
            Dim encodedBase64 As String = HttpUtility.UrlEncode(pdfBase64)

            Dim content As New StringContent($"pdf_file={encodedBase64}", Encoding.UTF8, "application/x-www-form-urlencoded")
            Dim response As HttpResponseMessage = client.PostAsync($"{OCRUrl}?access_token={accessToken}", content).Result
            response.EnsureSuccessStatusCode()

            Dim responseContent As String = response.Content.ReadAsStringAsync().Result
            Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", $"OCR 服务成功返回结果，文件: {pdfFilePath}")
            Return responseContent
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "OCRProcessor", $"OCR 请求失败，文件: {pdfFilePath}, 错误: {ex.Message}")
            Return $"OCR 请求失败: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' 获取 OCR 服务的 Access Token。
    ''' </summary>
    ''' <returns>返回 Access Token 字符串。</returns>
    Private Function GetAccessToken() As String
        Try
            ' 检查 Settings 中的 Access Token 是否有效
            Dim cachedAccessToken As String = My.Settings.AccessToken
            Dim expiryTime As DateTime

            If Not String.IsNullOrEmpty(cachedAccessToken) AndAlso
           DateTime.TryParse(My.Settings.ExpiryTime, expiryTime) AndAlso
           DateTime.Now < expiryTime Then
                Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", "使用 Settings 中缓存的 Access Token。")
                Return cachedAccessToken
            End If

            Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", "开始获取新的 Access Token。")

            ' 获取新的 Access Token
            Dim client As New HttpClient()
            Dim requestUrl As String = $"{TokenUrl}?grant_type=client_credentials&client_id={ClientId}&client_secret={ClientSecret}"
            Dim response As HttpResponseMessage = client.PostAsync(requestUrl, Nothing).Result
            response.EnsureSuccessStatusCode()

            ' 解析响应
            Dim responseContent As String = response.Content.ReadAsStringAsync().Result
            Dim tokenResponse As Dictionary(Of String, String) = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(responseContent)

            If tokenResponse.ContainsKey("access_token") AndAlso tokenResponse.ContainsKey("expires_in") Then
                ' 更新 Access Token 和过期时间
                cachedAccessToken = tokenResponse("access_token")
                Dim expiresInSeconds As Integer = Integer.Parse(tokenResponse("expires_in"))
                expiryTime = DateTime.Now.AddSeconds(expiresInSeconds - 60) ' 提前 1 分钟刷新

                ' 保存到 Settings
                My.Settings.AccessToken = cachedAccessToken
                My.Settings.ExpiryTime = expiryTime.ToString("o") ' ISO 8601 格式
                My.Settings.Save()

                Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", "新的 Access Token 已保存到 Settings。")
                Return cachedAccessToken
            End If

        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "OCRProcessor", $"获取 Access Token 时发生错误: {ex.Message}")
        End Try

        Return String.Empty
    End Function


    ''' <summary>
    ''' 将 OCR 结果保存到文件中。
    ''' </summary>
    ''' <param name="folderPath">文件夹路径。</param>
    ''' <param name="ocrResult">OCR 结果内容。</param>
    Private Sub SaveOCRResult(folderPath As String, ocrResult As String)
        Try
            Dim resultFilePath As String = Path.Combine(folderPath, "OCR_Result.txt")
            File.WriteAllText(resultFilePath, ocrResult)
            Logger.WriteLog(Logger.LogLevel.Info, "OCRProcessor", $"OCR 结果已保存: {resultFilePath}")
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "OCRProcessor", $"保存 OCR 结果时发生错误: {ex.Message}")
        End Try
    End Sub
End Module
