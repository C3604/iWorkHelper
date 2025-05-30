Imports UglyToad.PdfPig
Imports UglyToad.PdfPig.Content
Imports System.Text
Imports System.Threading

' 定义OCR类
Public Class PdfpigOcr

    ' 重试配置
    Private Const MaxRetries As Integer = 3
    Private Const RetryDelayMs As Integer = 500

    ' 方法：接受PDF文件路径作为输入，返回OCR结果（长文本类型）
    Public Function ExtractTextFromPdf(pdfFilePath As String) As String
        ' 初始化一个StringBuilder来拼接PDF中的文本内容
        Dim result As New StringBuilder()

        ' 记录方法开始
        LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ExtractTextFromPdf", $"开始处理PDF文件：{pdfFilePath}")

        ' 检查文件是否存在
        If Not System.IO.File.Exists(pdfFilePath) Then
            LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf", $"PDF文件不存在：{pdfFilePath}")
            Return String.Empty
        End If

        ' 重试计数器
        Dim retryCount As Integer = 0
        Dim lastException As Exception = Nothing

        ' 重试循环
        While retryCount < MaxRetries
            Try
                ' 使用PdfDocument打开PDF文件
                Using document As PdfDocument = PdfDocument.Open(pdfFilePath)
                    ' 如果文档正常打开但页数为0，记录并返回空字符串
                    If document.NumberOfPages = 0 Then
                        LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf", $"PDF文件没有页面：{pdfFilePath}")
                        Return String.Empty
                    End If

                    ' 遍历PDF文件中的所有页面
                    For pageIndex As Integer = 0 To document.NumberOfPages - 1
                        Try
                            ' 获取当前页的文本
                            Dim pageText As String = document.GetPage(pageIndex + 1).Text

                            ' 将当前页的文本内容添加到result中
                            result.Append(pageText)

                            ' 记录每一页处理的情况
                            LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ExtractTextFromPdf",
                                             $"成功提取第 {pageIndex + 1} 页的文本内容")
                        Catch pageEx As Exception
                            ' 如果单页处理失败，记录错误但继续处理其他页面
                            LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf",
                                            $"处理第 {pageIndex + 1} 页时出错，跳过此页：{pageEx.Message}")
                        End Try
                    Next
                End Using

                ' 如果提取的文本为空，记录警告
                If result.Length = 0 Then
                    LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf",
                                     $"PDF文件不包含可提取的文本内容：{pdfFilePath}")
                End If

                ' 记录提取结果状态
                LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ExtractTextFromPdf",
                                 $"PDF文件处理完成，成功提取文本。文件路径：{pdfFilePath}")

                ' 成功完成，返回提取的文本内容
                Return result.ToString()
            Catch ex As system.Exception
                ' 捕获异常并记录错误信息
                lastException = ex
                LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf",
                                 $"处理PDF文件时出现错误（尝试 {retryCount + 1}/{MaxRetries}）: {ex.Message}")

                ' 在重试之前等待一段时间
                Thread.Sleep(RetryDelayMs)
                retryCount += 1

                ' 强制垃圾回收，释放资源
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End While

        ' 所有重试都失败
        LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf",
                         $"在 {MaxRetries} 次尝试后仍然无法处理PDF文件: {pdfFilePath}，最后错误: {lastException?.Message}")
        
        ' 返回空字符串表示发生错误
        Return String.Empty 
    End Function

    ' 尝试修复不完整或损坏的PDF文件（简单实现）
    Public Function TryRepairPdf(pdfFilePath As String) As Boolean
        Try
            LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.TryRepairPdf", $"尝试修复PDF文件：{pdfFilePath}")
            
            ' 读取整个PDF文件到内存
            Dim pdfBytes As Byte() = System.IO.File.ReadAllBytes(pdfFilePath)
            
            ' 创建一个临时修复文件
            Dim repairFilePath As String = pdfFilePath & ".repaired"
            
            ' 尝试使用PdfPig打开并另存为修复后的文件
            Using document As PdfDocument = PdfDocument.Open(pdfBytes)
                ' 如果能打开，则使用UglyToad.PdfPig库将文档保存到新文件
                ' 注意：PdfPig当前版本可能不支持直接保存，这里仅作为示意
                LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.TryRepairPdf", $"PDF文件似乎可以打开，但可能需要更专业的工具来修复")
                Return False
            End Using
            
            ' 如果成功创建了修复文件，则替换原文件
            If System.IO.File.Exists(repairFilePath) Then
                System.IO.File.Delete(pdfFilePath)
                System.IO.File.Move(repairFilePath, pdfFilePath)
                LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.TryRepairPdf", $"PDF文件修复成功：{pdfFilePath}")
                Return True
            End If
            
            Return False
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "PdfProcessor.TryRepairPdf", $"尝试修复PDF文件时出错：{ex.Message}")
            Return False
        End Try
    End Function

End Class
