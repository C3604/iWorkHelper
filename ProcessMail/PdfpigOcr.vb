Imports UglyToad.PdfPig
Imports UglyToad.PdfPig.Content
Imports System.Text

' 定义OCR类
Public Class PdfpigOcr

    ' 方法：接受PDF文件路径作为输入，返回OCR结果（长文本类型）
    Public Function ExtractTextFromPdf(pdfFilePath As String) As String
        ' 初始化一个StringBuilder来拼接PDF中的文本内容
        Dim result As New StringBuilder()

        ' 记录方法开始
        LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ExtractTextFromPdf", $"开始处理PDF文件：{pdfFilePath}")

        Try
            ' 使用PdfDocument打开PDF文件
            Using document As PdfDocument = PdfDocument.Open(pdfFilePath)
                ' 遍历PDF文件中的所有页面
                For pageIndex As Integer = 0 To document.NumberOfPages - 1
                    ' 获取当前页的文本
                    Dim pageText As String = document.GetPage(pageIndex + 1).Text

                    ' 将当前页的文本内容添加到result中
                    result.Append(pageText)

                    ' 记录每一页处理的情况
                    LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ExtractTextFromPdf",
                                     $"成功提取第 {pageIndex + 1} 页的文本内容")
                Next
            End Using
        Catch ex As Exception
            ' 捕获异常并输出错误信息
            LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ExtractTextFromPdf",
                             $"处理PDF文件时出现错误: {ex.Message}")
            Return String.Empty ' 返回空字符串表示发生错误
        End Try

        ' 记录提取结果状态
        LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ExtractTextFromPdf",
                         $"PDF文件处理完成，成功提取文本。文件路径：{pdfFilePath}")

        ' 返回提取的文本内容
        Return result.ToString()
    End Function

End Class
