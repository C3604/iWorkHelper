Imports System.Diagnostics
Imports System.IO
Imports System.Reflection

Module ShowHelpDocument
    Public Sub ShowHelpDocument()
        Try
            ' 获取嵌入资源的字节数组
            Dim pdfBytes As Byte() = My.Resources.OutlookPluginHelpDocument_V0_1

            ' 定义临时文件路径
            Dim tempFilePath As String = Path.Combine(Path.GetTempPath(), "HelpDocument.pdf")

            ' 将字节数组写入临时文件
            File.WriteAllBytes(tempFilePath, pdfBytes)

            ' 打开临时文件
            Process.Start(New ProcessStartInfo(tempFilePath) With {.UseShellExecute = True})
        Catch ex As Exception
            ' 错误处理
            MsgBox($"无法打开帮助文档: {ex.Message}", MsgBoxStyle.Critical, "错误")
        End Try
    End Sub





End Module
