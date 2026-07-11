Imports System.IO
Imports System.Collections.Generic
Imports UglyToad.PdfPig.Writer

''' <summary>
''' PDF 合并服务。使用 **PdfPig 自带的 `UglyToad.PdfPig.Writer.PdfMerger`**（无需引入新库）。
''' PdfMerger.Merge(String()) 返回合并后的 PDF 字节，可被常见阅读器打开。
''' 约定：调用方按“发票在前、行程单在后”传入顺序；无法分类时按原始附件顺序。
''' 不覆盖原始附件；输出到调用方指定的临时/目标路径。
''' </summary>
Public Class PdfMergeService

    ''' <summary>
    ''' 合并多个 PDF 到 outputPath。
    '''  - 传入 0 个：失败；
    '''  - 传入 1 个：直接复制该 PDF 到 outputPath（不调用合并，保证单 PDF 仍可用）；
    '''  - 传入 ≥2 个：调用 PdfMerger 合并。
    ''' 任何异常都转为 Result，不抛出。
    ''' </summary>
    Public Function Merge(pdfPaths As List(Of String), outputPath As String) As Result
        Try
            Dim valid As New List(Of String)()
            If pdfPaths IsNot Nothing Then
                For Each p As String In pdfPaths
                    If Not String.IsNullOrWhiteSpace(p) AndAlso File.Exists(p) Then
                        valid.Add(p)
                    End If
                Next
            End If

            If valid.Count = 0 Then
                Return Result.Fail("没有可合并的 PDF 文件。")
            End If

            Dim targetDir As String = Path.GetDirectoryName(outputPath)
            If Not PathHelper.EnsureDirectory(targetDir) Then
                Return Result.Fail("无法创建合并输出目录：" & targetDir)
            End If

            If valid.Count = 1 Then
                ' 单 PDF：复制到输出（不覆盖原件）。
                File.Copy(valid(0), outputPath, overwrite:=True)
                AppLogger.Info("单 PDF 无需合并，已复制到临时文件：" & Path.GetFileName(outputPath))
                Return Result.Ok(outputPath)
            End If

            Dim merged As Byte() = PdfMerger.Merge(valid.ToArray())
            If merged Is Nothing OrElse merged.Length = 0 Then
                Return Result.Fail("PDF 合并返回空结果。")
            End If

            File.WriteAllBytes(outputPath, merged)
            AppLogger.Info(String.Format("已合并 {0} 个 PDF -> {1}（{2} 字节）", valid.Count, Path.GetFileName(outputPath), merged.Length))
            Return Result.Ok(outputPath)

        Catch ex As Exception
            AppLogger.Error("PDF 合并失败。", ex)
            Return Result.Fail("PDF 合并失败：" & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Function

End Class
