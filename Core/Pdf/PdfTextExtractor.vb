Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports UglyToad.PdfPig

''' <summary>
''' 基于 PdfPig 的 PDF 文本层抽取器。
''' 重要：PdfPig 只能读取 PDF 中"已存在的文本层"，并非 OCR。
''' 对扫描件/图片型 PDF 会抽不到文本，此时应标记为疑似图片型并交给 OCR。
''' </summary>
Public Class PdfTextExtractor

    ''' <summary>
    ''' 有效文本的最小阈值。低于该阈值视为疑似图片型 PDF（按每页粗略估计）。
    ''' </summary>
    Private Const MinValidTextLength As Integer = 20

    ''' <summary>
    ''' 抽取指定 PDF 的文本层。任何异常都转换为结果对象，不向上抛出。
    ''' </summary>
    Public Function Extract(pdfPath As String) As PdfTextExtractResult
        Dim result As New PdfTextExtractResult()

        Try
            If String.IsNullOrWhiteSpace(pdfPath) OrElse Not File.Exists(pdfPath) Then
                result.Success = False
                result.ErrorMessage = "PDF 文件不存在: " & PrivacySafeFormatter.MaskPath(pdfPath)
                Return result
            End If

            Dim sb As New StringBuilder()
            Dim pageCount As Integer = 0

            Using document As PdfDocument = PdfDocument.Open(pdfPath)
                pageCount = document.NumberOfPages
                Dim pageIdx As Integer = 0
                For Each page In document.GetPages()
                    pageIdx += 1
                    Dim pageText As String = page.Text
                    If Not String.IsNullOrEmpty(pageText) Then
                        sb.AppendLine(pageText)
                    End If
                    ' 坐标行重建（用于表格版式稳定解析）
                    Try
                        Dim words As New List(Of PdfTextWord)()
                        For Each w In page.GetWords()
                            Dim bb = w.BoundingBox
                            words.Add(New PdfTextWord With {.Text = w.Text, .X = bb.Left, .Right = bb.Right, .Y = bb.Bottom})
                        Next
                        result.Lines.AddRange(PdfTextLayoutExtractor.ClusterIntoLines(words, pageIdx))
                    Catch exLayout As Exception
                        AppLogger.Warn("坐标行重建失败（忽略，回退全文正则）：" & exLayout.Message)
                    End Try
                Next
            End Using

            Dim text As String = sb.ToString()
            result.Success = True
            result.PageCount = pageCount
            result.Text = text
            result.TextLength = CountNonWhitespace(text)

            ' 判断是否疑似图片型：有效字符过少（考虑页数）。
            Dim threshold As Integer = MinValidTextLength * Math.Max(1, pageCount)
            result.IsLikelyImageBased = (result.TextLength < threshold)

            AppLogger.Info(String.Format("PDF 文本抽取完成: 页数={0}, 有效字符={1}, 疑似图片型={2}, 文件={3}",
                                         result.PageCount, result.TextLength, result.IsLikelyImageBased,
                                         PrivacySafeFormatter.MaskFileName(Path.GetFileName(pdfPath))))
            Return result

        Catch ex As Exception
            result.Success = False
            result.ErrorMessage = ExceptionFormatter.ToUserMessage(ex)
            result.IsLikelyImageBased = True ' 抽取失败时，保守地交给 OCR 决策
            AppLogger.Error("PDF 文本抽取失败: " & PrivacySafeFormatter.MaskPath(pdfPath), ex)
            Return result
        End Try
    End Function

    Private Function CountNonWhitespace(text As String) As Integer
        If String.IsNullOrEmpty(text) Then
            Return 0
        End If
        Dim count As Integer = 0
        For Each c As Char In text
            If Not Char.IsWhiteSpace(c) Then
                count += 1
            End If
        Next
        Return count
    End Function

End Class
