Imports System.Collections.Generic

''' <summary>
''' 坐标行重建：把带坐标的词按 Y 聚类为行、行内按 X 排序、行间按 Y 从上到下排序。
''' 用于表格版式的稳定解析（不依赖 page.Text 拼接顺序）。
''' </summary>
Public Module PdfTextLayoutExtractor

    ''' <summary>同一行的 Y 容差（PDF 单位）。行程单每行行高约 14，单元格换行 Y 差约 7，取 4 以内为同一主行。</summary>
    Public Const LineYTolerance As Double = 4.0

    ''' <summary>
    ''' 把一页的词聚类为行。words 需含坐标；返回按 Y 从上到下（Y 降序）的行列表，行内按 X 升序。
    ''' </summary>
    Public Function ClusterIntoLines(words As List(Of PdfTextWord), pageIndex As Integer) As List(Of PdfTextLine)
        Dim lines As New List(Of PdfTextLine)()
        If words Is Nothing OrElse words.Count = 0 Then
            Return lines
        End If

        ' 按 Y 降序（PDF 中 Y 越大越靠上），再按 X 升序，稳定分组。
        Dim sorted As New List(Of PdfTextWord)(words)
        sorted.Sort(Function(a, b)
                        Dim c As Integer = b.Y.CompareTo(a.Y)
                        If c <> 0 Then Return c
                        Return a.X.CompareTo(b.X)
                    End Function)

        For Each w As PdfTextWord In sorted
            Dim target As PdfTextLine = Nothing
            For Each ln As PdfTextLine In lines
                If Math.Abs(ln.Y - w.Y) <= LineYTolerance Then
                    target = ln
                    Exit For
                End If
            Next
            If target Is Nothing Then
                target = New PdfTextLine With {.Y = w.Y, .PageIndex = pageIndex}
                lines.Add(target)
            End If
            target.Words.Add(w)
        Next

        ' 行内按 X 升序
        For Each ln As PdfTextLine In lines
            ln.Words.Sort(Function(a, b) a.X.CompareTo(b.X))
        Next
        ' 行间按 Y 降序（从上到下）
        lines.Sort(Function(a, b) b.Y.CompareTo(a.Y))
        Return lines
    End Function

End Module
