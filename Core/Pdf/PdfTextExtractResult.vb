''' <summary>
''' PDF 文本层抽取结果。用于判断是否成功取到有效文本，或疑似图片型 PDF 需走 OCR。
''' </summary>
Public Class PdfTextExtractResult

    ''' <summary>抽取是否成功（未发生异常）。注意：成功不代表一定有文本。</summary>
    Public Property Success As Boolean

    ''' <summary>PDF 页数。</summary>
    Public Property PageCount As Integer

    ''' <summary>抽取到的文本。</summary>
    Public Property Text As String

    ''' <summary>文本长度（去除空白后的有效字符数）。</summary>
    Public Property TextLength As Integer

    ''' <summary>
    ''' 是否疑似图片型 PDF（文本层为空或极少）。
    ''' 为 True 时应交给 OCR，而非依赖本地文本解析。
    ''' </summary>
    Public Property IsLikelyImageBased As Boolean

    ''' <summary>错误信息（抽取失败时）。</summary>
    Public Property ErrorMessage As String

    ''' <summary>
    ''' 坐标行重建结果（跨页扁平；每行含页码/Y/词坐标）。用于表格版式的按列稳定解析。
    ''' 抽取失败或无词时为空列表。
    ''' </summary>
    Public Property Lines As System.Collections.Generic.List(Of PdfTextLine)

    Public Sub New()
        Lines = New System.Collections.Generic.List(Of PdfTextLine)()
    End Sub

End Class
