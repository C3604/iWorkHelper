''' <summary>
''' 识别输入上下文。同时携带 PDF 物理路径与本地已抽取文本，
''' 使本地识别器与在线 OCR 识别器可共用同一入参。
''' </summary>
Public Class RecognitionContext

    ''' <summary>待识别 PDF 的物理路径（临时文件）。</summary>
    Public Property PdfPath As String

    ''' <summary>PdfPig 抽取到的文本层内容（可能为空，表示疑似图片型）。</summary>
    Public Property ExtractedText As String

    ''' <summary>本地文本抽取结果（页数、是否疑似图片型等），可为空。</summary>
    Public Property TextExtractResult As PdfTextExtractResult

    ''' <summary>原始附件文件名，用于回退命名与日志。</summary>
    Public Property OriginalFileName As String

End Class
