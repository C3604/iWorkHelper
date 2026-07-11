''' <summary>
''' 票据识别器统一接口。本地文本识别器与百度 OCR 识别器均实现该接口，
''' 使 ParseMode（本地/在线）可在运行时选择实现，并支持“本地优先、必要时回退在线”的策略。
''' </summary>
Public Interface IInvoiceRecognizer

    ''' <summary>识别器名称，用于日志与结果标注。</summary>
    ReadOnly Property Name As String

    ''' <summary>
    ''' 在当前配置/输入下是否可用（如在线 OCR 未配置则返回 False）。
    ''' </summary>
    Function IsAvailable() As Boolean

    ''' <summary>
    ''' 执行识别。实现必须捕获自身异常并以 InvoiceRecognitionResult 反馈状态，
    ''' 不应向调用方抛出未处理异常。
    ''' </summary>
    Function Recognize(context As RecognitionContext) As InvoiceRecognitionResult

End Interface
