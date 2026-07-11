''' <summary>
''' 识别调度器：实现"本地优先 / 在线兜底"的完整策略，并标记识别来源。
''' 策略（对应配置 PreferLocalParse / AutoFallbackToOcr / OcrEnabled）：
'''  1. 始终先做本地文本抽取 + 本地解析（成本低，且作为 OCR 失败时的兜底）；
'''  2. 若优先本地且本地关键字段充分 → 直接采用本地，不调用 OCR；
'''  3. 若本地文本为空/疑似图片型/字段不足，或用户设置以在线为主 → 调用百度 OCR；
'''  4. OCR 成功 → 优先采用 OCR 结构化字段（本地也有字段则标记 Mixed）；
'''  5. OCR 失败但本地有部分字段 → 返回本地部分成功，并在日志说明原因；
'''  6. 本地与 OCR 都失败 → 返回失败；
'''  7. 结果标记来源：LocalText / BaiduOcr / Mixed / Failed。
''' </summary>
Public Class RecognitionPipeline

    Private ReadOnly _options As BaiduOcrOptions
    Private ReadOnly _extractor As PdfTextExtractor
    Private ReadOnly _local As IInvoiceRecognizer
    Private ReadOnly _online As IInvoiceRecognizer

    Public Sub New(options As BaiduOcrOptions)
        _options = If(options, New BaiduOcrOptions())
        _extractor = New PdfTextExtractor()
        _local = New LocalTextInvoiceRecognizer()
        _online = New BaiduOcrInvoiceRecognizer(_options)
    End Sub

    Public Function Recognize(pdfPath As String, originalFileName As String) As InvoiceRecognitionResult
        ' 1) 本地抽取 + 本地解析
        Dim extract As PdfTextExtractResult = _extractor.Extract(pdfPath)
        Dim context As New RecognitionContext With {
            .PdfPath = pdfPath,
            .OriginalFileName = originalFileName,
            .ExtractedText = extract.Text,
            .TextExtractResult = extract
        }

        Dim localResult As InvoiceRecognitionResult = _local.Recognize(context)
        Dim localHasFields As Boolean = localResult IsNot Nothing AndAlso localResult.Invoice.HasAnyCoreField()
        ' 充分度以“统一命名核心字段”为准（乘车日期/金额/出发地点/到达地点）：
        ' 本地拿不到起点/终点时不视为充分，从而回退在线 OCR。
        Dim localSufficient As Boolean = localResult IsNot Nothing AndAlso localResult.IsUsable _
            AndAlso KeyFieldEvaluator.IsNamingSufficient(localResult.Invoice, localResult.DocumentType)

        ' 2) 优先本地且充分 → 采用本地
        If _options.PreferLocalParse AndAlso localSufficient Then
            Return FinalizeLocal(localResult, extract)
        End If

        ' 3) 是否调用 OCR
        Dim canOnline As Boolean = _options.IsConfigured() AndAlso _online.IsAvailable()
        Dim wantOnline As Boolean = (Not _options.PreferLocalParse) OrElse _options.AutoFallbackToOcr OrElse extract.IsLikelyImageBased

        ' 本地不足时，明确记录：缺失字段 / 是否启用 OCR / 是否会回退（便于诊断常规发票命名字段不足场景）
        If Not localSufficient Then
            Dim missingText As String = If(localResult IsNot Nothing AndAlso localResult.MissingKeyFields IsNot Nothing AndAlso localResult.MissingKeyFields.Count > 0,
                                           String.Join("/", localResult.MissingKeyFields.ToArray()), "(无)")
            AppLogger.Info(String.Format("本地识别字段不足：文件={0}, 类型={1}, 缺失字段={2}, 在线OCR启用={3}, 配置完整={4}, 将回退OCR={5}",
                                         PrivacySafeFormatter.MaskFileName(originalFileName),
                                         If(localResult IsNot Nothing, localResult.DocumentType.ToString(), "?"),
                                         missingText, _options.Enabled, _options.IsConfigured(), (canOnline AndAlso wantOnline)))
        End If

        If canOnline AndAlso wantOnline Then
            AppLogger.Info("本地字段不足/需在线，开始回退在线百度 OCR：" & PrivacySafeFormatter.MaskFileName(originalFileName))
            Dim ocrResult As InvoiceRecognitionResult = _online.Recognize(context)

            If ocrResult.IsUsable Then
                ' 4) OCR 可用 → 优先 OCR 字段；本地也有字段则 Mixed
                ocrResult.Source = If(localHasFields, RecognitionSource.Mixed, RecognitionSource.BaiduOcr)
                ocrResult.MissingKeyFields = KeyFieldEvaluator.GetMissingKeyFields(ocrResult.Invoice, ocrResult.DocumentType)
                If String.IsNullOrEmpty(ocrResult.RawText) Then ocrResult.RawText = extract.Text
                AppLogger.Info("采用 OCR 结果，来源=" & ocrResult.Source.ToString())
                Return ocrResult
            End If

            ' 5) OCR 失败但本地有部分字段 → 本地部分成功
            If localHasFields Then
                localResult.Status = RecognitionStatus.PartialSuccess
                localResult.Source = RecognitionSource.LocalText
                localResult.Message = "在线 OCR 未成功（" & ocrResult.Message & "），采用本地部分字段。"
                localResult.MissingKeyFields = KeyFieldEvaluator.GetMissingKeyFields(localResult.Invoice, localResult.DocumentType)
                AppLogger.Warn("OCR 未成功，回退本地部分字段：" & ocrResult.Message)
                Return localResult
            End If

            ' 6) 都失败
            If ocrResult.Status = RecognitionStatus.ConfigurationMissing Then
                Return ocrResult ' 配置缺失原样返回
            End If
            ocrResult.Source = RecognitionSource.Failed
            Return ocrResult
        End If

        ' 不调用 OCR 的情形
        If localResult IsNot Nothing AndAlso localResult.IsUsable Then
            Return FinalizeLocal(localResult, extract)
        End If

        ' 本地需要 OCR 但在线不可用/未开启
        Dim reason As String
        If Not _options.Enabled Then
            reason = "疑似需要 OCR，但在线 OCR 未启用（可在设置中开启）。"
        ElseIf Not _options.IsConfigured() Then
            reason = "疑似需要 OCR，但在线 OCR 配置不完整。"
        Else
            reason = "疑似需要 OCR，但当前策略未启用在线识别。"
        End If
        Dim needs As InvoiceRecognitionResult = InvoiceRecognitionResult.NeedsOcr(reason, extract.Text)
        needs.Source = RecognitionSource.LocalText
        Return needs
    End Function

    Private Function FinalizeLocal(localResult As InvoiceRecognitionResult, extract As PdfTextExtractResult) As InvoiceRecognitionResult
        localResult.Source = RecognitionSource.LocalText
        localResult.MissingKeyFields = KeyFieldEvaluator.GetMissingKeyFields(localResult.Invoice, localResult.DocumentType)
        If String.IsNullOrEmpty(localResult.RawText) Then localResult.RawText = extract.Text
        Return localResult
    End Function

End Class
