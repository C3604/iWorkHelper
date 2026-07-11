''' <summary>
''' 未识别 PDF 命名规则：未识别_{原始文件名}.pdf。清理 Windows 非法字符；
''' 同名冲突由 ArchivePlanner 追加序号处理，不覆盖已有文件。内部规则，不作为设置项暴露。
''' </summary>
Public Module UnknownPdfNamingRule

    ''' <summary>生成未识别 PDF 文件名（含 .pdf）。</summary>
    Public Function BuildFileName(originalFileName As String) As String
        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(If(originalFileName, "未识别"))
        Return FileNameSanitizer.BuildFileName(NamingTemplates.UnknownPdfPrefix & baseName, ".pdf", "未识别")
    End Function

End Module
