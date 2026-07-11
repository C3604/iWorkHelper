''' <summary>
''' 从邮件中导出的单个 PDF 附件。记录来源邮件信息与临时保存路径，供后续识别/归档使用。
''' </summary>
Public Class MailAttachmentItem

    ''' <summary>来源邮件主题。</summary>
    Public Property MailSubject As String

    ''' <summary>来源邮件发件人（显示名或地址）。</summary>
    Public Property SenderName As String

    ''' <summary>来源邮件接收时间（原始文本，避免时区/格式歧义）。</summary>
    Public Property ReceivedTime As String

    ''' <summary>附件原始文件名。</summary>
    Public Property OriginalFileName As String

    ''' <summary>导出到临时工作目录后的完整路径。</summary>
    Public Property TempFilePath As String

    ''' <summary>附件字节大小（如可获取）。</summary>
    Public Property SizeBytes As Long

End Class
