Imports System.Collections.Generic

''' <summary>
''' 一封邮件的 PDF 附件分组。归档以“邮件”为单位：同一封邮件内的多个 PDF（发票/行程单等）
''' 归为一组，后续合并为一个 PDF 再识别、命名、归档。
''' </summary>
Public Class MailPdfGroup

    Public Sub New()
        Pdfs = New List(Of MailAttachmentItem)()
    End Sub

    ''' <summary>邮件在选择集中的序号（从 1 开始）。</summary>
    Public Property Index As Integer

    Public Property MailSubject As String
    Public Property SenderName As String
    Public Property ReceivedTime As String

    ''' <summary>该邮件内的 PDF 附件（已导出到临时目录）。</summary>
    Public Property Pdfs As List(Of MailAttachmentItem)

    Public ReadOnly Property PdfCount As Integer
        Get
            Return If(Pdfs Is Nothing, 0, Pdfs.Count)
        End Get
    End Property

    Public ReadOnly Property HasPdf As Boolean
        Get
            Return PdfCount > 0
        End Get
    End Property

End Class
