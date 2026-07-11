Imports System.Collections.Generic

''' <summary>
''' 批量读取邮件附件的汇总结果。
''' </summary>
Public Class MailReadResult

    Public Sub New()
        Attachments = New List(Of MailAttachmentItem)()
        Messages = New List(Of String)()
    End Sub

    ''' <summary>成功导出的 PDF 附件集合。</summary>
    Public Property Attachments As List(Of MailAttachmentItem)

    ''' <summary>选中项总数（含非邮件项）。</summary>
    Public Property SelectedItemCount As Integer

    ''' <summary>其中被识别为邮件（MailItem）的数量。</summary>
    Public Property MailItemCount As Integer

    ''' <summary>被跳过的非邮件项数量。</summary>
    Public Property SkippedNonMailCount As Integer

    ''' <summary>无 PDF 附件的邮件数量。</summary>
    Public Property MailsWithoutPdfCount As Integer

    ''' <summary>过程消息（跳过原因、异常等）。</summary>
    Public Property Messages As List(Of String)

    Public Sub AddMessage(text As String)
        If Not String.IsNullOrEmpty(text) Then
            Messages.Add(text)
        End If
    End Sub

End Class
