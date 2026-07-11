Imports System.Collections.Generic

''' <summary>
''' 批量读取并按邮件分组的结果。每个 MailPdfGroup 对应一封邮件。
''' </summary>
Public Class MailPdfGroupingResult

    Public Sub New()
        Groups = New List(Of MailPdfGroup)()
        Messages = New List(Of String)()
    End Sub

    ''' <summary>按邮件分组（仅含至少有一个 PDF 的邮件）。</summary>
    Public Property Groups As List(Of MailPdfGroup)

    Public Property SelectedItemCount As Integer
    Public Property MailItemCount As Integer
    Public Property SkippedNonMailCount As Integer
    Public Property MailsWithoutPdfCount As Integer

    Public Property Messages As List(Of String)

    Public Sub AddMessage(text As String)
        If Not String.IsNullOrEmpty(text) Then
            Messages.Add(text)
        End If
    End Sub

    Public ReadOnly Property TotalPdfCount As Integer
        Get
            Dim n As Integer = 0
            For Each g As MailPdfGroup In Groups
                n += g.PdfCount
            Next
            Return n
        End Get
    End Property

End Class
