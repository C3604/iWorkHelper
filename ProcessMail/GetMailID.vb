Imports Microsoft.Office.Interop.Outlook
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Windows.Forms

' 创建一个类来处理选中的邮件
Public Class GetMailID
    ' 声明一个列表来保存包含附件的邮件的EntryID
    Private mailWithAttachments As New List(Of String)()

    ' 构造函数，初始化邮件附件处理类
    Public Sub New()
        ' 可以在这里进行一些初始化操作（如果需要）
    End Sub

    ' 获取选中的邮件并处理每封邮件
    Public Sub ProcessSelectedEmails()
        ' 获取Outlook应用程序对象
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Nothing
        Try
            outlookApp = Globals.ThisAddIn.Application
            If outlookApp Is Nothing Then
                Throw New InvalidOperationException("无法访问Outlook应用程序")
            End If
        Catch ex As System.Exception
            ' 处理初始化失败的情况
            LogError("Outlook应用程序访问失败: " & ex.Message)
            MessageBox.Show("无法访问Outlook应用程序: " & ex.Message)
            Return
        End Try

        ' 获取当前选中的邮件项集合
        Dim selectedItems As Selection = outlookApp.ActiveExplorer().Selection

        ' 遍历每一封选中的邮件
        For Each item As Object In selectedItems
            Try
                ' 确保item是邮件对象
                If TypeOf item Is MailItem Then
                    Dim mailItem As MailItem = CType(item, MailItem)
                    ' 获取邮件的EntryID
                    Dim mailEntryID As String = mailItem.EntryID

                    ' 判断邮件是否包含附件
                    If mailItem.Attachments.Count > 0 Then
                        ' 只处理PDF附件
                        For Each attachment As Attachment In mailItem.Attachments
                            If attachment.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                                mailWithAttachments.Add(mailEntryID)
                                Exit For ' 只要有一个PDF附件就认为该邮件有附件
                            End If
                        Next
                    End If
                    ' 释放COM对象
                    Marshal.ReleaseComObject(mailItem)
                End If
            Catch ex As System.Exception
                ' 其他未处理的异常
                LogError("处理邮件时出错: " & ex.Message)
            End Try
        Next

        ' 输出包含附件的邮件ID
        For Each entryID As String In mailWithAttachments
            Console.WriteLine("邮件包含附件: " & entryID)
        Next
    End Sub

    ' 获取包含附件的邮件EntryID列表
    Public Function GetMailWithAttachments() As List(Of String)
        Return mailWithAttachments
    End Function

    ' 错误日志记录
    Private Sub LogError(message As String)
        If Not EventLog.SourceExists("GetMailID") Then
            EventLog.CreateEventSource("GetMailID", "Application")
        End If
        EventLog.WriteEntry("GetMailID", message, EventLogEntryType.Error)
    End Sub
End Class
