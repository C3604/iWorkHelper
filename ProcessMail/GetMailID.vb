Imports Microsoft.Office.Interop.Outlook
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Threading

' 创建一个类来处理选中的邮件
Public Class GetMailID
    ' 声明一个列表来保存包含附件的邮件的EntryID
    Private mailWithAttachments As New List(Of String)()

    ' 构造函数，初始化邮件附件处理类
    Public Sub New()
        ' 可以在这里进行一些初始化操作（如果需要）
        LogManager.WriteLog(LogLevel.INFO, "GetMailID.New", "初始化 GetMailID 实例")
    End Sub

    ' 获取选中的邮件并处理每封邮件
    Public Sub ProcessSelectedEmails()
        ' 获取Outlook应用程序对象
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Nothing
        Dim explorer As Explorer = Nothing
        Dim selection As Selection = Nothing

        Try
            ' 尝试获取当前运行的Outlook实例
            Try
                outlookApp = GetRunningOutlookInstance()
                If outlookApp Is Nothing Then
                    LogManager.WriteLog(LogLevel.Info, "GetMailID.ProcessSelectedEmails", "未找到运行中的Outlook实例，正在创建新实例")
                    outlookApp = New Microsoft.Office.Interop.Outlook.Application()
                End If
            Catch ex As System.Exception
                LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"创建Outlook应用程序实例失败: {ex.Message}")
                MessageBox.Show($"无法访问Outlook应用程序: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' 获取当前活动的资源管理器窗口
            Try
                explorer = outlookApp.ActiveExplorer()
                If explorer Is Nothing Then
                    LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", "无法获取活动的Explorer窗口")
                    MessageBox.Show("请确保Outlook已打开且有邮件窗口处于活动状态", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            Catch ex As System.Exception
                LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"获取ActiveExplorer失败: {ex.Message}")
                MessageBox.Show($"无法访问Outlook窗口: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' 获取当前选中的邮件项集合
            Try
                selection = explorer.Selection
                If selection Is Nothing OrElse selection.Count = 0 Then
                    LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", "未选中任何邮件")
                    MessageBox.Show("请先选择包含PDF附件的邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If
                
                LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", $"成功获取选中的邮件，共 {selection.Count} 封")
            Catch ex As System.Exception
                LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"获取选中邮件失败: {ex.Message}")
                MessageBox.Show($"无法获取选中的邮件: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' 遍历每一封选中的邮件
            For i As Integer = 1 To selection.Count
                Dim item As Object = Nothing
                Dim mailItem As MailItem = Nothing

                Try
                    ' 获取邮件项
                    item = selection.Item(i)
                    
                    ' 确保item是邮件对象
                    If TypeOf item Is MailItem Then
                        mailItem = CType(item, MailItem)
                        
                        ' 获取邮件的EntryID
                        Dim mailEntryID As String = mailItem.EntryID
                        Dim subject As String = mailItem.Subject

                        LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", 
                                       $"处理第 {i} 封邮件，主题: {subject}, EntryID: {mailEntryID}")

                        ' 判断邮件是否包含附件
                        If mailItem.Attachments.Count > 0 Then
                            Dim hasPdfAttachment As Boolean = False
                            
                            ' 只处理PDF附件
                            For Each attachment As Attachment In mailItem.Attachments
                                If attachment.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                                    LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", 
                                                  $"找到PDF附件: {attachment.FileName}")
                                    hasPdfAttachment = True
                                    Exit For ' 只要有一个PDF附件就认为该邮件有附件
                                End If
                            Next

                            If hasPdfAttachment Then
                                mailWithAttachments.Add(mailEntryID)
                                LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", 
                                             $"邮件已添加到处理列表，主题: {subject}, EntryID: {mailEntryID}")
                            Else
                                LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", 
                                             $"邮件不包含PDF附件，已跳过，主题: {subject}")
                            End If
                        Else
                            LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", 
                                         $"邮件没有附件，已跳过，主题: {subject}")
                        End If
                    Else
                        LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", 
                                     $"选中的项目不是邮件，类型: {item.GetType().Name}")
                    End If
                Catch ex As System.Exception
                    ' 处理单个邮件的错误，继续处理其他邮件
                    LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"处理邮件项时出错: {ex.Message}")
                Finally
                    ' 释放COM对象
                    If mailItem IsNot Nothing Then
                        Try
                            Marshal.ReleaseComObject(mailItem)
                        Catch ex As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"释放邮件COM对象时出错: {ex.Message}")
                        End Try
                        mailItem = Nothing
                    End If
                    
                    If item IsNot Nothing Then
                        Try
                            Marshal.ReleaseComObject(item)
                        Catch ex As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"释放项目COM对象时出错: {ex.Message}")
                        End Try
                        item = Nothing
                    End If
                End Try
            Next

            LogManager.WriteLog(LogLevel.INFO, "GetMailID.ProcessSelectedEmails", 
                         $"成功处理选中的邮件，共找到 {mailWithAttachments.Count} 封包含PDF附件的邮件")
        Catch ex As System.Exception
            ' 处理整体处理过程中的未捕获异常
            LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"处理邮件过程中发生未捕获的异常: {ex.Message}")
            MessageBox.Show($"处理邮件时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 确保所有COM对象都被释放
            If selection IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(selection)
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"释放Selection COM对象时出错: {ex.Message}")
                End Try
                selection = Nothing
            End If
            
            If explorer IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(explorer)
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"释放Explorer COM对象时出错: {ex.Message}")
                End Try
                explorer = Nothing
            End If
            
            If outlookApp IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(outlookApp)
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "GetMailID.ProcessSelectedEmails", $"释放Outlook应用程序COM对象时出错: {ex.Message}")
                End Try
                outlookApp = Nothing
            End If
            
            ' 强制垃圾回收
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    ' 获取包含附件的邮件EntryID列表
    Public Function GetMailWithAttachments() As List(Of String)
        Return mailWithAttachments
    End Function

    ' 错误日志记录
    Private Sub LogError(message As String)
        Try
            If Not EventLog.SourceExists("GetMailID") Then
                EventLog.CreateEventSource("GetMailID", "Application")
            End If
            EventLog.WriteEntry("GetMailID", message, EventLogEntryType.Error)
        Catch ex As system.Exception
            ' 如果无法写入事件日志，至少尝试通过LogManager记录
            LogManager.WriteLog(LogLevel.Error, "GetMailID.LogError", $"写入事件日志失败: {ex.Message}，原始错误: {message}")
        End Try
    End Sub

    ' 获取当前运行的Outlook实例
    Private Function GetRunningOutlookInstance() As Microsoft.Office.Interop.Outlook.Application
        Try
            Return DirectCast(Marshal.GetActiveObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Info, "GetMailID.GetRunningOutlookInstance", $"未找到运行中的Outlook实例: {ex.Message}")
            Return Nothing
        End Try
    End Function
End Class
