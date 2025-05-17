Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class iWorkHelperRibbon

    Private Sub Btn_Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_Settings.Click
        ' 确保插件已初始化
        ThisAddIn.GetInstance().EnsureInitialized()
        
        ' 创建SettingsForm的实例
        Dim settingsForm As New SettingsForm()

        ' 显示SettingsForm窗体
        settingsForm.ShowDialog() ' 使用ShowDialog()可以使SettingsForm作为模态窗体显示
    End Sub

    Private Sub btn_Archive_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_Archive.Click
        ' 确保插件已初始化
        ThisAddIn.GetInstance().EnsureInitialized()
        
        ' 检查是否设置了归档路径
        If String.IsNullOrEmpty(My.Settings.ArchivePath) Then
            ' 如果归档路径为空，提示用户并跳转到设置窗体
            MessageBox.Show("请先设置归档位置", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' 打开设置窗体
            Dim settingsForm As New SettingsForm()
            settingsForm.ShowDialog()
            Return ' 退出当前操作，等待用户设置归档路径
        End If

        ' 记录按钮点击事件
        If My.Settings.DebugStatus Then
            ' 使用异步方式记录日志
            System.Threading.Tasks.Task.Run(Sub()
                LogManager.WriteLog(LogLevel.INFO, "MainRibbon.btn_Archive_Click", "开始处理邮件归档操作")
            End Sub)
        End If

        Try
            ' 只有在需要时才创建MailProcessor实例
            Dim mailProcessor As New MailProcessor()
            
            ' 调用ProcessMails方法，开始处理邮件
            If My.Settings.DebugStatus Then
                System.Threading.Tasks.Task.Run(Sub()
                    LogManager.WriteLog(LogLevel.INFO, "MainRibbon.btn_Archive_Click", "调用ProcessMails方法开始处理邮件")
                End Sub)
            End If
            
            mailProcessor.ProcessMails()
            
            If My.Settings.DebugStatus Then
                System.Threading.Tasks.Task.Run(Sub()
                    LogManager.WriteLog(LogLevel.INFO, "MainRibbon.btn_Archive_Click", "邮件处理完成")
                End Sub)
            End If

            ' 处理完成后打开归档文件夹
            Dim folderOpener As New FolderOpener(My.Settings.ArchivePath)
        Catch ex As Exception
            ' 捕获异常并记录错误信息
            If My.Settings.DebugStatus Then
                System.Threading.Tasks.Task.Run(Sub()
                    LogManager.WriteLog(LogLevel.Error, "MainRibbon.btn_Archive_Click", $"处理邮件时发生错误：{ex.Message}")
                End Sub)
            End If
            MessageBox.Show($"处理邮件时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
