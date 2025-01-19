Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

Module MainInvoicearchive
    Public Sub CheckLicense()
        ' 检查 ApiKey、SecretKey 和 ArchivePath 是否为空
        If String.IsNullOrEmpty(My.Settings.ApiKey) OrElse
       String.IsNullOrEmpty(My.Settings.SecretKey) OrElse
       String.IsNullOrEmpty(My.Settings.archivePath) Then
            ' 如果有空值，则跳转到设置窗体
            MessageBox.Show("请检查设置，确保 ApiKey、SecretKey 和 ArchivePath 已正确配置。", "设置检查", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Dim settingsForm As New Form_Settings()
            settingsForm.ShowDialog()
            Return ' 退出当前操作
        End If


        ' 所有检查通过，执行邮件处理
        StartMailProcessing()
    End Sub

    ' 检查当前邮件的收件人和许可有效性
    Private Function IsValidLicenseUser() As Boolean
        Try
            ' 获取当前选择的邮件
            Dim explorer As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()
            Dim selection As Outlook.Selection = explorer.Selection

            If selection.Count = 0 Then
                MessageBox.Show("请先选择一封邮件进行检查。", "无邮件", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' 如果所有检查通过
            Return True
        Catch ex As system.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "CheckLicense", $"许可验证时发生错误: {ex.Message}")
            'MessageBox.Show("验证许可时发生错误，请查看日志获取更多信息。", "许可验证失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function



    ' 开始邮件处理的功能
    Private Sub StartMailProcessing()
        ' 打开 MailProgressBarForm 窗体
        Dim progressForm As New MailProgressBarForm()

        ' 显示窗体（使用 ShowDialog() 模态显示）
        progressForm.ShowDialog()

        ' 你可以在这里继续其他处理逻辑，确保进度条窗体关闭后继续执行
        ' 如果需要，可以在窗体关闭后执行一些后续处理，如：
        ' MessageBox.Show("邮件处理完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


End Module
