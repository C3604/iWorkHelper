Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

Module MainInvoicearchive
    Public Sub CheckLicense()
        ' 检查 ClientId、ClientSecret 和 ArchivePath 是否为空
        If String.IsNullOrEmpty(My.Settings.ClientId) OrElse
       String.IsNullOrEmpty(My.Settings.ClientSecret) OrElse
       String.IsNullOrEmpty(My.Settings.ArchivePath) Then
            ' 如果有空值，则跳转到设置窗体
            MessageBox.Show("请检查设置，确保 ClientId、ClientSecret 和 ArchivePath 已正确配置。", "设置检查", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Dim settingsForm As New FormSetting()
            settingsForm.ShowDialog()
            Return ' 退出当前操作
        End If

        ' 检查 TempMode 是否启用
        If My.Settings.TempMode Then
            ' 检查是否是许可用户
            If Not IsValidLicenseUser() Then
                MessageBox.Show("当前用户的许可无效或过期。", "许可无效", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return ' 退出当前操作
            End If
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

            ' 获取邮件的收件人
            Dim mailItem As Outlook.MailItem = CType(selection.Item(1), Outlook.MailItem)
            Dim recipients As Outlook.Recipients = mailItem.Recipients

            Dim licenseMail As String = My.Settings.LicenseMail
            Dim licenseData As DateTime = DateTime.Parse(My.Settings.LicenseData)

            ' 1. 检查收件人是否包含 LicenseMail
            Dim validRecipient As Boolean = False
            For Each recipient As Outlook.Recipient In recipients
                ' 获取邮箱地址，确保是正确的邮件格式
                Dim recipientAddress As String = String.Empty

                ' 如果是 Exchange 收件人
                If recipient.AddressEntry.Type = "EX" Then
                    ' 获取 Exchange 用户并获取其邮件地址
                    Dim exchangeUser As Outlook.ExchangeUser = recipient.AddressEntry.GetExchangeUser()
                    If exchangeUser IsNot Nothing Then
                        recipientAddress = exchangeUser.PrimarySmtpAddress.ToLower()
                    End If
                Else
                    ' 如果是常规邮件地址
                    recipientAddress = recipient.AddressEntry.Address.ToLower()
                End If

                ' 检查地址并与 LicenseMail 比较
                If recipientAddress.Equals(licenseMail.ToLower()) Then
                    validRecipient = True
                    Logger.WriteLog(Logger.LogLevel.Info, "IsValidLicenseUser", $"匹配到许可邮箱: {licenseMail}")
                    Exit For
                End If
            Next

            ' 如果没有匹配的收件人
            If Not validRecipient Then
                Logger.WriteLog(Logger.LogLevel.Warning, "IsValidLicenseUser", $"未找到匹配的许可邮箱: {licenseMail}，邮件收件人: {String.Join(", ", recipients.Cast(Of Outlook.Recipient)().Select(Function(r) r.AddressEntry.Address))}")
                MessageBox.Show($"邮件的收件人中未找到指定的许可邮箱 {licenseMail}。", "许可验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' 2. 检查当前日期是否小于或等于 LicenseData
            If DateTime.Now > licenseData Then
                Logger.WriteLog(Logger.LogLevel.Warning, "IsValidLicenseUser", $"当前日期 {DateTime.Now.ToString("yyyy-MM-dd")} 已超过许可有效期 ({licenseData.ToString("yyyy-MM-dd")})。")
                MessageBox.Show($"当前日期 {DateTime.Now.ToString("yyyy-MM-dd")} 已超过许可有效期 ({licenseData.ToString("yyyy-MM-dd")})。", "许可过期", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
