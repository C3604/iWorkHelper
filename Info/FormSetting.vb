Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class FormSetting
    ''' <summary>
    ''' 窗体加载时的初始化操作。
    ''' </summary>
    Private Sub FormSetting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' 加载设置
            LoadSettings()

            ' 更新 CheckBoxTempMode 的状态
            CheckBoxTempMode.Checked = My.Settings.TempMode

            ' 根据模式状态更新 UI
            If CheckBoxTempMode.Checked Then
                EnableTempMode()
            Else
                DisableTempMode()
            End If

            ' 日志记录
            Logger.WriteLog(Logger.LogLevel.Info, "FormSetting", "设置已加载成功。")
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"窗体加载时发生错误: {ex.Message}")
        End Try
    End Sub





    ''' <summary>
    ''' 加载用户设置并更新到界面控件。
    ''' </summary>
    Private Sub LoadSettings()
        Try
            CheckBoxdebug.Checked = My.Settings.DebugMode
            CheckBoxlog.Checked = My.Settings.LogModel
            TextBoxArchivePath.Text = My.Settings.ArchivePath
            TextBoxOCRRequestAddress.Text = My.Settings.OCRRequestAddress

            ' 无论是否是临时模式，都加载 ClientId 和 ClientSecret
            TextBoxClientId.Text = My.Settings.ClientId
            TextBoxClientSecret.Text = My.Settings.ClientSecret

            LabelLicenseData.Text = If(String.IsNullOrEmpty(My.Settings.LicenseData), "未设置", My.Settings.LicenseData)
            LabelLicenseMail.Text = If(String.IsNullOrEmpty(My.Settings.LicenseMail), "未授权", My.Settings.LicenseMail)
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"加载设置时发生错误: {ex.Message}")
        End Try
    End Sub



    ''' <summary>
    ''' 根据设置状态动态更新界面 UI。
    ''' </summary>
    Private Sub UpdateUI()
        Try
            Buttonlog.Enabled = CheckBoxlog.Checked
            Buttonlog.Visible = CheckBoxlog.Checked
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"更新 UI 时发生错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 保存当前界面控件中的设置到用户配置。
    ''' </summary>
    Private Sub SaveSettings()
        Try
            ' 保存所有设置
            My.Settings.ClientId = TextBoxClientId.Text
            My.Settings.ClientSecret = TextBoxClientSecret.Text
            My.Settings.DebugMode = CheckBoxdebug.Checked
            My.Settings.LogModel = CheckBoxlog.Checked
            My.Settings.ArchivePath = TextBoxArchivePath.Text
            My.Settings.OCRRequestAddress = TextBoxOCRRequestAddress.Text

            ' 保存到配置文件
            My.Settings.Save()
            Logger.WriteLog(Logger.LogLevel.Info, "FormSetting", "设置已保存成功。")
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"保存设置时发生错误: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' 日志模式状态切换时的操作。
    ''' </summary>
    Private Sub CheckBoxlog_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxlog.CheckedChanged
        Try
            My.Settings.LogModel = CheckBoxlog.Checked
            UpdateUI()

            ' 如果日志模式被禁用，删除日志文件
            If Not CheckBoxlog.Checked Then
                Dim logFilePath As String = GetLogFilePath()
                If File.Exists(logFilePath) Then
                    File.Delete(logFilePath)
                End If
            End If
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"切换日志模式时发生错误: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' 获取日志文件路径。
    ''' </summary>
    ''' <returns>日志文件路径。</returns>
    Private Function GetLogFilePath() As String
        Return Path.Combine(Path.GetTempPath(), "OutlookPlugin.log")
    End Function



    Private Sub CheckBoxshowtoken_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxShowToken.CheckedChanged
        SetTextBoxPasswordMode(CheckBoxShowToken.Checked)
    End Sub

    ''' <summary>
    ''' 设置密码框显示模式
    ''' </summary>
    Private Sub SetTextBoxPasswordMode(showToken As Boolean)
        TextBoxClientId.UseSystemPasswordChar = Not showToken
        TextBoxClientSecret.UseSystemPasswordChar = Not showToken
    End Sub

    Private Sub ButtonSubmit_Click(sender As Object, e As EventArgs) Handles ButtonSubmit.Click

        SaveSettings()
        Me.Close()
    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        Try
            SaveSettings()
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"保存设置时发生错误: {ex.Message}")
            MessageBox.Show("保存设置失败，请检查日志！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

    Private Sub Buttonlog_Click(sender As Object, e As EventArgs) Handles Buttonlog.Click
        Try
            ' 定义日志文件路径
            Dim logFilePath As String = Path.Combine(Path.GetTempPath(), "OutlookPlugin.log")

            ' 检查日志文件是否存在，如果不存在则创建
            If Not File.Exists(logFilePath) Then
                File.WriteAllText(logFilePath, "日志文件已创建。" & Environment.NewLine)
                Logger.WriteLog(Logger.LogLevel.Info, "FormSetting", "日志文件不存在，已创建新日志文件。")
            End If

            ' 打开日志文件
            Process.Start("notepad.exe", logFilePath)
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"打开日志文件时发生错误: {ex.Message}")
            MessageBox.Show("无法打开日志文件，请检查日志路径或文件权限！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonBrowserpath_Click(sender As Object, e As EventArgs) Handles ButtonBrowserpath.Click
        Try
            Using folderBrowserDialog As New FolderBrowserDialog()
                ' 设置描述文本
                folderBrowserDialog.Description = "请选择归档文件夹"
                folderBrowserDialog.ShowNewFolderButton = True ' 允许创建新文件夹

                ' 设置初始路径
                If Directory.Exists(TextBoxArchivePath.Text) Then
                    folderBrowserDialog.SelectedPath = TextBoxArchivePath.Text
                Else
                    folderBrowserDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If

                ' 显示对话框
                If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
                    ' 将选择的路径更新到文本框和设置
                    TextBoxArchivePath.Text = folderBrowserDialog.SelectedPath
                End If
            End Using
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"选择文件夹路径时发生错误: {ex.Message}")
            MessageBox.Show("选择文件夹路径失败，请检查日志！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CheckBoxdebug_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxdebug.CheckedChanged
        Try
            ' 如果 DebugMode 被勾选，则强制勾选 LogModel，并禁用 CheckBoxlog
            If CheckBoxdebug.Checked Then
                CheckBoxlog.Checked = True
                CheckBoxlog.Enabled = False
                ButtonEmptySetting.Visible = True
                CheckBoxTempMode.Visible = True
                TextBoxOCRRequestAddress.Enabled = True

            Else
                ' 如果取消 DebugMode，则恢复 CheckBoxlog 的可用状态
                CheckBoxlog.Enabled = True
                ButtonEmptySetting.Visible = False
                CheckBoxTempMode.Visible = False
                TextBoxOCRRequestAddress.Enabled = False
            End If
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"切换 DebugMode 时发生错误: {ex.Message}")
        End Try
    End Sub
    Private Sub ButtonEmptySetting_Click(sender As Object, e As EventArgs) Handles ButtonEmptySetting.Click
        My.Settings.AccessToken = String.Empty
        My.Settings.ArchivePath = String.Empty
        My.Settings.ClientId = String.Empty
        My.Settings.ClientSecret = String.Empty
        My.Settings.License = String.Empty
        My.Settings.LicenseData = String.Empty
        My.Settings.LicenseMail = String.Empty
        My.Settings.TempMode = False
        My.Settings.AccessToken = String.Empty
        My.Settings.OCRRequestAddress = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"

        CheckBoxTempMode.Checked = False
        CheckBoxShowToken.Enabled = True
        CheckBoxdebug.Enabled = True

        TextBoxClientId.Text = String.Empty
        TextBoxClientSecret.Text = String.Empty
        TextBoxClientId.Enabled = True
        TextBoxClientSecret.Enabled = True

        Label1.Visible = False
        Label2.Visible = False
        Label4.Visible = False
        LabelLicenseMail.Visible = False
        LabelLicenseData.Visible = False
        LabelOtherinfo.Visible = False




        File.Delete(Path.Combine(Path.GetTempPath(), "OutlookPlugin.log")）
        My.Settings.Save()
        LoadSettings()
        MessageBox.Show("已重置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click

    End Sub

    Private Sub CheckBoxTempMode_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxTempMode.CheckedChanged
        Try
            If CheckBoxTempMode.Checked Then
                EnableTempMode()
            Else
                My.Settings.LicenseMail = String.Empty
                My.Settings.LicenseData = String.Empty
                My.Settings.Save()
                DisableTempMode()
            End If
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"切换临时模式时发生错误: {ex.Message}")
            MessageBox.Show($"切换临时模式时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonEmptySetting_Click()
        Throw New NotImplementedException()
    End Sub

    Private Sub EnableTempMode()
        Try
            ' 如果临时模式已启用且设置完整，直接应用
            If My.Settings.TempMode AndAlso IsTempModeSettingsValid() Then
                ApplyTempModeSettings()
                Return
            End If

            ' 弹出许可输入框
            Dim base64Input As String = InputBox("请输入临时许可（Base64 编码）:", "输入临时许可")
            If String.IsNullOrWhiteSpace(base64Input) Then
                CheckBoxTempMode.Checked = False
                MessageBox.Show("未输入许可，临时模式未启用。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 解析并验证许可
            Try
                ParseAndApplyTempLicense(base64Input)
                Logger.WriteLog(Logger.LogLevel.Info, "FormSetting", "临时模式已成功启用。")
            Catch ex As Exception
                Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"解析临时许可时发生错误: {ex.Message}")
                MessageBox.Show("临时许可无效，请检查后重试！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                CheckBoxTempMode.Checked = False
            End Try
        Catch ex As Exception
            Logger.WriteLog(Logger.LogLevel.Error, "FormSetting", $"启用临时模式时发生错误: {ex.Message}")
            CheckBoxTempMode.Checked = False
        End Try
    End Sub


    ''' <summary>
    ''' 验证临时模式的设置是否完整。
    ''' </summary>
    ''' <returns>如果设置完整，返回 True；否则返回 False。</returns>
    Private Function IsTempModeSettingsValid() As Boolean
        Return Not String.IsNullOrEmpty(My.Settings.ClientId) AndAlso
           Not String.IsNullOrEmpty(My.Settings.ClientSecret) AndAlso
           Not String.IsNullOrEmpty(My.Settings.LicenseMail) AndAlso
           Not String.IsNullOrEmpty(My.Settings.LicenseData)
    End Function



    ''' <summary>
    ''' 解析并应用临时许可。
    ''' </summary>
    ''' <param name="base64Input">临时许可的 Base64 字符串。</param>
    Private Sub ParseAndApplyTempLicense(base64Input As String)
        Try
            Dim decodedString As String = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(base64Input))
            Dim parameters As String() = decodedString.Split("_"c)

            ' 验证参数完整性
            If parameters.Length < 4 Then
                Throw New ArgumentException("许可参数不完整！")
            End If

            ' 应用许可参数
            My.Settings.ClientId = GetValueFromParameters(parameters, "ClientId")
            My.Settings.ClientSecret = GetValueFromParameters(parameters, "ClientSecret")
            My.Settings.LicenseMail = GetValueFromParameters(parameters, "LicenseMail")
            My.Settings.LicenseData = GetValueFromParameters(parameters, "LicenseData")
            My.Settings.TempMode = True
            My.Settings.Save()

            ApplyTempModeSettings()
        Catch ex As Exception
            Throw New ArgumentException("临时许可解析失败: " & ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' 从参数中获取指定键的值。
    ''' </summary>
    ''' <param name="parameters">参数数组。</param>
    ''' <param name="key">键。</param>
    ''' <returns>对应的值。</returns>
    Private Function GetValueFromParameters(parameters As String(), key As String) As String
        Return parameters.FirstOrDefault(Function(p) p.StartsWith($"{key}:"))?.Split(":"c)(1)
    End Function


    ' 用于应用临时模式的设置
    Private Sub ApplyTempModeSettings()
        ' 根据临时许可设置更新 UI 和状态
        TextBoxClientId.Text = My.Settings.ClientId ' 确保加载值
        TextBoxClientSecret.Text = My.Settings.ClientSecret ' 确保加载值
        TextBoxClientId.Enabled = False
        TextBoxClientSecret.Enabled = False
        CheckBoxShowToken.Checked = False
        CheckBoxShowToken.Enabled = False
        CheckBoxdebug.Enabled = False

        Label1.Visible = True
        Label2.Visible = True
        Label4.Visible = True
        LabelLicenseMail.Visible = True
        LabelLicenseData.Visible = True
        LabelOtherinfo.Visible = True
    End Sub


    Private Sub DisableTempMode()
        ' 恢复界面状态
        TextBoxClientId.Enabled = True
        TextBoxClientSecret.Enabled = True

        CheckBoxShowToken.Enabled = True
        CheckBoxdebug.Enabled = True

        ' 清除与临时模式相关的设置
        My.Settings.TempMode = False
        ' 清空与临时模式相关的 UI
        TextBoxClientId.Text = My.Settings.ClientId ' 恢复非临时模式数据
        TextBoxClientSecret.Text = My.Settings.ClientSecret ' 恢复非临时模式数据

        LabelLicenseMail.Text = "未授权"
        LabelLicenseData.Text = "未设置"

        ' 隐藏与临时模式相关的 UI 元素
        Label1.Visible = False
        Label2.Visible = False
        Label4.Visible = False
        LabelLicenseMail.Visible = False
        LabelLicenseData.Visible = False
        LabelOtherinfo.Visible = False

    End Sub





End Class
