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
                TextBoxOCRRequestAddress.Enabled = True

            Else
                ' 如果取消 DebugMode，则恢复 CheckBoxlog 的可用状态
                CheckBoxlog.Enabled = True
                ButtonEmptySetting.Visible = False
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

        CheckBoxShowToken.Enabled = True
        CheckBoxdebug.Enabled = True

        TextBoxClientId.Text = String.Empty
        TextBoxClientSecret.Text = String.Empty
        TextBoxClientId.Enabled = True
        TextBoxClientSecret.Enabled = True




        File.Delete(Path.Combine(Path.GetTempPath(), "OutlookPlugin.log")）
        My.Settings.Save()
        LoadSettings()
        MessageBox.Show("已重置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click
        ShowHelpDocument.ShowHelpDocument()
    End Sub

    Private Sub ButtonEmptySetting_Click()
        Throw New NotImplementedException()
    End Sub

    ''' <summary>
    ''' 从参数中获取指定键的值。
    ''' </summary>
    Private Function GetValueFromParameters(parameters As String(), key As String) As String
        Return parameters.FirstOrDefault(Function(p) p.StartsWith($"{key}:"))?.Split(":"c)(1)
    End Function

End Class
