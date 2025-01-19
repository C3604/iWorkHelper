Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Public Class Form_Settings

    ' 定义日志文件路径
    Private logFolderPath As String = Path.Combine(Path.GetTempPath(), "OutlookPlugin")
    Private logFilePath As String = Path.Combine(logFolderPath, "OutlookPlugin.log")

    ' 创建文件夹选择对话框并设置归档路径
    Private Sub btnSelectPath_Click(sender As Object, e As EventArgs) Handles btnSelectPath.Click
        Try
            Dim folderDialog As New FolderBrowserDialog()
            If Not String.IsNullOrEmpty(txtArchivePath.Text) AndAlso IO.Directory.Exists(txtArchivePath.Text) Then
                folderDialog.SelectedPath = txtArchivePath.Text
            Else
                folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            If folderDialog.ShowDialog() = DialogResult.OK Then
                txtArchivePath.Text = folderDialog.SelectedPath
            End If
        Catch ex As Exception
            LogError("btnSelectPath_Click", ex)
            MessageBox.Show("选择文件夹时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 单选框事件处理：离线模式选中
    Private Sub radOfflineMode_CheckedChanged(sender As Object, e As EventArgs) Handles radOfflineMode.CheckedChanged
        ' 如果选择了离线模式
        If radOfflineMode.Checked Then
            ' 禁用与百度OCR相关的控件
            txtRequestUrl.Enabled = False
            txtApiKey.Enabled = False
            txtSecretKey.Enabled = False
            chkShowPlainText.Enabled = False
        End If
    End Sub

    ' 单选框事件处理：百度OCR选中
    Private Sub radBaiduOCR_CheckedChanged(sender As Object, e As EventArgs) Handles radBaiduOCR.CheckedChanged
        ' 如果选择了百度OCR
        If radBaiduOCR.Checked Then
            ' 启用与百度OCR相关的控件
            txtRequestUrl.Enabled = True
            txtApiKey.Enabled = True
            txtSecretKey.Enabled = True
            chkShowPlainText.Enabled = True
        End If
    End Sub


    ' 窗体加载时初始化控件状态
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' 加载设置并填充控件
            txtArchivePath.Text = My.Settings.archivePath
            LoadOCRMode(My.Settings.OCRMode)
            txtRequestUrl.Text = My.Settings.RequestUrl
            txtApiKey.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(My.Settings.ApiKey))
            txtSecretKey.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(My.Settings.SecretKey))
            chkDebugMode.Checked = My.Settings.DebugMode
            chkLog.Checked = My.Settings.LogMode
            SetupLogModeForDebug(chkDebugMode.Checked)

            If chkDebugMode.Checked Then
                chkDebugMode.Top = 450
            Else
                chkDebugMode.Top = 465
            End If

            If chkLog.Checked Then
                chkLog.Top = 450
            Else
                chkLog.Top = 465
            End If


        Catch ex As Exception
            LogError("Form_Load", ex)
            MessageBox.Show("加载设置时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 设置OCR模式
    Private Sub LoadOCRMode(ocrMode As String)
        If ocrMode = "LocalMode" Then
            radOfflineMode.Checked = True
        ElseIf ocrMode = "BaiduOCR" Then
            radBaiduOCR.Checked = True
        End If
    End Sub

    ' 设置调试模式下的日志模式
    Private Sub SetupLogModeForDebug(isDebugMode As Boolean)
        If isDebugMode Then
            chkLog.Checked = True
            chkLog.Enabled = False
        Else
            chkLog.Enabled = True
        End If
    End Sub

    ' 处理Debug模式状态变化
    Private Sub chkDebugMode_CheckedChanged(sender As Object, e As EventArgs) Handles chkDebugMode.CheckedChanged
        Try
            If chkDebugMode.Checked Then
                btnReset.Visible = chkDebugMode.Checked
                chkLog.Checked = chkDebugMode.Checked
                chkLog.Enabled = Not chkDebugMode.Checked
                chkDebugMode.Top = 450
            Else
                btnReset.Visible = chkDebugMode.Checked
                chkLog.Enabled = Not chkDebugMode.Checked
                chkDebugMode.Top = 465
            End If

        Catch ex As Exception
            LogError("chkDebugMode_CheckedChanged", ex)
            MessageBox.Show("设置调试模式时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 处理Log模式状态变化
    Private Sub chkLog_CheckedChanged(sender As Object, e As EventArgs) Handles chkLog.CheckedChanged
        Try
            If chkLog.Checked Then
                btnViewLog.Visible = chkLog.Checked
                chkLog.Top = 450
                CheckLogFileExistence()
            Else
                btnViewLog.Visible = chkLog.Checked
                File.Delete(logFilePath)
                chkLog.Top = 465
            End If
        Catch ex As Exception
            LogError("chkLog_CheckedChanged", ex)
            MessageBox.Show("日志设置时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 检查日志文件是否存在
    Private Sub CheckLogFileExistence()
        If Not Directory.Exists(logFolderPath) Then
            Directory.CreateDirectory(logFolderPath)
        End If

        If Not File.Exists(logFilePath) Then
            File.WriteAllText(logFilePath, "日志文件已创建。" & Environment.NewLine)
        End If
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        Try
            ' 获取并保存设置
            My.Settings.archivePath = txtArchivePath.Text
            My.Settings.OCRMode = If(radOfflineMode.Checked, "LocalMode", "BaiduOCR")
            My.Settings.RequestUrl = txtRequestUrl.Text
            My.Settings.ApiKey = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(txtApiKey.Text))
            My.Settings.SecretKey = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(txtSecretKey.Text))
            My.Settings.DebugMode = chkDebugMode.Checked
            My.Settings.LogMode = chkLog.Checked
            My.Settings.Save()

            ' 记录保存的设置内容到日志
            Logger.WriteLog(Logger.LogLevel.Info, "btnConfirm_Click", $"设置已保存：{My.Settings.archivePath}, OCR模式：{My.Settings.OCRMode}，请求地址：{My.Settings.RequestUrl}，APIKEY：{My.Settings.ApiKey}，SecretKey：{My.Settings.SecretKey}，DebugMode：{My.Settings.DebugMode}")

            ' 调用检查方法并获取错误信息
            Dim errorMessages As String = CheckSettingsValidity()

            ' 如果有错误信息，统一提示用户
            If Not String.IsNullOrEmpty(errorMessages) Then
                MessageBox.Show(errorMessages, "设置错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ' 关闭窗口
            Me.Close()
        Catch ex As Exception
            LogError("btnConfirm_Click", ex)
            MessageBox.Show("保存设置时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 检查设置的有效性，路径和网址格式
    Private Function CheckSettingsValidity() As String
        Dim errorMessages As New Text.StringBuilder()

        ' 检查文件夹路径格式
        If String.IsNullOrEmpty(My.Settings.archivePath) Then
            errorMessages.AppendLine("设置已保存，但归档路径为空，将无法正常保存发票文件")
        ElseIf Not Directory.Exists(My.Settings.archivePath) Then
            errorMessages.AppendLine("设置已保存，但归档路径无效，将无法正常保存发票文件")
        End If

        ' 检查BaiduOCR的时候检查内容格式
        If My.Settings.OCRMode = "BaiduOCR" Then
            If String.IsNullOrEmpty(My.Settings.RequestUrl) Then
                errorMessages.AppendLine("设置已保存，但请求地址不能为空，请填写网址。")
            Else
                Dim uriResult As Uri = Nothing
                If Not Uri.TryCreate(My.Settings.RequestUrl, UriKind.Absolute, uriResult) Then
                    errorMessages.AppendLine("设置已保存，但请求地址无效，将无法正常识别发票文件。")
                End If
            End If

            If String.IsNullOrEmpty(My.Settings.ApiKey) Then
                errorMessages.AppendLine("设置已保存，但Api Key为空，将无法正常识别发票文件")
            End If

            If String.IsNullOrEmpty(My.Settings.SecretKey) Then
                errorMessages.AppendLine("设置已保存，但Secret Key为空，将无法正常识别发票文件")
            End If




        End If

        ' 返回收集到的错误信息
        Return errorMessages.ToString()
    End Function

    ' 记录错误日志
    Private Sub LogError(methodName As String, ex As Exception)
        Logger.WriteLog(Logger.LogLevel.Error, methodName, $"{methodName} 发生错误: {ex.Message}")
    End Sub

    Private Sub btnViewLog_Click(sender As Object, e As EventArgs) Handles btnViewLog.Click
        Try
            ' 定义日志文件路径
            Dim logFolderPath As String = Path.Combine(Path.GetTempPath(), "OutlookPlugin")
            Dim logFilePath As String = Path.Combine(logFolderPath, "OutlookPlugin.log")

            ' 检查日志文件是否存在
            If File.Exists(logFilePath) Then
                ' 使用默认的文本编辑器（如记事本）打开日志文件
                Process.Start("notepad.exe", logFilePath)
            Else
                ' 如果日志文件不存在，提示用户
                MessageBox.Show("日志文件不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            ' 异常处理
            LogError("chkLog_CheckedChanged", ex)
            MessageBox.Show("打开日志文件时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' chkShowPlainText 勾选状态改变时处理
    Private Sub chkShowPlainText_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowPlainText.CheckedChanged
        If chkShowPlainText.Checked Then
            ' 勾选时，显示为明文
            txtApiKey.PasswordChar = Char.MinValue
            txtSecretKey.PasswordChar = Char.MinValue
        Else
            ' 取消勾选时，显示为密码形式
            txtApiKey.PasswordChar = "*"c  ' 或者其他你喜欢的字符，通常是 * 
            txtSecretKey.PasswordChar = "*"c  ' 或者其他你喜欢的字符，通常是 *
        End If
    End Sub

    ' 点击“重置”按钮时清空所有设置
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Try
            ' 显示确认对话框，提示用户是否确定要重置所有设置
            Dim result As DialogResult = MessageBox.Show("确定要重置所有设置吗?", "确认重置", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                ' 清空 My.Settings 中的所有设置（手动）
                My.Settings.archivePath = String.Empty
                My.Settings.OCRMode = "radLocalMode" ' 设置为默认值
                My.Settings.RequestUrl = String.Empty
                My.Settings.ApiKey = String.Empty
                My.Settings.SecretKey = String.Empty
                My.Settings.DebugMode = False
                My.Settings.LogMode = False

                ' 重新加载窗体界面，恢复到默认状态
                LoadDefaultSettings()

                ' 保存设置
                My.Settings.Save()

                ' 提示用户已重置
                MessageBox.Show("设置已重置为默认值！", "重置成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            ' 异常处理
            MessageBox.Show("重置设置时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 重新加载窗体界面为默认状态
    Private Sub LoadDefaultSettings()
        ' 归档位置
        txtArchivePath.Text = String.Empty

        ' OCR模式
        radOfflineMode.Checked = True ' 或者根据默认值设置
        radBaiduOCR.Checked = False

        ' 请求地址
        txtRequestUrl.Text = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"

        ' API Key 和 Secret Key
        txtApiKey.Text = String.Empty
        txtSecretKey.Text = String.Empty

        ' 调试模式
        chkDebugMode.Checked = False

        ' 记录日志
        chkLog.Checked = False

        ' 重置其他控件状态（如果有）
        chkLog.Enabled = True ' 启用日志复选框
    End Sub

    ' 在窗体的代码中处理 HelpButtonClicked 事件
    Private Sub Form_Settings_HelpButtonClicked(sender As Object, e As CancelEventArgs) Handles Me.HelpButtonClicked
        Try
            ShowHelpDocument.ShowHelpDocument()
        Catch ex As Exception
            MessageBox.Show("无法打开帮助文件: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub



End Class