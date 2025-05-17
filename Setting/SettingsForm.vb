Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class SettingsForm
    ' 当用户点击btn_ArchivePath按钮时触发此事件
    Private Sub btn_ArchivePath_Click(sender As Object, e As EventArgs) Handles btn_ArchivePath.Click
        ' 创建文件夹选择对话框
        Dim folderDialog As New FolderBrowserDialog()

        ' 设置初始路径为txt_ArchivePath中的内容
        If Not String.IsNullOrEmpty(txt_ArchivePath.Text) Then
            folderDialog.SelectedPath = txt_ArchivePath.Text
        End If

        ' 显示文件夹选择对话框
        If folderDialog.ShowDialog() = DialogResult.OK Then
            ' 用户选择了文件夹，更新txt_ArchivePath文本框的内容为所选文件夹路径
            txt_ArchivePath.Text = folderDialog.SelectedPath
        End If
    End Sub

    ' 在SettingsForm加载时读取My.Settings中的值并填充控件
    Private Sub SettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' 记录日志：SettingsForm 加载开始
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "SettingsForm 加载开始。")

            ' 从My.Settings读取归档路径并填充到txt_ArchivePath
            If My.Settings("ArchivePath") IsNot Nothing AndAlso Not String.IsNullOrEmpty(My.Settings("ArchivePath").ToString()) Then
                Dim archivePath As String = My.Settings("ArchivePath").ToString()
                txt_ArchivePath.Text = archivePath
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "读取归档路径成功，归档路径：" & archivePath)
            Else
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "归档路径为空或未设置。")
            End If

            ' 从My.Settings读取调试模式设置并填充到ckb_Debug
            If My.Settings("DebugStatus") IsNot Nothing Then
                Dim debugStatus As Boolean = CBool(My.Settings("DebugStatus"))
                ckb_Debug.Checked = debugStatus
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "读取调试模式设置成功，调试模式：" & debugStatus.ToString())
            Else
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "调试模式未设置或为空。")
            End If

            ' 从My.Settings读取合并滴滴文件设置并填充到ckb_MergeDidiFiles
            If My.Settings("MergeDidiFiles") IsNot Nothing Then
                Dim mergeDidiFiles As Boolean = CBool(My.Settings("MergeDidiFiles"))
                ckb_MergeDidiFiles.Checked = mergeDidiFiles
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "读取合并滴滴文件设置成功，是否合并：" & mergeDidiFiles.ToString())
            Else
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "合并滴滴文件设置未设置或为空。")
            End If

            Dim fileVersion As String = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion

            ' 显示版本号
            lblVersion.Text = "版本号: " & fileVersion
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm_Load", "软件" & lblVersion.Text)




        Catch ex As Exception
            ' 记录日志：捕获异常并记录错误信息
            LogManager.WriteLog(LogLevel.Error, "SettingsForm_Load", "发生错误: " & ex.Message)
        End Try
    End Sub


    Private Sub btn_SaveSettings_Click(sender As Object, e As EventArgs) Handles btn_SaveSettings.Click
        ' 记录日志：按钮点击事件开始
        LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "保存设置操作开始。")

        Try
            ' 获取txt_ArchivePath文本框的内容
            Dim archivePath As String = txt_ArchivePath.Text
            LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "获取归档路径： " & archivePath)

            ' 获取ckb_Debug复选框的选中状态
            Dim debugStatus As String = If(ckb_Debug.Checked, "True", "False")
            LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "获取调试模式状态： " & debugStatus)

            ' 获取ckb_MergeDidiFiles复选框的选中状态
            Dim mergeDidiFiles As String = If(ckb_MergeDidiFiles.Checked, "True", "False")
            LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "获取合并滴滴文件状态： " & mergeDidiFiles)

            ' 调试：打印传入的设置
            Console.WriteLine("保存设置：归档路径 = " & archivePath)
            Console.WriteLine("保存设置：调试模式 = " & debugStatus)
            Console.WriteLine("保存设置：合并滴滴文件 = " & mergeDidiFiles)

            ' 定义一个变量，用来记录是否有保存失败的设置
            Dim saveFailed As Boolean = False
            Dim errorMessage As String = ""

            ' 保存归档路径
            LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "尝试保存归档路径： " & archivePath)
            Dim resultArchivePath As Integer = WriteSetting.SaveSetting("ArchivePath", "String", archivePath)
            If resultArchivePath = 0 Then
                saveFailed = True
                errorMessage &= "归档路径设置保存失败！" & Environment.NewLine
                LogManager.WriteLog(LogLevel.Error, "btn_SaveSettings_Click", "归档路径设置保存失败！")
            Else
                LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "归档路径设置保存成功。")
            End If

            ' 保存调试模式
            LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "尝试保存调试模式： " & debugStatus)
            Dim resultDebugStatus As Integer = WriteSetting.SaveSetting("DebugStatus", "Boolean", debugStatus)
            If resultDebugStatus = 0 Then
                saveFailed = True
                errorMessage &= "调试模式设置保存失败！" & Environment.NewLine
                LogManager.WriteLog(LogLevel.Error, "btn_SaveSettings_Click", "调试模式设置保存失败！")
            Else
                LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "调试模式设置保存成功。")
            End If

            ' 保存合并滴滴文件设置
            LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "尝试保存合并滴滴文件设置： " & mergeDidiFiles)
            Dim resultMergeDidiFiles As Integer = WriteSetting.SaveSetting("MergeDidiFiles", "Boolean", mergeDidiFiles)
            If resultMergeDidiFiles = 0 Then
                saveFailed = True
                errorMessage &= "合并滴滴文件设置保存失败！" & Environment.NewLine
                LogManager.WriteLog(LogLevel.Error, "btn_SaveSettings_Click", "合并滴滴文件设置保存失败！")
            Else
                LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "合并滴滴文件设置保存成功。")
            End If

            ' 如果没有保存失败，则退出窗体
            If Not saveFailed Then
                LogManager.WriteLog(LogLevel.INFO, "btn_SaveSettings_Click", "设置保存成功，关闭窗体。")
                Me.Close()  ' 关闭窗体
            Else
                ' 如果有保存失败，则显示错误信息并不退出窗体
                LogManager.WriteLog(LogLevel.Error, "btn_SaveSettings_Click", "设置保存失败，显示错误信息。")
                MessageBox.Show(errorMessage, "设置保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            ' 记录日志：如果发生错误，记录错误信息
            LogManager.WriteLog(LogLevel.Error, "btn_SaveSettings_Click", "发生错误: " & ex.Message)
            MessageBox.Show("保存设置失败，发生错误：" & ex.Message, "保存设置失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' 在SettingsForm中添加重置按钮的点击事件
    Private Sub btn_Reset_Click(sender As Object, e As EventArgs)
        ' 记录日志：按钮点击事件开始
        LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "重置操作开始。")

        Try
            ' 清空txt_ArchivePath文本框的内容
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "清空txt_ArchivePath文本框的内容。")
            txt_ArchivePath.Text = String.Empty

            ' 取消勾选ckb_Debug复选框
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "取消勾选ckb_Debug复选框。")
            ckb_Debug.Checked = False

            ' 清空My.Settings中的ArchivePath设置
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "清空My.Settings中的ArchivePath设置。")
            My.Settings.ArchivePath = String.Empty

            ' 将My.Settings中的DebugStatus设置为False
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "将My.Settings中的DebugStatus设置为False。")
            My.Settings.DebugStatus = False

            ' 保存设置
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "保存设置。")
            My.Settings.Save()

            Lab_PathError.Visible = False


            ' 提示用户重置成功
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "提示用户重置成功。")
            MessageBox.Show("设置已重置到默认值", "重置成功", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' 记录日志：重置操作成功
            LogManager.WriteLog(LogLevel.INFO, "btn_Reset_Click", "重置操作成功。")
        Catch ex As Exception
            ' 记录日志：如果发生错误，记录错误信息
            LogManager.WriteLog(LogLevel.Error, "btn_Reset_Click", "发生错误: " & ex.Message)
            MessageBox.Show("重置失败，发生错误：" & ex.Message, "重置失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    'Private Sub txt_ArchivePath_TextChanged(sender As Object, e As EventArgs) Handles txt_ArchivePath.TextChanged
    '    ' 获取当前文本框内容
    '    Dim folderPath As String = txt_ArchivePath.Text

    '    ' 使用正则表达式验证文件夹路径的格式
    '    If Not IsValidFolderPath(folderPath) Then
    '        ' 如果路径格式不符合要求，显示错误标签
    '        Lab_PathError.Visible = True
    '    Else
    '        ' 如果路径格式正确，隐藏错误标签
    '        Lab_PathError.Visible = False
    '    End If
    'End Sub

    '' 用于检查文件夹路径格式的辅助方法
    'Private Function IsValidFolderPath(path As String) As Boolean
    '    ' 文件夹路径的正则表达式（包含Windows路径格式的完整验证）
    '    Dim pattern As String = "^(?:(?:[a-zA-Z]:\\|\\\\[^<>:""/\\|?*\r\n]+\\[^<>:""/\\|?*\r\n]+))(?:\\[^<>:""/\\|?*\r\n]+)*\\?$"

    '    ' 使用正则表达式验证路径
    '    Dim regex As New Regex(pattern)

    '    ' 返回路径是否匹配正则表达式
    '    Return regex.IsMatch(path)
    'End Function

    ' 版本标签点击事件处理程序
    Private Sub lblVersion_Click(sender As Object, e As EventArgs) Handles lblVersion.Click
        ' 获取程序集版本信息
        Dim assembly As Assembly = Assembly.GetExecutingAssembly()
        Dim fileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location)

        ' 构建详细版本信息
        Dim versionInfo As String = $"程序名称: {fileVersionInfo.ProductName}" & Environment.NewLine &
                                   $"版本号: {fileVersionInfo.ProductVersion}" & Environment.NewLine &
                                   $"文件版本: {fileVersionInfo.FileVersion}" & Environment.NewLine &
                                   $"编译时间: {System.IO.File.GetLastWriteTime(assembly.Location)}"

        ' 显示版本信息
        MessageBox.Show(versionInfo, "版本信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class