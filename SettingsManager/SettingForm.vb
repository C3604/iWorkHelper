Imports System.IO
Imports System.Windows.Forms
Imports iWorkHelper.LogManager

Public Class SettingForm




    ' 统一日志实例（仅此窗体内使用），避免散落的 MessageBox 调用
    Private Shared _logger As New Logger()

    Private Sub SettingForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 启动日志后台写入（关键路径）
        Try
            _logger.Start()
        Catch
        End Try
        ' 在窗口加载时，将设置项读取到文本框
        Dim tmpText As String = String.Empty
        Try
            Dim obj = My.Settings("tmppath")
            tmpText = If(TryCast(obj, String), String.Empty)
        Catch
            tmpText = String.Empty
        End Try
        txt_tmppath.Text = tmpText

        Dim archText As String = String.Empty
        Try
            Dim obj2 = My.Settings("archivepath")
            archText = If(TryCast(obj2, String), String.Empty)
        Catch
            archText = String.Empty
        End Try
        txt_archivepath.Text = archText
    End Sub

    Private Sub btn_archive_Click(sender As Object, e As EventArgs) Handles btn_archive.Click, btn_temp.Click
        ' 打开文件夹浏览对话框
        If FolderBrowser_setting.ShowDialog() = DialogResult.OK Then
            Dim selectedPath As String = FolderBrowser_setting.SelectedPath

            ' 检查路径是否存在
            If Directory.Exists(selectedPath) Then
                ' 根据触发的按钮写入到相应文本框
                If sender Is btn_temp Then
                    txt_tmppath.Text = selectedPath
                Else
                    txt_archivepath.Text = selectedPath
                End If
            Else
                ' 由日志系统接管警告弹窗
                _logger.LogWarn("选择的文件夹不存在，请重新选择。")
            End If
        End If
    End Sub

    Private Sub btn_accept_Click(sender As Object, e As EventArgs) Handles btn_accept.Click
        ' 关键路径：先保存临时路径到 My.Settings（字符串 tmppath），再处理归档路径
        Dim tmpPath As String = If(txt_tmppath.Text, String.Empty).Trim()
        Try
            ' 保存 txt_tmppath 到 My.Settings，变量名为 tmppath，类型为 String
            SettingsWriter.WriteSetting("tmppath", "String", tmpPath)
        Catch
            ' 忽略异常，不影响归档路径的保存流程
        End Try

        ' 保存 tmppath 后重启 Logger（先 Stop 再 New+Start），避免频繁读取设置
        Try
            _logger.Stop()
        Catch
        End Try
        Try
            _logger = New Logger()
            _logger.Start()
        Catch
        End Try

        ' 获取文本框内容、校验、写入 My.Settings 并反馈（archivepath）
        Dim inputPath As String = If(txt_archivepath.Text, String.Empty).Trim()

        ' 1) 非空校验
        If String.IsNullOrWhiteSpace(inputPath) Then
            _logger.LogWarn("路径不能为空，请输入有效的归档路径。")
            Exit Sub
        End If

        ' 2) 格式校验（不校验是否存在，仅校验字符串有效性）
        If Not IsValidPathFormat(inputPath) Then
            _logger.LogWarn("路径格式无效，请输入形如 C:\\... 或 \\服务器\\共享 的有效路径。")
            Exit Sub
        End If

        ' 3) 写入 My.Settings.archivepath
        Try
            Dim result = SettingsWriter.WriteSetting("archivepath", "String", inputPath)
            If result.Item1 Then
                ' 成功提示改为 Info 级日志（不弹窗）
                _logger.LogInfo("归档路径已保存。")
                ' 完成工作后关闭设置窗体
                Me.Close()
            Else
                _logger.LogError($"保存失败：{result.Item2}")
            End If
        Catch ex As Exception
            _logger.LogError($"写入过程出现异常：{ex.Message}")
        End Try


    End Sub

    ' 路径格式校验：必须为绝对路径且不包含无效字符（不检查存在性）
    Private Function IsValidPathFormat(ByVal pathText As String) As Boolean
        Try
            If String.IsNullOrWhiteSpace(pathText) Then Return False
            ' 必须是绝对路径（支持盘符路径或 UNC 路径）
            If Not System.IO.Path.IsPathRooted(pathText) Then Return False

            ' 不包含系统定义的无效字符
            Dim invalidChars = System.IO.Path.GetInvalidPathChars()
            For Each ch In invalidChars
                If pathText.Contains(ch) Then Return False
            Next

            ' 简单的盘符开头合法性（如 C:\）或 UNC 前缀（如 \\server\share）
            Dim rootedDrive As Boolean = (pathText.Length >= 3 AndAlso Char.IsLetter(pathText(0)) AndAlso pathText(1) = ":"c AndAlso (pathText(2) = "\"c Or pathText(2) = "/"c))
            Dim isUNC As Boolean = pathText.StartsWith("\\")
            If Not (rootedDrive OrElse isUNC) Then Return False

            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub SettingForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ' 停止日志后台任务
        Try
            _logger.Stop()
        Catch
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub
End Class