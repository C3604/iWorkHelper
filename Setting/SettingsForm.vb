Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports System.Text

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




        Catch ex As system.Exception
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
        Catch ex As system.Exception
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
        Catch ex As system.Exception
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
    Private Async Sub lblVersion_Click(sender As Object, e As EventArgs) Handles lblVersion.Click
        Try
            ' 记录日志
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.lblVersion_Click", "用户点击了版本标签，直接显示更新日志")
            
            ' 临时禁用标签，防止重复点击
            lblVersion.Enabled = False
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.lblVersion_Click", "版本标签已临时禁用")
            
            ' 创建更新检查器并检查更新
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.lblVersion_Click", "创建Checkupdata实例获取更新信息")
            Dim checker As New Checkupdata()
            Dim result = Await checker.CheckForUpdates(FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion)
            
            ' 获取更新日志JSON数据
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.lblVersion_Click", "开始获取更新日志数据")
            Dim changelogJson As String = ""
            
            If result.UpdateInfo IsNot Nothing AndAlso Not String.IsNullOrEmpty(result.UpdateInfo.ChangelogUrl) Then
                changelogJson = Await GetChangelogJsonAsync(result.UpdateInfo.ChangelogUrl)
                
                ' 直接显示更新日志窗口
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.lblVersion_Click", "直接显示更新日志")
                ShowChangelogWindow(changelogJson)
            Else
                ' 如果获取不到更新日志URL，显示错误信息
                LogManager.WriteLog(LogLevel.Error, "SettingsForm.lblVersion_Click", "无法获取更新日志URL")
                MessageBox.Show("无法获取更新日志信息", "更新日志", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            
        Catch ex As Exception
            ' 记录异常
            LogManager.WriteLog(LogLevel.Error, "SettingsForm.lblVersion_Click", "处理更新日志时发生异常: " & ex.Message)
            MessageBox.Show("显示更新日志时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 恢复标签状态
            lblVersion.Enabled = True
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.lblVersion_Click", "版本标签已恢复状态")
        End Try
    End Sub
    
    ' 获取更新日志的JSON数据
    Private Async Function GetChangelogJsonAsync(changelogUrl As String) As Task(Of String)
        Try
            ' 检查URL是否为空，如果为空则使用默认URL
            If String.IsNullOrEmpty(changelogUrl) Then
                changelogUrl = "https://gitee.com/Xoin-Yang/master-updater/raw/master/iWorkHelper/changelog.json"
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", 
                    "更新日志URL为空，使用默认URL: " & changelogUrl)
            End If
            
            ' 使用提供的URL获取更新日志
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", "开始获取更新日志JSON数据: " & changelogUrl)
            
            ' 下载JSON数据
            Using client As New Net.WebClient()
                client.Encoding = System.Text.Encoding.UTF8
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", "开始下载JSON数据...")
                Dim jsonData = Await client.DownloadStringTaskAsync(changelogUrl)
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", 
                    "成功下载JSON数据，长度: " & jsonData.Length & " 字符")
                
                ' 如果获取到的不是完整的changelog数据，可能是更新信息JSON，尝试使用默认URL再次获取
                If jsonData.Contains("""latestVersion""") AndAlso Not jsonData.Contains("""versions""") Then
                    LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", 
                        "获取到的是更新信息JSON而不是更新日志，将尝试使用默认URL")
                    
                    changelogUrl = "https://gitee.com/Xoin-Yang/master-updater/raw/master/iWorkHelper/changelog.json"
                    LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", 
                        "使用默认更新日志URL: " & changelogUrl)
                    
                    ' 重新下载正确的更新日志JSON
                    jsonData = Await client.DownloadStringTaskAsync(changelogUrl)
                    LogManager.WriteLog(LogLevel.INFO, "SettingsForm.GetChangelogJsonAsync", 
                        "使用默认URL成功下载JSON数据，长度: " & jsonData.Length & " 字符")
                End If
                
                Return jsonData
            End Using
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "SettingsForm.GetChangelogJsonAsync", "获取更新日志JSON数据失败: " & ex.Message)
            Return String.Empty
        End Try
    End Function
    
    ' 显示更新日志窗口
    Private Sub ShowChangelogWindow(jsonData As String)
        Try
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ShowChangelogWindow", "开始显示更新日志窗口")
            
            ' 临时解决方案：不使用ChangelogForm类，直接解析并显示更新日志
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ShowChangelogWindow", "开始解析更新日志JSON数据")
            Dim changelog As String = ParseChangelogJson(jsonData)
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ShowChangelogWindow", "成功解析更新日志，准备显示")
            MessageBox.Show(changelog, "更新日志", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ShowChangelogWindow", "更新日志显示完成")
            
            ' 原方案（需要添加ChangelogForm到项目中才能使用）
            ' Dim changelogForm As New ChangelogForm(jsonData)
            ' changelogForm.ShowDialog(Me)
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "SettingsForm.ShowChangelogWindow", "显示更新日志窗口失败: " & ex.Message)
            MessageBox.Show("无法显示更新日志: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' 临时解析changelog JSON数据的函数
    Private Function ParseChangelogJson(jsonData As String) As String
        Try
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "开始解析更新日志数据")
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "完整JSON数据: " & jsonData)
            
            Dim result As New StringBuilder()
            result.AppendLine("软件更新日志")
            result.AppendLine("====================")
            result.AppendLine()
            
            ' 先检查是否为有效JSON
            If String.IsNullOrEmpty(jsonData) OrElse Not (jsonData.Trim().StartsWith("{") AndAlso jsonData.Trim().EndsWith("}")) Then
                LogManager.WriteLog(LogLevel.Error, "SettingsForm.ParseChangelogJson", "JSON格式无效")
                Return "无法解析更新日志: 格式无效"
            End If
            
            ' 查找versions数组
            Dim versionsStartIndex As Integer = jsonData.IndexOf("""versions""")
            
            If versionsStartIndex < 0 Then
                LogManager.WriteLog(LogLevel.Error, "SettingsForm.ParseChangelogJson", "未找到versions数组")
                Return "无法解析更新日志: 未找到版本信息"
            End If
            
            ' 找到versions数组的开始和结束位置
            Dim arrayStartIndex As Integer = jsonData.IndexOf("[", versionsStartIndex)
            Dim arrayEndIndex As Integer = FindMatchingBracket(jsonData, arrayStartIndex)
            
            If arrayStartIndex < 0 OrElse arrayEndIndex < 0 Then
                LogManager.WriteLog(LogLevel.Error, "SettingsForm.ParseChangelogJson", "无法找到versions数组的边界")
                Return "无法解析更新日志: 版本数组格式错误"
            End If
            
            ' 提取versions数组内容
            Dim versionsArrayContent As String = jsonData.Substring(arrayStartIndex + 1, arrayEndIndex - arrayStartIndex - 1)
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "提取的versions数组内容: " & versionsArrayContent)
            
            ' 手动查找和处理每个版本对象
            Dim versionObjects As New List(Of String)()
            Dim currentIndex As Integer = 0
            
            While currentIndex < versionsArrayContent.Length
                ' 跳过空白字符
                While currentIndex < versionsArrayContent.Length AndAlso Char.IsWhiteSpace(versionsArrayContent(currentIndex))
                    currentIndex += 1
                End While
                
                If currentIndex >= versionsArrayContent.Length Then Exit While
                
                If versionsArrayContent(currentIndex) = "{"c Then
                    ' 找到对象的结束位置
                    Dim objectEndIndex As Integer = FindMatchingBracket(versionsArrayContent, currentIndex)
                    If objectEndIndex > 0 Then
                        Dim versionObject As String = versionsArrayContent.Substring(currentIndex, objectEndIndex - currentIndex + 1)
                        versionObjects.Add(versionObject)
                        currentIndex = objectEndIndex + 1
                    Else
                        currentIndex += 1
                    End If
                Else
                    currentIndex += 1
                End If
                
                ' 跳过逗号和空白
                While currentIndex < versionsArrayContent.Length AndAlso 
                      (versionsArrayContent(currentIndex) = ","c OrElse Char.IsWhiteSpace(versionsArrayContent(currentIndex)))
                    currentIndex += 1
                End While
            End While
            
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "找到 " & versionObjects.Count & " 个版本对象")
            
            ' 解析每个版本对象
            For Each versionJson In versionObjects
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "处理版本对象: " & versionJson)
                
                ' 直接查找version字段
                Dim versionStartIndex As Integer = versionJson.IndexOf("""version""")
                Dim versionValueStart As Integer = versionJson.IndexOf(":", versionStartIndex) + 1
                versionValueStart = versionJson.IndexOf("""", versionValueStart) + 1
                Dim versionValueEnd As Integer = versionJson.IndexOf("""", versionValueStart)
                
                ' 直接查找date字段
                Dim dateStartIndex As Integer = versionJson.IndexOf("""date""")
                Dim dateValueStart As Integer = versionJson.IndexOf(":", dateStartIndex) + 1
                dateValueStart = versionJson.IndexOf("""", dateValueStart) + 1
                Dim dateValueEnd As Integer = versionJson.IndexOf("""", dateValueStart)
                
                ' 提取版本号和日期
                Dim version As String = "未知版本"
                Dim releaseDate As String = "未知日期"
                
                If versionStartIndex >= 0 And versionValueStart > 0 And versionValueEnd > versionValueStart Then
                    version = versionJson.Substring(versionValueStart, versionValueEnd - versionValueStart)
                    LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "找到版本号: " & version)
                End If
                
                If dateStartIndex >= 0 And dateValueStart > 0 And dateValueEnd > dateValueStart Then
                    releaseDate = versionJson.Substring(dateValueStart, dateValueEnd - dateValueStart)
                    LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ParseChangelogJson", "找到日期: " & releaseDate)
                End If
                
                ' 添加版本和日期信息
                result.AppendLine($"版本 {version} - {releaseDate}")
                result.AppendLine("--------------------")
                
                ' 查找changes数组
                Dim changesStartIndex As Integer = versionJson.IndexOf("""changes""")
                If changesStartIndex >= 0 Then
                    Dim changesArrayStart As Integer = versionJson.IndexOf("[", changesStartIndex)
                    Dim changesArrayEnd As Integer = FindMatchingBracket(versionJson, changesArrayStart)
                    
                    If changesArrayStart >= 0 AndAlso changesArrayEnd >= 0 Then
                        Dim changesContent As String = versionJson.Substring(changesArrayStart + 1, changesArrayEnd - changesArrayStart - 1)
                        
                        ' 分割changes条目
                        Dim currentPos As Integer = 0
                        Dim inQuotes As Boolean = False
                        Dim changeStart As Integer = 0
                        Dim changes As New List(Of String)()
                        
                        While currentPos < changesContent.Length
                            Dim c As Char = changesContent(currentPos)
                            
                            If c = """"c Then
                                inQuotes = Not inQuotes
                                If inQuotes AndAlso currentPos > 0 AndAlso changesContent(currentPos - 1) <> "\"c Then
                                    changeStart = currentPos + 1
                                ElseIf Not inQuotes AndAlso currentPos > 0 AndAlso changesContent(currentPos - 1) <> "\"c Then
                                    If changeStart < currentPos Then
                                        changes.Add(changesContent.Substring(changeStart, currentPos - changeStart))
                                    End If
                                End If
                            End If
                            
                            currentPos += 1
                        End While
                        
                        ' 输出所有changes
                        For Each change In changes
                            If Not String.IsNullOrEmpty(change) Then
                                result.AppendLine($"• {change}")
                            End If
                        Next
                    End If
                End If
                
                result.AppendLine()
            Next
            
            Return result.ToString()
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "SettingsForm.ParseChangelogJson", "解析更新日志时出错: " & ex.Message & vbCrLf & ex.StackTrace)
            Return "解析更新日志时出错: " & ex.Message
        End Try
    End Function
    
    ' 从JSON对象中提取指定键的值
    Private Function ExtractJsonValue(jsonObject As String, key As String) As String
        Try
            ' 记录完整的输入数据（前100个字符）
            LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", 
                "尝试从JSON中提取键: " & key & ", JSON前100字符: " & jsonObject.Substring(0, Math.Min(100, jsonObject.Length)))
            
            ' 首先尝试使用标准模式匹配
            Dim keyPattern As String = """" & key & """\s*:"
            Dim keyIndex As Integer = jsonObject.IndexOf(keyPattern)
            
            If keyIndex < 0 Then 
                ' 尝试不同的大小写和空格变体
                Dim alternativePatterns As New List(Of String) From {
                    """" & key.ToLower() & """\s*:",
                    """" & key.ToUpper() & """\s*:",
                    """" & key & """\s*\n*\r*:",
                    """" & key.ToLower() & """\s*\n*\r*:",
                    """" & key.ToUpper() & """\s*\n*\r*:"
                }
                
                For Each pattern In alternativePatterns
                    keyIndex = jsonObject.IndexOf(pattern)
                    If keyIndex >= 0 Then
                        LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", 
                            "找到键使用替代模式: " & pattern)
                        Exit For
                    End If
                Next
                
                If keyIndex < 0 Then
                    LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", 
                        "未找到键: " & key & " 在JSON中")
                    Return ""
                End If
            End If
            
            ' 找到了键，现在提取值
            Dim colonIndex As Integer = jsonObject.IndexOf(":", keyIndex)
            If colonIndex < 0 Then
                LogManager.WriteLog(LogLevel.Error, "SettingsForm.ExtractJsonValue", 
                    "找到键但没有找到冒号，键位置: " & keyIndex)
                Return ""
            End If
            
            Dim valueStart As Integer = colonIndex + 1
            Dim valueEnd As Integer = -1
            
            ' 跳过空白字符
            While valueStart < jsonObject.Length AndAlso (Char.IsWhiteSpace(jsonObject(valueStart)) OrElse 
                                                         jsonObject(valueStart) = vbCr OrElse 
                                                         jsonObject(valueStart) = vbLf)
                valueStart += 1
            End While
            
            If valueStart >= jsonObject.Length Then 
                LogManager.WriteLog(LogLevel.Error, "SettingsForm.ExtractJsonValue", "值的起始位置超出了JSON长度")
                Return ""
            End If
            
            ' 检查值的类型
            If jsonObject(valueStart) = """"c Then
                ' 字符串值
                valueStart += 1
                valueEnd = jsonObject.IndexOf("""", valueStart)
                If valueEnd < 0 Then
                    LogManager.WriteLog(LogLevel.ERROR, "SettingsForm.ExtractJsonValue", "字符串值没有闭合的引号")
                    Return ""
                End If
                
                Dim result = jsonObject.Substring(valueStart, valueEnd - valueStart)
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", "提取的字符串值: " & key & " = " & result)
                Return result
            ElseIf jsonObject(valueStart) = "["c Then
                ' 数组
                valueEnd = FindMatchingBracket(jsonObject, valueStart)
                If valueEnd < 0 Then
                    LogManager.WriteLog(LogLevel.ERROR, "SettingsForm.ExtractJsonValue", "数组没有闭合的括号")
                    Return ""
                End If
                
                Dim result = jsonObject.Substring(valueStart, valueEnd - valueStart + 1)
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", "提取的数组值: " & key & " (长度: " & result.Length & ")")
                Return result
            ElseIf jsonObject(valueStart) = "{"c Then
                ' 对象
                valueEnd = FindMatchingBracket(jsonObject, valueStart)
                If valueEnd < 0 Then
                    LogManager.WriteLog(LogLevel.ERROR, "SettingsForm.ExtractJsonValue", "对象没有闭合的括号")
                    Return ""
                End If
                
                Dim result = jsonObject.Substring(valueStart, valueEnd - valueStart + 1)
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", "提取的对象值: " & key & " (长度: " & result.Length & ")")
                Return result
            Else
                ' 数字、布尔值等
                valueEnd = jsonObject.IndexOf(",", valueStart)
                If valueEnd < 0 Then
                    valueEnd = jsonObject.IndexOf("}", valueStart)
                End If
                
                If valueEnd < 0 Then 
                    LogManager.WriteLog(LogLevel.ERROR, "SettingsForm.ExtractJsonValue", "无法确定值的结束位置")
                    Return ""
                End If
                
                Dim result = jsonObject.Substring(valueStart, valueEnd - valueStart).Trim()
                LogManager.WriteLog(LogLevel.INFO, "SettingsForm.ExtractJsonValue", "提取的原始值: " & key & " = " & result)
                Return result
            End If
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "SettingsForm.ExtractJsonValue", "提取值时出错: " & ex.Message & vbCrLf & ex.StackTrace)
            Return ""
        End Try
    End Function
    
    ' 查找匹配的括号
    Private Function FindMatchingBracket(text As String, openIndex As Integer) As Integer
        If openIndex < 0 OrElse openIndex >= text.Length Then Return -1
        
        Dim openChar As Char = text(openIndex)
        Dim closeChar As Char
        
        If openChar = "["c Then
            closeChar = "]"c
        ElseIf openChar = "{"c Then
            closeChar = "}"c
        Else
            Return -1
        End If
        
        Dim depth As Integer = 1
        Dim i As Integer = openIndex + 1
        
        While i < text.Length AndAlso depth > 0
            If text(i) = openChar Then
                depth += 1
            ElseIf text(i) = closeChar Then
                depth -= 1
            End If
            
            i += 1
        End While
        
        If depth = 0 Then
            Return i - 1
        Else
            Return -1
        End If
    End Function
    
    ' 分割JSON对象
    Private Function SplitJsonObjects(json As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim i As Integer = 0
        
        While i < json.Length
            ' 跳过空白字符
            While i < json.Length AndAlso Char.IsWhiteSpace(json(i))
                i += 1
            End While
            
            If i >= json.Length Then Exit While
            
            If json(i) = "{"c Then
                Dim objEnd As Integer = FindMatchingBracket(json, i)
                If objEnd < 0 Then Exit While
                
                result.Add(json.Substring(i, objEnd - i + 1))
                i = objEnd + 1
            Else
                i += 1
            End If
            
            ' 跳过逗号和空白字符
            While i < json.Length AndAlso (json(i) = ","c OrElse Char.IsWhiteSpace(json(i)))
                i += 1
            End While
        End While
        
        Return result
    End Function
    
    ' 分割JSON数组元素
    Private Function SplitJsonArray(json As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim i As Integer = 0
        Dim itemStart As Integer = 0
        Dim inQuotes As Boolean = False
        Dim depth As Integer = 0
        
        While i < json.Length
            Dim c As Char = json(i)
            
            If c = """"c AndAlso (i = 0 OrElse json(i - 1) <> "\"c) Then
                inQuotes = Not inQuotes
            ElseIf Not inQuotes Then
                If c = "{"c OrElse c = "["c Then
                    depth += 1
                ElseIf c = "}"c OrElse c = "]"c Then
                    depth -= 1
                ElseIf c = ","c AndAlso depth = 0 Then
                    result.Add(json.Substring(itemStart, i - itemStart).Trim())
                    itemStart = i + 1
                End If
            End If
            
            i += 1
        End While
        
        ' 添加最后一个元素
        If itemStart < json.Length Then
            result.Add(json.Substring(itemStart).Trim())
        End If
        
        Return result
    End Function

End Class