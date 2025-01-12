Imports System.Diagnostics
Imports System.Dynamic
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

Public Class MailProgressBarForm
    Private Async Sub MailProgressBarForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim entryIds As List(Of String) = GetSelectedMailEntryIds()

        If entryIds Is Nothing OrElse entryIds.Count = 0 Then
            MessageBox.Show("未选择任何邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.Close()
            Return
        End If

        Dim totalMails As Integer = entryIds.Count
        Dim currentMailIndex As Integer = 0

        For Each entryId As String In entryIds
            Try
                currentMailIndex += 1
                Me.Text = $"共 {totalMails} 封邮件，正在处理第 {currentMailIndex} 封邮件..."
                Me.Refresh()

                ' 异步调用处理单封邮件的方法
                Await ProcessSingleMail(entryId)

            Catch ex As system.Exception
                Logger.WriteLog(Logger.LogLevel.Error, "MailProgressBarForm", $"处理邮件失败，EntryID: {entryId}, 错误: {ex.Message}")
            End Try
        Next

        UpdateProgressBar(100, "处理完成")
        ActivateOrOpenFolder(My.Settings.ArchivePath)
        Me.Close()
    End Sub



    ''' <summary>
    ''' 异步处理单封邮件。
    ''' </summary>
    ''' <param name="entryId">邮件的 EntryID。</param>
    Private Async Function ProcessSingleMail(entryId As String) As Task
        Dim skipMail As Boolean = False ' 标志是否跳过该邮件
        Try
            Dim mailItem As Outlook.MailItem = GetMailItemByEntryID(entryId)
            Dim mailInfo As String = If(mailItem IsNot Nothing, FormatMailInfo(mailItem), $"EntryID: {entryId}")

            ' Step 2.1: 创建临时文件夹
            UpdateProgressBar(0, "开始处理……")
            Dim tempFolderPath As String = Path.Combine(Path.GetTempPath(), entryId)
            Directory.CreateDirectory(tempFolderPath)

            ' Step 2.2: 保存附件
            UpdateProgressBar(30, "正在保存附件……")
            SaveMailAttachments(entryId, tempFolderPath)

            ' Step 2.3: 检查 PDF 文件
            Dim pdfFiles As String() = Directory.GetFiles(tempFolderPath, "*.pdf")
            If pdfFiles Is Nothing OrElse pdfFiles.Length = 0 Then
                Logger.WriteLog(Logger.LogLevel.Warning, "MailProgressBarForm", $"未找到 PDF 文件，跳过邮件: {mailInfo}")
                Return
            End If

            ' Step 2.4: 调用 OCR
            UpdateProgressBar(50, "正在执行 OCR 识别……")
            Dim ocrResultPath As String = Path.Combine(tempFolderPath, "OCR_Result.txt")
            Dim ocrStatus As Integer = Await Task.Run(Function() OCRProcessor.ProcessOCR(tempFolderPath))
            If ocrStatus < 0 Then
                Logger.WriteLog(Logger.LogLevel.Warning, "MailProgressBarForm", $"OCR 识别失败，跳过邮件: {mailInfo}")
                Return
            End If

            ' Step 2.5: 重命名附件
            UpdateProgressBar(70, "正在重命名附件……")
            RenameAttachments(tempFolderPath, ocrResultPath, ocrStatus)

            ' Step 2.6: 移动附件到归档路径
            UpdateProgressBar(80, "正在归档附件……")
            ArchiveAttachments(tempFolderPath)

            UpdateProgressBar(90, "正在清理临时文件……")
            If Not My.Settings.DebugMode Then
                Directory.Delete(tempFolderPath, recursive:=True)
                Logger.WriteLog(Logger.LogLevel.Info, "MailProgressBarForm", $"临时文件夹已删除: {tempFolderPath}")
            Else
                Logger.WriteLog(Logger.LogLevel.Info, "MailProgressBarForm", $"DebugMode 启用，保留临时文件夹: {tempFolderPath}")
            End If
            Logger.WriteLog(Logger.LogLevel.Info, "MailProgressBarForm", $"邮件 {mailInfo} 处理完成。")

            ' 标记邮件为已完成
            If Not skipMail Then
                MarkMailAsCompleted(mailItem)
            Else
                Logger.WriteLog(Logger.LogLevel.Warning, "MailProgressBarForm", $"邮件 {mailInfo} 处理被跳过，未标记为完成")
            End If
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "MailProgressBarForm", $"处理邮件时发生错误: {ex.Message}")
        End Try
    End Function


    ''' <summary>
    ''' 重命名附件文件。
    ''' </summary>
    ''' <param name="tempFolderPath">临时文件夹路径。</param>
    ''' <param name="ocrResultPath">OCR 结果文件路径。</param>
    ''' <param name="ocrStatus">OCR 返回值，用于选择正则规则。</param>
    Private Sub RenameAttachments(tempFolderPath As String, ocrResultPath As String, ocrStatus As Integer)
        Try
            ' 读取 OCR 结果内容
            Dim ocrContent As String = File.ReadAllText(ocrResultPath)

            ' 定义正则规则
            Dim DateRegex1 As String = "[行程起止日期1]{4,8}：?\s*(\d{4}-\d{2}-\d{2})"
            Dim AmountRegex1 As String = "合计(\d+\.\d{2})元"

            Dim DateRegex2 As String = "开票日期：(\d{4})年(\d{2})月(\d{2})日"
            Dim AmountRegex2 As String = "\(小写\)￥(\d+\.\d{2})"


            Dim dateMatch As Match
            Dim amountMatch As Match

            ' 根据 OCR 返回值选择正则
            If ocrStatus = 1 Then
                dateMatch = Regex.Match(ocrContent, DateRegex1)
                amountMatch = Regex.Match(ocrContent, AmountRegex1)
            ElseIf ocrStatus = 2 Then
                dateMatch = Regex.Match(ocrContent, DateRegex2)
                amountMatch = Regex.Match(ocrContent, AmountRegex2)
            Else
                Logger.WriteLog(Logger.LogLevel.Warning, "RenameAttachments", "未知的 OCR 状态值，跳过附件重命名。")
                Return
            End If

            ' 提取日期和金额
            Dim extractedDate As String = Nothing
            Dim extractedAmount As String = Nothing

            If dateMatch.Success Then
                If ocrStatus = 2 Then
                    extractedDate = $"{dateMatch.Groups(1).Value}-{dateMatch.Groups(2).Value}-{dateMatch.Groups(3).Value}"
                Else
                    extractedDate = dateMatch.Groups(1).Value
                End If
            End If

            If amountMatch.Success Then
                extractedAmount = amountMatch.Groups(1).Value
            End If

            ' 检查是否提取到有效数据
            If String.IsNullOrEmpty(extractedDate) OrElse String.IsNullOrEmpty(extractedAmount) Then
                Logger.WriteLog(Logger.LogLevel.Warning, "RenameAttachments", "OCR 结果中未找到有效信息，跳过附件重命名。")
                Return
            End If

            ' 遍历文件并重命名
            For Each filePath As String In Directory.GetFiles(tempFolderPath, "*.pdf")
                Dim originalFileName As String = Path.GetFileName(filePath)
                Dim newFileName As String = $"{extractedDate}_{extractedAmount}_{originalFileName}"
                Dim newFilePath As String = Path.Combine(tempFolderPath, newFileName)
                File.Move(filePath, newFilePath)
                Logger.WriteLog(Logger.LogLevel.Info, "RenameAttachments", $"附件已重命名: {newFileName}")
            Next

        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "RenameAttachments", $"重命名附件时发生错误: {ex.Message}")
        End Try
    End Sub



    ''' <summary>
    ''' 归档附件到指定路径。
    ''' </summary>
    ''' <param name="tempFolderPath">临时文件夹路径。</param>
    Private Sub ArchiveAttachments(tempFolderPath As String)
        Try
            Dim archivePath As String = My.Settings.ArchivePath
            Directory.CreateDirectory(archivePath)

            For Each filePath As String In Directory.GetFiles(tempFolderPath, "*.pdf")
                Dim destinationPath As String = Path.Combine(archivePath, Path.GetFileName(filePath))
                File.Move(filePath, destinationPath)
            Next

        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "MailProgressBarForm", $"归档附件时发生错误: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' 标记邮件为已完成。
    ''' </summary>
    ''' <param name="mailItem">邮件对象。</param>
    Private Sub MarkMailAsCompleted(mailItem As Outlook.MailItem)
        If mailItem Is Nothing Then Return

        Try
            mailItem.MarkAsTask(Outlook.OlMarkInterval.olMarkComplete)
            mailItem.TaskCompletedDate = DateTime.Now
            mailItem.Save()
            Logger.WriteLog(Logger.LogLevel.Info, "MailProgressBarForm", $"邮件已标记为完成: {FormatMailInfo(mailItem)}")
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "MailProgressBarForm", $"标记邮件为完成时发生错误: {ex.Message}")
        End Try
    End Sub


    ''' 更新进度条和文本标签
    Private Sub UpdateProgressBar(progressValue As Integer, progressText As String)
        Try
            ' 检查跨线程调用的情况
            If MailProgressBar.InvokeRequired Then
                MailProgressBar.Invoke(Sub() UpdateProgressBar(progressValue, progressText))
            Else
                ' 更新进度条值
                If progressValue >= MailProgressBar.Minimum AndAlso progressValue <= MailProgressBar.Maximum Then
                    MailProgressBar.Value = progressValue
                Else
                    Logger.WriteLog(Logger.LogLevel.Warning, "UpdateProgressBar", $"进度值超出范围: {progressValue}")
                End If

                ' 更新标签文本
                LabelProgressBar.Text = progressText
            End If
            Me.Refresh()
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "UpdateProgressBar", $"更新进度条和文本时发生错误: {ex.Message}")
        End Try
    End Sub


    ''' 获取当前选择的邮件的 EntryId 列表
    Private Function GetSelectedMailEntryIds() As List(Of String)
        Dim selectedEntryIds As New List(Of String)()

        Try
            ' 获取当前选择的对象集合
            Dim explorer As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()
            Dim selection As Outlook.Selection = explorer.Selection

            ' 遍历选择的对象
            For i As Integer = 1 To selection.Count
                Dim item As Object = selection.Item(i)

                ' 检查对象是否为邮件类型
                If TypeOf item Is Outlook.MailItem Then
                    Dim mailItem As Outlook.MailItem = CType(item, Outlook.MailItem)
                    selectedEntryIds.Add(mailItem.EntryID)
                End If
            Next

        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "GetSelectedMailEntryIds", $"获取选中邮件 EntryId 时出错: {ex.Message}")
        End Try

        Return selectedEntryIds
    End Function

    ''' 根据 EntryID 获取邮件对象。
    ''' 返回对应的 Outlook.MailItem 对象，如果未找到则返回 Nothing。
    Private Function GetMailItemByEntryID(entryId As String) As Outlook.MailItem
        Try
            ' 获取 Outlook 的命名空间对象
            Dim outlookNamespace As Outlook.NameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI")

            ' 使用 EntryID 获取邮件对象
            Dim item As Object = outlookNamespace.GetItemFromID(entryId)

            ' 确保对象是 MailItem 类型
            If TypeOf item Is Outlook.MailItem Then
                Return CType(item, Outlook.MailItem)
            End If

        Catch ex As System.Exception
            ' 记录错误日志
            Logger.WriteLog(Logger.LogLevel.Error, "GetMailItemByEntryID", $"获取邮件失败，EntryID: {entryId}, 错误: {ex.Message}")
        End Try

        ' 如果发生错误或对象不是 MailItem，则返回 Nothing
        Return Nothing
    End Function

    ''' 格式化邮件信息为"标题_收件时间"。
    Private Function FormatMailInfo(mailItem As Outlook.MailItem) As String
        Try
            Dim subject As String = If(mailItem.Subject, "无标题")
            Dim receivedTime As String = mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
            Return $"{subject}_{receivedTime}"
        Catch ex As System.Exception
            Return "未知邮件信息"
        End Try
    End Function




    ''' 保存指定邮件的附件到目标文件夹。
    Private Sub SaveMailAttachments(entryId As String, folderPath As String)
        Try
            ' 根据 EntryId 获取邮件对象
            Dim mailItem As Outlook.MailItem = GetMailItemByEntryID(entryId)
            If mailItem Is Nothing Then
                Logger.WriteLog(Logger.LogLevel.Warning, "SaveMailAttachments", $"未找到邮件，EntryId: {entryId}")
                Return
            End If

            ' 检查邮件是否有附件
            If mailItem.Attachments.Count = 0 Then
                Logger.WriteLog(Logger.LogLevel.Info, "SaveMailAttachments", $"邮件没有附件，EntryId: {entryId}")
                Return
            End If

            ' 遍历附件并保存到目标文件夹
            For i As Integer = 1 To mailItem.Attachments.Count
                Dim attachment As Outlook.Attachment = mailItem.Attachments.Item(i)

                ' 获取附件文件名
                Dim attachmentFileName As String = attachment.FileName

                ' 生成目标文件路径
                Dim savePath As String = Path.Combine(folderPath, attachmentFileName)

                ' 保存附件
                attachment.SaveAsFile(savePath)

                Logger.WriteLog(Logger.LogLevel.Info, "SaveMailAttachments", $"附件已保存: {savePath}")
            Next

        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "SaveMailAttachments", $"保存附件时发生错误: {ex.Message}")
        End Try
    End Sub


    ''' 激活已打开的文件夹窗口或打开新的文件夹窗口。
    Private Sub ActivateOrOpenFolder(folderPath As String)
        Try
            ' 确保路径一致性
            Dim absolutePath As String = Path.GetFullPath(folderPath)

            ' 检查是否有匹配的文件夹窗口
            Dim folderHandle As IntPtr = FindFolderWindow(absolutePath)
            If folderHandle <> IntPtr.Zero Then
                ' 如果窗口已最小化，则恢复窗口
                If IsIconic(folderHandle) Then
                    ShowWindow(folderHandle, SW_RESTORE)
                    Logger.WriteLog(Logger.LogLevel.Info, "AttachmentSaver", $"文件夹窗口已恢复: {absolutePath}")
                End If

                ' 激活窗口
                SetForegroundWindow(folderHandle)
                Logger.WriteLog(Logger.LogLevel.Info, "AttachmentSaver", $"文件夹已激活: {absolutePath}")
            Else
                ' 如果未找到匹配的窗口，则打开文件夹
                Process.Start("explorer.exe", absolutePath)
                Logger.WriteLog(Logger.LogLevel.Info, "AttachmentSaver", $"文件夹已打开: {absolutePath}")
            End If
        Catch ex As System.Exception
            Logger.WriteLog(Logger.LogLevel.Error, "AttachmentSaver", $"激活或打开文件夹时发生错误: {ex.Message}")
        End Try
    End Sub


    ''' 查找与指定路径匹配的文件夹窗口句柄。
    Private Function FindFolderWindow(folderPath As String) As IntPtr
        Dim foundHandle As IntPtr = IntPtr.Zero

        EnumWindows(Function(hwnd, lParam)
                        Dim title As New StringBuilder(256)
                        GetWindowText(hwnd, title, title.Capacity)

                        ' 检查窗口标题是否与目标路径相关
                        If Not String.IsNullOrWhiteSpace(title.ToString()) AndAlso title.ToString().Contains(Path.GetFileName(folderPath)) Then
                            ' 检查窗口是否属于文件夹
                            Dim processId As Integer
                            GetWindowThreadProcessId(hwnd, processId)
                            Dim process As Process = Process.GetProcessById(processId)

                            If process.ProcessName = "explorer" Then
                                ' 进一步验证窗口路径是否与目标一致
                                foundHandle = hwnd
                                Return False ' 停止枚举
                            End If
                        End If
                        Return True ' 继续枚举
                    End Function, IntPtr.Zero)

        Return foundHandle
    End Function


    ''' 检查窗口是否最小化。
    <DllImport("user32.dll")>
    Private Shared Function IsIconic(ByVal hwnd As IntPtr) As Boolean
    End Function


    ''' 显示窗口。
    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function


    ''' 设置窗口到前台。
    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hwnd As IntPtr) As Boolean
    End Function


    ''' 获取窗口标题。
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    End Function


    ''' 获取窗口线程进程 ID。
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function


    ''' 枚举所有窗口。
    <DllImport("user32.dll")>
    Private Shared Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Boolean
    End Function

    Private Delegate Function EnumWindowsProc(ByVal hwnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    Private Const SW_RESTORE As Integer = 9




End Class