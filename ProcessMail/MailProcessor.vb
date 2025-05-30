Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Threading
Imports iTextSharp.text.pdf

Public Class MailProcessor
    Private Delegate Sub CloseFormDelegate()
    Public Sub ProcessMails()
        Try
            ' 初始化日志
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", "开始处理邮件")

            Dim getMailID As New GetMailID()
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", "创建 GetMailID 实例")

            Try
                getMailID.ProcessSelectedEmails()
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", "处理已选中的邮件")
            Catch ex As system.Exception
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"处理选中邮件时出错：{ex.Message}")
                MessageBox.Show($"处理选中邮件时出错：{ex.Message}，程序将继续处理可访问的邮件", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ' 继续执行，不要终止整个流程
            End Try

            Dim mailIDs As List(Of String) = getMailID.GetMailWithAttachments()
            Dim totalMails As Integer = mailIDs.Count
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"获取邮件ID成功，共 {totalMails} 封邮件")

            If totalMails = 0 Then
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", "没有找到包含附件的邮件")
                MessageBox.Show("没有找到包含 PDF 附件的邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 创建并显示进度条窗体，启动独立线程
            Dim progressForm As New PdfPigProgressBarForm(totalMails)
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", "创建进度条窗体")

            Dim progressThread As New Thread(Sub() progressForm.ShowDialog())
            progressThread.Start()
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", "启动进度条线程")

            ' 遍历邮件ID
            For i As Integer = 0 To mailIDs.Count - 1
                Dim entryID As String = mailIDs(i)
                Dim tempFolder As String = String.Empty

                Try
                    ' 更新窗体标题，显示当前正在处理的邮件
                    progressForm.UpdateWindowTitle(i + 1, totalMails)
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"正在处理第 {i + 1} 封邮件，EntryID: {entryID}")

                    ' 创建临时文件夹
                    tempFolder = CreateTempFolder(entryID)

                    ' 第1步：保存附件 (25%)
                    progressForm.UpdateStepStatus("保存附件中...")
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"保存附件，EntryID: {entryID}")

                    Try
                        SaveAttachments(entryID, tempFolder)
                        progressForm.UpdateMailProgress(25) ' 第一阶段完成：25%
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"附件保存完毕，EntryID: {entryID}")
                    Catch ex As system.Exception
                        LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"保存附件时出错：{ex.Message}")
                        progressForm.UpdateStepStatus($"保存附件时出错：{ex.Message}")
                        Continue For ' 跳过当前邮件，继续处理下一封
                    End Try

                    ' 第2步：OCR识别和文件重命名 (50%)
                    progressForm.UpdateStepStatus("正在进行OCR识别和重命名...")
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"开始进行OCR识别和重命名，EntryID: {entryID}")

                    Try
                        ProcessPdfFiles(tempFolder) ' OCR处理和文件重命名在这个方法中完成
                        progressForm.UpdateMailProgress(50) ' 第二阶段完成：50%
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"OCR识别和重命名完毕，EntryID: {entryID}")
                    Catch ex As system.Exception
                        LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"OCR识别和重命名时出错：{ex.Message}")
                        progressForm.UpdateStepStatus($"OCR识别和重命名时出错：{ex.Message}")
                        ' 不退出，尝试继续下一步
                    End Try

                    ' 第3步：合并滴滴文件（如果启用）(75%)
                    If My.Settings.MergeDidiFiles Then
                        progressForm.UpdateStepStatus("正在合并滴滴文件...")
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"开始合并滴滴文件，EntryID: {entryID}")
                        Try
                            If MergeDidiPdfFiles(tempFolder) Then
                                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"滴滴文件合并成功，EntryID: {entryID}")
                            Else
                                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"无需合并滴滴文件或合并失败，EntryID: {entryID}")
                            End If
                        Catch ex As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"合并滴滴文件时出错：{ex.Message}")
                            progressForm.UpdateStepStatus($"合并滴滴文件时出错：{ex.Message}")
                            ' 继续执行，不要终止流程
                        End Try
                    End If
                    progressForm.UpdateMailProgress(75) ' 第三阶段完成：75%

                    ' 第4步：归档附件和清理临时文件 (100%)
                    progressForm.UpdateStepStatus("正在归档附件...")
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"开始归档附件，EntryID: {entryID}")
                    Try
                        ' 调用 MarkerMail 方法来归档附件
                        Dim outlookApp As Outlook.Application = Nothing
                        Dim mailItem As Outlook.MailItem = Nothing
                        Try
                            outlookApp = New Outlook.Application()
                            mailItem = outlookApp.Session.GetItemFromID(entryID) ' 获取邮件项
                            If MarkerMail(tempFolder, mailItem) Then
                                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"附件归档成功，EntryID: {entryID}")
                            Else
                                progressForm.UpdateStepStatus("归档失败: 移动文件失败")
                                LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"附件归档失败，EntryID: {entryID}，移动文件失败")
                            End If
                        Catch ex As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"获取或处理邮件对象时出错：{ex.Message}")
                            progressForm.UpdateStepStatus($"处理邮件对象时出错：{ex.Message}")
                        Finally
                            If mailItem IsNot Nothing Then
                                Try
                                    Marshal.ReleaseComObject(mailItem)
                                Catch ex As system.Exception
                                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"释放邮件对象时出错：{ex.Message}")
                                End Try
                                mailItem = Nothing
                            End If

                            If outlookApp IsNot Nothing Then
                                Try
                                    Marshal.ReleaseComObject(outlookApp)
                                Catch ex As system.Exception
                                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"释放Outlook应用程序对象时出错：{ex.Message}")
                                End Try
                                outlookApp = Nothing
                            End If
                        End Try
                    Catch ex As system.Exception
                        progressForm.UpdateStepStatus("归档失败: " & ex.Message)
                        LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"归档失败，EntryID: {entryID}，错误信息: {ex.Message}")
                    End Try

                    ' 清理临时文件
                    progressForm.UpdateStepStatus("清理临时文件...")
                    Try
                        SafeDeleteDirectory(tempFolder)
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"清理临时文件夹：{tempFolder}")
                    Catch ex As system.Exception
                        LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"清理临时文件夹出错：{ex.Message}")
                        progressForm.UpdateStepStatus($"清理临时文件夹出错：{ex.Message}")
                    End Try

                    progressForm.UpdateMailProgress(100) ' 第四阶段完成：100%
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", $"清理完毕，EntryID: {entryID}")

                    ' 处理完每一封邮件后，进度条归0，准备下一封
                    progressForm.UpdateMailProgress(0)
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"处理邮件时出现未处理的异常：{ex.Message}，EntryID: {entryID}")
                    progressForm.UpdateStepStatus($"处理邮件时出现错误：{ex.Message}")

                    ' 尝试清理临时文件夹
                    If Not String.IsNullOrEmpty(tempFolder) AndAlso Directory.Exists(tempFolder) Then
                        Try
                            SafeDeleteDirectory(tempFolder)
                        Catch cleanEx As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"清理临时文件夹失败：{cleanEx.Message}")
                        End Try
                    End If

                    progressForm.UpdateMailProgress(0) ' 重置进度条
                End Try
            Next

            ' 所有邮件处理完毕后，关闭进度条窗体
            Try
                If progressForm.InvokeRequired Then
                    ' 使用委托来跨线程关闭窗体
                    progressForm.Invoke(New CloseFormDelegate(AddressOf progressForm.Close))
                Else
                    progressForm.Close()
                End If
            Catch ex As system.Exception
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"关闭进度条窗体时出错：{ex.Message}")
            End Try

            ' 终止进度条线程
            Try
                progressThread.Abort()
            Catch ex As system.Exception
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"终止进度条线程时出错：{ex.Message}")
            End Try

            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.ProcessMails", "所有邮件处理完毕，进度条关闭")
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"邮件处理主流程发生严重错误：{ex.Message}")
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.ProcessMails", $"错误详情：{ex.StackTrace}")
            MessageBox.Show($"邮件处理过程中发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 创建临时文件夹，使用更短的名称
    Private Function CreateTempFolder(entryID As String) As String
        ' 使用哈希值缩短文件夹名称
        Dim tempFolderName As String = "iWorkHelper_" & Math.Abs(entryID.GetHashCode()).ToString()
        Dim tempFolder As String = Path.Combine(Environment.GetEnvironmentVariable("TEMP"), tempFolderName)

        If Not Directory.Exists(tempFolder) Then
            Directory.CreateDirectory(tempFolder)
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.CreateTempFolder", $"创建临时文件夹：{tempFolder}，对应EntryID: {entryID}")
        End If

        Return tempFolder
    End Function

    ' 保存邮件附件到指定文件夹
    Private Sub SaveAttachments(entryID As String, folderPath As String)
        ' 连接Outlook并获取邮件对象
        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"开始连接Outlook，获取邮件：EntryID = {entryID}")
        Dim outlookApp As New Outlook.Application()
        Dim mailItem As Outlook.MailItem = Nothing
        Try
            mailItem = outlookApp.Session.GetItemFromID(entryID)
            Dim subject As String = mailItem.Subject
            Dim receivedTime As String = mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"成功获取邮件：EntryID = {entryID}，主题：{subject}，收件时间：{receivedTime}")
        Catch ex As System.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.SaveAttachments", $"获取邮件失败：EntryID = {entryID}，错误信息：{ex.Message}")
            Exit Sub
        End Try

        ' 遍历邮件附件并保存到指定文件夹
        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"开始遍历邮件附件，EntryID = {entryID}")
        For Each attachment As Outlook.Attachment In mailItem.Attachments
            ' 仅保存PDF文件
            If attachment.FileName.ToLower().EndsWith(".pdf") Then
                Dim filePath As String = Path.Combine(folderPath, attachment.FileName)
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"发现PDF附件，文件名：{attachment.FileName}，准备保存到：{filePath}")
                Try
                    attachment.SaveAsFile(filePath)
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"附件保存成功：{attachment.FileName}，路径：{filePath}，EntryID：{entryID}，主题：{mailItem.Subject}")
                Catch ex As System.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.SaveAttachments", $"保存附件失败：文件名 = {attachment.FileName}，路径 = {filePath}，错误信息：{ex.Message}，EntryID：{entryID}，主题：{mailItem.Subject}")
                End Try
            Else
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"跳过非PDF附件：{attachment.FileName}，EntryID：{entryID}，主题：{mailItem.Subject}")
            End If
        Next

        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SaveAttachments", $"邮件附件保存完成，EntryID = {entryID}，主题：{mailItem.Subject}")
    End Sub


    ' 处理文件夹中的PDF文件并执行OCR
    Private Sub ProcessPdfFiles(folderPath As String)
        ' 获取文件夹中的所有PDF文件
        Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
        LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", $"找到{pdfFiles.Length}个PDF文件，路径：{folderPath}")

        ' 遍历每个PDF文件，执行OCR
        For Each pdfFile As String In pdfFiles
            LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", $"开始处理PDF文件：{Path.GetFileName(pdfFile)}，路径：{pdfFile}")

            Dim ocrProcessor As PdfpigOcr = Nothing
            Try
                ocrProcessor = New PdfpigOcr()
                ' 调用实例方法
                LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", $"执行OCR：{Path.GetFileName(pdfFile)}")
                Dim ocrResult As String = ocrProcessor.ExtractTextFromPdf(pdfFile)

                ' 将OCR结果保存为txt文件，文件名与源PDF文件名一致
                Dim txtFile As String = Path.Combine(folderPath, Path.GetFileNameWithoutExtension(pdfFile) & ".txt")
                File.WriteAllText(txtFile, ocrResult)
                LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", $"OCR处理完成，结果保存为文本文件：{txtFile}")
            Catch ex As System.Exception
                LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ProcessPdfFiles", $"OCR处理失败：{Path.GetFileName(pdfFile)}，错误信息：{ex.Message}")
            Finally
                ocrProcessor = Nothing
            End Try
        Next

        ' 使用通用方法检查是否包含滴滴文件
        Dim fileInfo As DidiFileInfo = CheckDidiFiles(folderPath)

        ' 根据是否包含滴滴行程单来决定调用哪个重命名方法
        Dim renameResult As Boolean = False
        If fileInfo.HasTripFile Then
            ' 如果包含，执行DidiRename子程序
            LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", "检测到滴滴行程单，调用DidiRename方法进行文件重命名")
            renameResult = DidiRename(folderPath)
            If renameResult Then
                LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", "DidiRename重命名成功，无需后续重命名操作")
            Else
                LogManager.WriteLog(LogLevel.Error, "PdfProcessor.ProcessPdfFiles", "DidiRename重命名失败，但将继续处理流程")
            End If
        Else
            ' 如果不包含，执行Rename子程序
            LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", "没有检测到滴滴行程单，调用通用Rename方法进行文件重命名")
            Dim renameInfo As String = Rename(folderPath)
            LogManager.WriteLog(LogLevel.INFO, "PdfProcessor.ProcessPdfFiles", $"Rename方法执行结果：{renameInfo}，无需后续重命名操作")
        End If
    End Sub

    ' 用于存储滴滴文件检查结果的结构
    Private Structure DidiFileInfo
        Public HasInvoiceFile As Boolean
        Public HasTripFile As Boolean
        Public InvoiceFilePath As String
        Public TripFilePath As String
    End Structure

    ' 通用方法：检查文件夹中是否包含滴滴发票和行程单
    Private Function CheckDidiFiles(folderPath As String) As DidiFileInfo
        Dim result As New DidiFileInfo With {
            .HasInvoiceFile = False,
            .HasTripFile = False,
            .InvoiceFilePath = "",
            .TripFilePath = ""
        }

        ' 获取文件夹中的所有PDF文件
        Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.CheckDidiFiles", $"检查文件夹：{folderPath}，共有{pdfFiles.Length}个PDF文件")

        ' 检查是否包含滴滴相关文件
        For Each pdfFile As String In pdfFiles
            Dim fileName As String = Path.GetFileName(pdfFile).ToLower()

            ' 检查是否为滴滴发票
            If fileName.Contains("滴滴") AndAlso
               (fileName.Contains("发票") OrElse fileName.Contains("电子发票")) AndAlso
               Not fileName.Contains("合并") Then
                result.HasInvoiceFile = True
                result.InvoiceFilePath = pdfFile
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.CheckDidiFiles", $"检测到滴滴发票文件：{fileName}")
            End If

            ' 检查是否为滴滴行程单
            If fileName.Contains("滴滴") AndAlso
               (fileName.Contains("行程") OrElse fileName.Contains("行程单")) AndAlso
               Not fileName.Contains("合并") Then
                result.HasTripFile = True
                result.TripFilePath = pdfFile
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.CheckDidiFiles", $"检测到滴滴行程单文件：{fileName}")
            End If
        Next

        Return result
    End Function

    ' DidiRename 读取文件夹内的 "滴滴出行行程报销单.txt" 文件并提取金额和日期
    Public Function DidiRename(folderPath As String) As Boolean
        ' 在一个方法中处理所有OCR文本文件，避免多次读取文件夹
        Dim txtFiles As String() = Directory.GetFiles(folderPath, "*.txt")
        Dim tripTxtFile As String = txtFiles.FirstOrDefault(Function(f) Path.GetFileName(f).ToLower().Contains("滴滴") AndAlso Path.GetFileName(f).ToLower().Contains("行程"))

        If String.IsNullOrEmpty(tripTxtFile) Then
            LogManager.WriteLog(LogLevel.Error, "FileProcessor.DidiRename", $"未找到滴滴行程单文本文件")
            Return False
        End If

        LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"找到滴滴行程单文本文件：{Path.GetFileName(tripTxtFile)}")

        ' 读取文件内容
        Dim fileContent As String = File.ReadAllText(tripTxtFile)
        LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"成功读取文件内容：{tripTxtFile}")

        ' 改进的正则表达式，更可靠地匹配金额和日期
        Dim amountPattern As String = "合计\s*([\d.]+)\s*元"
        Dim datePattern As String = "行程起止日期[:：]\s*(\d{4}-\d{2}-\d{2})"

        ' 使用正则表达式提取金额和日期
        Dim amountMatch As Match = Regex.Match(fileContent, amountPattern)
        Dim dateMatch As Match = Regex.Match(fileContent, datePattern)

        ' 如果没有找到日期，尝试其他格式
        If Not dateMatch.Success Then
            ' 尝试其他可能的日期格式
            datePattern = "(\d{4}[-/年]\d{1,2}[-/月]\d{1,2})[日\s]?"
            dateMatch = Regex.Match(fileContent, datePattern)
        End If

        ' 检查是否成功提取金额和日期
        If amountMatch.Success AndAlso dateMatch.Success Then
            Dim amount As String = amountMatch.Groups(1).Value
            Dim transactionDate As String = dateMatch.Groups(1).Value

            ' 标准化日期格式
            If transactionDate.Contains("年") Then
                transactionDate = transactionDate.Replace("年", "-").Replace("月", "-").Replace("日", "")
            End If

            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"提取成功：金额 = {amount}，日期 = {transactionDate}")

            ' 获取文件夹中的所有PDF文件
            Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"找到{pdfFiles.Length}个PDF文件，开始重命名")

            ' 对每个PDF文件添加日期和金额前缀
            Dim successCount As Integer = 0
            For Each pdfFile As String In pdfFiles
                ' 跳过已经合并或已经有日期和金额前缀的文件
                Dim fileName = Path.GetFileName(pdfFile)
                If fileName.Contains("合并") OrElse Regex.IsMatch(fileName, "^\d{4}-\d{2}-\d{2}_\d+\.\d{2}_") Then
                    LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"跳过已处理的文件：{fileName}")
                    Continue For
                End If

                ' 创建新文件名：日期_金额_源文件名.pdf
                Dim newFileName As String = $"{transactionDate}_{amount}_{fileName}"
                Dim newFilePath As String = Path.Combine(folderPath, newFileName)

                ' 如果目标路径中已经存在同名文件，则修改文件名
                If File.Exists(newFilePath) Then
                    newFilePath = Path.Combine(folderPath, $"{Path.GetFileNameWithoutExtension(pdfFile)}_new{Path.GetExtension(pdfFile)}")
                    LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"目标文件已存在，重命名为：{Path.GetFileName(newFilePath)}")
                End If

                ' 尝试重命名文件
                Try
                    File.Move(pdfFile, newFilePath)
                    LogManager.WriteLog(LogLevel.INFO, "FileProcessor.DidiRename", $"成功重命名文件：{fileName} -> {Path.GetFileName(newFilePath)}")
                    successCount += 1
                Catch ex As System.Exception
                    LogManager.WriteLog(LogLevel.Error, "FileProcessor.DidiRename", $"重命名文件时发生错误：{fileName} -> {Path.GetFileName(newFilePath)}，错误信息：{ex.Message}")
                End Try
            Next

            Return successCount > 0 ' 如果至少有一个文件重命名成功，则返回True
        Else
            ' 如果未成功提取金额或日期，记录日志并返回False
            If Not amountMatch.Success Then
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.DidiRename", "未能从文本中提取金额信息")
            End If
            If Not dateMatch.Success Then
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.DidiRename", "未能从文本中提取日期信息")
            End If
            Return False
        End If
    End Function

    ' 默认的重命名方法
    Function Rename(folderPath As String) As String
        ' 日期正则表达式：匹配格式为"2023年5月6日"或"2023-5-6"之类的日期
        Dim datePattern As String = "(20\d{2})\s*（?年?）?(\d{1,2})\s*（?月?）?(\d{1,2})\s*（?日?）?|20\d{2}\s*[-/－]\s*(\d{1,2})\s*[-/－]\s*(\d{1,2})"
        ' 金额正则表达式：匹配可能带有小数点的金额
        Dim amountPattern As String = "(\d+(\.\d{1,2})?)"

        ' 用来保存所有匹配到的日期和金额
        Dim allDates As New List(Of DateTime)
        Dim allAmounts As New List(Of Decimal)

        ' 确保文件夹路径存在
        If Directory.Exists(folderPath) Then
            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"文件夹路径存在：{folderPath}")

            ' 获取所有txt文件
            Dim txtFiles = Directory.GetFiles(folderPath, "*.txt")
            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"找到{txtFiles.Length}个txt文件，路径：{folderPath}")

            ' 遍历每个txt文件
            For Each file As String In txtFiles
                ' 读取文件内容
                Dim fileContent As String = System.IO.File.ReadAllText(file)
                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"成功读取文件：{file}")

                ' 匹配文件中的所有日期
                Dim dateMatches = Regex.Matches(fileContent, datePattern)
                For Each match As Match In dateMatches
                    If match.Success Then
                        ' 提取并格式化日期
                        Dim year As Integer = Integer.Parse(match.Groups(1).Value)
                        Dim month As Integer = Integer.Parse(match.Groups(2).Value)
                        Dim day As Integer = Integer.Parse(match.Groups(3).Value)

                        ' 尝试创建日期对象并添加到日期列表
                        Try
                            allDates.Add(New DateTime(year, month, day))
                            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"成功提取日期：{year}-{month}-{day}")
                        Catch ex As System.Exception
                            ' 如果无法解析为日期，跳过
                            LogManager.WriteLog(LogLevel.Error, "FileProcessor.Rename", $"无法解析日期：{year}-{month}-{day}，错误信息：{ex.Message}")
                            Continue For
                        End Try
                    End If
                Next

                ' 匹配文件中的所有金额
                Dim amountMatches = Regex.Matches(fileContent, amountPattern)
                For Each match As Match In amountMatches
                    If match.Success Then
                        ' 提取金额并转换为Decimal
                        Dim amount As Decimal = Decimal.Parse(match.Groups(1).Value)
                        allAmounts.Add(amount)
                        LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"成功提取金额：{amount}")
                    End If
                Next
            Next

            ' 判断是否有匹配到日期和金额
            If allDates.Any() AndAlso allAmounts.Any() Then
                ' 获取最小的日期（最早的日期）
                Dim earliestDate As DateTime = allDates.Min()
                ' 获取最大的金额
                Dim maxAmount As Decimal = allAmounts.Max()

                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"最早的日期：{earliestDate.ToString("yyyy年MM月dd日")}")
                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.Rename", $"最大金额：{maxAmount.ToString("C2")}")

                ' 返回结果字符串
                Return $"最早的日期：{earliestDate.ToString("yyyy年MM月dd日")}" & vbCrLf &
               $"最大金额：{maxAmount.ToString("C2")}"
            Else
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.Rename", "未找到有效的日期或金额")
                Return "未找到有效的日期或金额"
            End If
        Else
            LogManager.WriteLog(LogLevel.Error, "FileProcessor.Rename", $"指定的文件夹路径不存在：{folderPath}")
            Return "指定的文件夹路径不存在"
        End If
    End Function


    ' 移动PDF文件，增强版，包含多项容错机制
    Public Function MovePdfFiles(folderPath As String) As Integer
        ' 获取目标路径
        Dim archivePath As String = My.Settings.ArchivePath

        LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"开始移动PDF文件，源文件夹: {folderPath}，目标文件夹: {archivePath}")

        ' 验证目标路径是否有效
        If String.IsNullOrEmpty(archivePath) Then
            LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", "目标路径未配置或为空")
            Return 0
        End If

        ' 如果目标路径不存在，则创建它
        Try
            If Not Directory.Exists(archivePath) Then
                Directory.CreateDirectory(archivePath)
                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"目标路径不存在，已创建：{archivePath}")
            Else
                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"目标路径已存在：{archivePath}")
            End If

            ' 验证目标路径可写
            Try
                ' 尝试创建一个临时文件来测试写入权限
                Dim testFilePath As String = Path.Combine(archivePath, "test_write.tmp")
                File.WriteAllText(testFilePath, "Test")
                File.Delete(testFilePath)
                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", "目标路径有写入权限")
            Catch ex As system.Exception
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"目标路径没有写入权限：{ex.Message}")
                Return 0
            End Try
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"创建或验证目标路径时出错：{ex.Message}")
            Return 0
        End Try

        ' 获取文件夹中的所有PDF文件
        Dim pdfFiles As String() = {}
        Try
            pdfFiles = Directory.GetFiles(folderPath, "*.pdf")
            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"在源文件夹 {folderPath} 中找到 {pdfFiles.Length} 个PDF文件")
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"获取源文件夹中的PDF文件时出错：{ex.Message}")
            Return 0
        End Try

        ' 如果没有找到PDF文件，记录并返回
        If pdfFiles.Length = 0 Then
            LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"源文件夹 {folderPath} 中没有找到PDF文件")
            Return 0
        End If

        ' 计数成功移动的文件数量
        Dim movedFilesCount As Integer = 0

        ' 遍历每个PDF文件并移动
        For Each pdfFile In pdfFiles
            ' 获取文件名
            Dim fileName As String = Path.GetFileName(pdfFile)
            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"开始处理文件：{fileName}")

            ' 确保源文件存在且可访问
            If Not File.Exists(pdfFile) Then
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"源文件不存在：{pdfFile}")
                Continue For
            End If

            ' 检查源文件是否可读
            Try
                Using fs As New FileStream(pdfFile, FileMode.Open, FileAccess.Read, FileShare.Read)
                    ' 只是测试打开，然后立即关闭
                End Using
            Catch ex As system.Exception
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"无法访问源文件：{pdfFile}，错误：{ex.Message}")
                Continue For
            End Try

            ' 检查目标文件是否已存在，若存在则修改文件名
            Dim targetPath As String = Path.Combine(archivePath, fileName)
            Dim uniqueCounter As Integer = 0

            While File.Exists(targetPath)
                uniqueCounter += 1
                Dim newFileName As String = Path.GetFileNameWithoutExtension(pdfFile) & $"_({uniqueCounter})" & Path.GetExtension(pdfFile)
                targetPath = Path.Combine(archivePath, newFileName)
                LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"目标文件已存在，文件重命名为：{Path.GetFileName(targetPath)}")
            End While

            ' 设置最大重试次数
            Const maxRetries As Integer = 3
            Dim retryCount As Integer = 0
            Dim moved As Boolean = False

            While retryCount < maxRetries AndAlso Not moved
                Try
                    ' 移动文件到目标路径
                    File.Move(pdfFile, targetPath)
                    movedFilesCount += 1
                    moved = True
                    LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"成功移动文件：{fileName} 到 {targetPath}")
                Catch ex As IOException
                    ' IO错误（文件可能被锁定）
                    retryCount += 1
                    LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"移动文件 {fileName} 时发生IO错误（尝试 {retryCount}/{maxRetries}）：{ex.Message}")

                    ' 尝试先复制后删除的方式
                    If retryCount = maxRetries - 1 Then
                        Try
                            File.Copy(pdfFile, targetPath, False)
                            File.Delete(pdfFile)
                            movedFilesCount += 1
                            moved = True
                            LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"通过复制+删除方式成功移动文件：{fileName} 到 {targetPath}")
                        Catch copyEx As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"尝试复制+删除方式移动文件 {fileName} 时出错：{copyEx.Message}")
                        End Try
                    Else
                        ' 等待一段时间再重试
                        Thread.Sleep(500)
                    End If
                Catch ex As UnauthorizedAccessException
                    ' 权限错误
                    LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"移动文件 {fileName} 时发生权限错误：{ex.Message}")

                    ' 尝试修改文件属性
                    Try
                        Dim fileInfo As New FileInfo(pdfFile)
                        fileInfo.Attributes = FileAttributes.Normal
                        retryCount += 1
                        Thread.Sleep(500)
                    Catch attrEx As system.Exception
                        LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"修改文件 {fileName} 属性时出错：{attrEx.Message}")
                        retryCount = maxRetries  ' 强制退出循环
                    End Try
                Catch ex As system.Exception
                    ' 其他错误
                    LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"移动文件 {fileName} 时发生未知错误：{ex.Message}")
                    retryCount = maxRetries  ' 强制退出循环
                End Try
            End While

            If Not moved Then
                LogManager.WriteLog(LogLevel.Error, "FileProcessor.MovePdfFiles", $"在 {maxRetries} 次尝试后仍无法移动文件：{fileName}")
            End If
        Next

        ' 返回成功移动的文件数量
        LogManager.WriteLog(LogLevel.INFO, "FileProcessor.MovePdfFiles", $"成功移动 {movedFilesCount} 个PDF文件，共 {pdfFiles.Length} 个")
        Return movedFilesCount
    End Function


    ' 标记邮件为"已完成"状态
    Public Function MarkerMail(folderPath As String, mailItem As Outlook.MailItem) As Boolean
        ' 获取文件夹中的所有PDF文件
        Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MarkerMail", $"在文件夹 {folderPath} 中找到 {pdfFiles.Length} 个PDF文件")

        ' 检查文件夹中是否有PDF文件
        If pdfFiles.Length = 0 Then
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MarkerMail", $"没有找到PDF文件，返回 False")
            Return False
        End If

        ' 调用 MovePdfFiles 函数，返回成功移动的文件数量
        Dim movedFiles As Integer = MovePdfFiles(folderPath)
        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MarkerMail", $"调用 MovePdfFiles 函数，成功移动了 {movedFiles} 个文件")

        ' 如果成功移动的文件数量与文件夹中的PDF文件数量一致且不为0
        If movedFiles = pdfFiles.Length AndAlso movedFiles > 0 Then
            ' 标记邮件为已完成
            mailItem.FlagStatus = Outlook.OlFlagStatus.olFlagComplete
            mailItem.Save()
            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MarkerMail", $"邮件 {mailItem.Subject} (ID: {mailItem.EntryID}) 标记为已完成")
            Return True
        Else
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.MarkerMail", $"文件移动失败，未能成功移动所有PDF文件。成功移动 {movedFiles} 个，文件夹中共有 {pdfFiles.Length} 个")
            Return False
        End If
    End Function

    ' 合并滴滴发票和行程单，增强版
    Private Function MergeDidiPdfFiles(folderPath As String) As Boolean
        ' 记录日志：开始合并滴滴文件
        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"开始合并滴滴文件，文件夹路径：{folderPath}")

        ' 使用统一的检查方法
        Dim fileInfo As DidiFileInfo = CheckDidiFiles(folderPath)

        ' 如果不存在滴滴发票或行程单，则退出
        If Not fileInfo.HasInvoiceFile OrElse Not fileInfo.HasTripFile Then
            If Not fileInfo.HasInvoiceFile Then
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", "未找到滴滴发票文件")
            End If
            If Not fileInfo.HasTripFile Then
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", "未找到滴滴行程单文件")
            End If
            Return False
        End If

        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"找到滴滴发票：{fileInfo.InvoiceFilePath}和行程单：{fileInfo.TripFilePath}")

        ' 验证文件的可访问性和有效性
        If Not VerifyPdfFile(fileInfo.InvoiceFilePath) Then
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"发票文件无效或无法访问：{fileInfo.InvoiceFilePath}")
            Return False
        End If

        If Not VerifyPdfFile(fileInfo.TripFilePath) Then
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"行程单文件无效或无法访问：{fileInfo.TripFilePath}")
            Return False
        End If

        Dim mergedPdf As iTextSharp.text.Document = Nothing
        Dim writer As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim invoiceReader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim tripReader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim mergedFilePath As String = ""
        Dim invoiceFileStream As FileStream = Nothing
        Dim tripFileStream As FileStream = Nothing
        Dim mergedFileStream As FileStream = Nothing

        ' 最大重试次数
        Const maxRetries As Integer = 3

        For retryCount As Integer = 0 To maxRetries - 1
            Try
                ' 从发票文件名中提取日期和金额
                Dim invoiceFileName As String = Path.GetFileNameWithoutExtension(fileInfo.InvoiceFilePath)
                Dim dateAmountMatch As Match = Regex.Match(invoiceFileName, "(\d{4}-\d{2}-\d{2})_(\d+\.\d{2})")

                ' 创建合并后的文件名
                Dim mergedFileName As String
                If dateAmountMatch.Success Then
                    ' 如果文件名中包含日期和金额，使用这些信息
                    Dim dateStr As String = dateAmountMatch.Groups(1).Value
                    Dim amountStr As String = dateAmountMatch.Groups(2).Value
                    mergedFileName = $"{dateStr}_{amountStr}_滴滴出行_合并.pdf"
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"从文件名提取信息：日期={dateStr}，金额={amountStr}")
                Else
                    ' 否则使用发票文件名作为基础
                    mergedFileName = Path.GetFileNameWithoutExtension(fileInfo.InvoiceFilePath).Replace("发票", "").Replace("电子发票", "") & "_合并.pdf"
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"未从文件名提取到日期和金额信息，使用基础文件名")
                End If

                mergedFilePath = Path.Combine(folderPath, mergedFileName)

                ' 检查目标路径是否已存在文件
                If File.Exists(mergedFilePath) Then
                    Dim uniqueCounter As Integer = 1
                    Dim newMergedFilePath As String = mergedFilePath

                    ' 如果文件已存在，添加序号直到找到可用名称
                    While File.Exists(newMergedFilePath)
                        newMergedFilePath = Path.Combine(folderPath, Path.GetFileNameWithoutExtension(mergedFilePath) & $"_({uniqueCounter})" & Path.GetExtension(mergedFilePath))
                        uniqueCounter += 1
                    End While

                    mergedFilePath = newMergedFilePath
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"调整合并文件路径以避免覆盖：{mergedFilePath}")
                End If

                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"准备创建合并文件：{mergedFilePath}")

                ' 使用iTextSharp合并PDF文件
                mergedPdf = New iTextSharp.text.Document()
                mergedFileStream = New FileStream(mergedFilePath, FileMode.Create)
                writer = New iTextSharp.text.pdf.PdfCopy(mergedPdf, mergedFileStream)
                mergedPdf.Open()
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", "开始合并PDF文件")

                ' 添加发票文件
                invoiceFileStream = New FileStream(fileInfo.InvoiceFilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                invoiceReader = New iTextSharp.text.pdf.PdfReader(invoiceFileStream)
                If invoiceReader.NumberOfPages > 0 Then
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"添加发票文件，页数：{invoiceReader.NumberOfPages}")
                    For i As Integer = 1 To invoiceReader.NumberOfPages
                        writer.AddPage(writer.GetImportedPage(invoiceReader, i))
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"已添加发票第 {i} 页")
                    Next
                Else
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", "发票文件没有页面")
                    CloseAllResources(mergedPdf, writer, invoiceReader, tripReader, mergedFileStream, invoiceFileStream, tripFileStream)
                    Return False
                End If

                ' 添加行程单文件
                tripFileStream = New FileStream(fileInfo.TripFilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                tripReader = New iTextSharp.text.pdf.PdfReader(tripFileStream)
                If tripReader.NumberOfPages > 0 Then
                    LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"添加行程单文件，页数：{tripReader.NumberOfPages}")
                    For i As Integer = 1 To tripReader.NumberOfPages
                        writer.AddPage(writer.GetImportedPage(tripReader, i))
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"已添加行程单第 {i} 页")
                    Next
                Else
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", "行程单文件没有页面")
                    CloseAllResources(mergedPdf, writer, invoiceReader, tripReader, mergedFileStream, invoiceFileStream, tripFileStream)
                    Return False
                End If

                ' 关闭所有资源
                CloseAllResources(mergedPdf, writer, invoiceReader, tripReader, mergedFileStream, invoiceFileStream, tripFileStream)
                mergedPdf = Nothing
                writer = Nothing
                invoiceReader = Nothing
                tripReader = Nothing
                mergedFileStream = Nothing
                invoiceFileStream = Nothing
                tripFileStream = Nothing

                ' 等待更长时间确保文件完全关闭
                Thread.Sleep(1000)

                ' 验证合并后的文件是否有效
                If Not VerifyPdfFile(mergedFilePath) Then
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"合并文件创建失败或无效：{mergedFilePath}")

                    ' 如果不是最后一次重试，则删除失败的文件并重试
                    If retryCount < maxRetries - 1 Then
                        Try
                            If File.Exists(mergedFilePath) Then
                                File.Delete(mergedFilePath)
                            End If
                            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"删除无效的合并文件，准备重试: {retryCount + 1}/{maxRetries}")
                            Thread.Sleep(500)
                            Continue For
                        Catch ex As system.Exception
                            LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"删除无效合并文件时出错：{ex.Message}")
                        End Try
                    End If

                    Return False
                End If

                ' 删除原始文件
                Try
                    ' 在删除原始文件前，确保合并文件存在且大小正常
                    Dim mergedFileInfo As New FileInfo(mergedFilePath)
                    If mergedFileInfo.Exists AndAlso mergedFileInfo.Length > 0 Then
                        Dim deleteSuccess As Boolean = True

                        ' 尝试删除发票文件
                        Try
                            File.Delete(fileInfo.InvoiceFilePath)
                            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"成功删除发票文件: {fileInfo.InvoiceFilePath}")
                        Catch ex As system.Exception
                            deleteSuccess = False
                            LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"删除发票文件时出错：{ex.Message}")
                        End Try

                        ' 尝试删除行程单文件
                        Try
                            File.Delete(fileInfo.TripFilePath)
                            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"成功删除行程单文件: {fileInfo.TripFilePath}")
                        Catch ex As system.Exception
                            deleteSuccess = False
                            LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"删除行程单文件时出错：{ex.Message}")
                        End Try

                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles",
                                         $"成功合并文件{If(deleteSuccess, "并删除原始文件", "但删除原始文件失败")}，合并后文件：{mergedFilePath}")
                        LogManager.WriteLog(LogLevel.INFO, "MailProcessor.MergeDidiPdfFiles", $"合并后的文件名已包含日期和金额信息，无需再次重命名")
                    Else
                        LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles",
                                          $"合并文件不存在或大小异常，不删除原始文件: {mergedFilePath}, 大小: {mergedFileInfo.Length} 字节")
                        Return False
                    End If
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"检查合并文件或删除原始文件时出错：{ex.Message}")
                    ' 即使删除失败也返回True，因为合并已经成功
                End Try

                Return True  ' 成功完成合并
            Catch ex As IOException
                ' 处理IO异常，可能是文件正在被使用
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles",
                                  $"合并文件时发生IO异常 (尝试 {retryCount + 1}/{maxRetries})：{ex.Message}")

                CloseAllResources(mergedPdf, writer, invoiceReader, tripReader, mergedFileStream, invoiceFileStream, tripFileStream)
                mergedPdf = Nothing
                writer = Nothing
                invoiceReader = Nothing
                tripReader = Nothing
                mergedFileStream = Nothing
                invoiceFileStream = Nothing
                tripFileStream = Nothing

                ' 进行垃圾回收并等待一段时间后重试
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Thread.Sleep(1000 * (retryCount + 1))  ' 随着重试次数增加等待时间
            Catch ex As system.Exception
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles",
                                  $"合并文件时发生错误 (尝试 {retryCount + 1}/{maxRetries})：{ex.Message}")
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"错误堆栈：{ex.StackTrace}")

                CloseAllResources(mergedPdf, writer, invoiceReader, tripReader, mergedFileStream, invoiceFileStream, tripFileStream)
                mergedPdf = Nothing
                writer = Nothing
                invoiceReader = Nothing
                tripReader = Nothing
                mergedFileStream = Nothing
                invoiceFileStream = Nothing
                tripFileStream = Nothing

                ' 进行垃圾回收并等待一段时间后重试
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Thread.Sleep(1000 * (retryCount + 1))  ' 随着重试次数增加等待时间
            End Try
        Next

        ' 如果所有重试都失败
        LogManager.WriteLog(LogLevel.Error, "MailProcessor.MergeDidiPdfFiles", $"在 {maxRetries} 次尝试后仍无法合并PDF文件")
        Return False
    End Function

    ' 验证PDF文件是否有效且可访问
    Private Function VerifyPdfFile(filePath As String) As Boolean
        If String.IsNullOrEmpty(filePath) OrElse Not File.Exists(filePath) Then
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.VerifyPdfFile", $"文件不存在：{filePath}")
            Return False
        End If

        ' 检查文件大小
        Try
            Dim fileInfo As New FileInfo(filePath)
            If fileInfo.Length = 0 Then
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.VerifyPdfFile", $"文件大小为0：{filePath}")
                Return False
            End If
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.VerifyPdfFile", $"检查文件大小时出错：{ex.Message}")
            Return False
        End Try

        ' 尝试打开文件进行读取
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                ' 验证文件是否为PDF格式
                Dim buffer As Byte() = New Byte(4) {}
                fs.Read(buffer, 0, 5)
                Dim fileHeader As String = System.Text.Encoding.ASCII.GetString(buffer)

                If Not fileHeader.StartsWith("%PDF") Then
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.VerifyPdfFile", $"文件不是有效的PDF格式：{filePath}")
                    Return False
                End If
            End Using

            ' 尝试用PdfReader打开（更严格的验证）
            Using reader As New iTextSharp.text.pdf.PdfReader(filePath)
                ' 检查页数
                If reader.NumberOfPages <= 0 Then
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.VerifyPdfFile", $"PDF文件没有页面：{filePath}")
                    Return False
                End If
            End Using

            LogManager.WriteLog(LogLevel.INFO, "MailProcessor.VerifyPdfFile", $"PDF文件验证通过：{filePath}")
            Return True
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.VerifyPdfFile", $"验证PDF文件时出错：{ex.Message}，文件路径：{filePath}")
            Return False
        End Try
    End Function

    ' 关闭所有资源的辅助方法
    Private Sub CloseAllResources(ByRef doc As iTextSharp.text.Document,
                                 ByRef writer As iTextSharp.text.pdf.PdfCopy,
                                 ByRef invoiceReader As iTextSharp.text.pdf.PdfReader,
                                 ByRef tripReader As iTextSharp.text.pdf.PdfReader,
                                 ByRef mergedFileStream As FileStream,
                                 ByRef invoiceFileStream As FileStream,
                                 ByRef tripFileStream As FileStream)
        Try
            ' 关闭所有资源，处理可能的异常
            If writer IsNot Nothing Then
                Try
                    writer.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭PdfCopy时出错：{ex.Message}")
                End Try
                writer = Nothing
            End If

            If doc IsNot Nothing Then
                Try
                    doc.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭Document时出错：{ex.Message}")
                End Try
                doc = Nothing
            End If

            If invoiceReader IsNot Nothing Then
                Try
                    invoiceReader.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭发票Reader时出错：{ex.Message}")
                End Try
                invoiceReader = Nothing
            End If

            If tripReader IsNot Nothing Then
                Try
                    tripReader.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭行程单Reader时出错：{ex.Message}")
                End Try
                tripReader = Nothing
            End If

            If mergedFileStream IsNot Nothing Then
                Try
                    mergedFileStream.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭合并文件流时出错：{ex.Message}")
                End Try
                mergedFileStream = Nothing
            End If

            If invoiceFileStream IsNot Nothing Then
                Try
                    invoiceFileStream.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭发票文件流时出错：{ex.Message}")
                End Try
                invoiceFileStream = Nothing
            End If

            If tripFileStream IsNot Nothing Then
                Try
                    tripFileStream.Close()
                Catch ex As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭行程单文件流时出错：{ex.Message}")
                End Try
                tripFileStream = Nothing
            End If
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseAllResources", $"关闭资源过程中发生未预期的错误：{ex.Message}")
        End Try
    End Sub

    ' 安全删除目录，包含重试机制和处理文件锁定的情况
    Private Sub SafeDeleteDirectory(directoryPath As String)
        If String.IsNullOrEmpty(directoryPath) OrElse Not Directory.Exists(directoryPath) Then
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.SafeDeleteDirectory", $"目录不存在或为空：{directoryPath}")
            Return
        End If

        ' 最大重试次数
        Const maxRetries As Integer = 3
        ' 每次重试的等待时间（毫秒）
        Const retryDelay As Integer = 500

        For retryCount As Integer = 0 To maxRetries - 1
            Try
                ' 首先尝试解锁并关闭目录中的所有文件
                CloseFilesInDirectory(directoryPath)

                ' 删除目录及其内容
                Directory.Delete(directoryPath, True)
                LogManager.WriteLog(LogLevel.INFO, "MailProcessor.SafeDeleteDirectory", $"成功删除目录：{directoryPath}")
                Return ' 成功删除，退出方法
            Catch ex As IOException
                ' 如果是IO异常（文件可能被锁定），等待一段时间后重试
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.SafeDeleteDirectory", $"删除目录时遇到IO异常，正在重试 ({retryCount + 1}/{maxRetries})：{ex.Message}")
                Thread.Sleep(retryDelay)
            Catch ex As UnauthorizedAccessException
                ' 如果是权限问题
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.SafeDeleteDirectory", $"删除目录时遇到权限问题：{ex.Message}")
                ' 尝试修改文件属性后重试
                Try
                    SetAttributesNormal(directoryPath)
                    Thread.Sleep(retryDelay)
                Catch attrEx As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.SafeDeleteDirectory", $"修改文件属性时出错：{attrEx.Message}")
                End Try
            Catch ex As system.Exception
                ' 其他类型的异常
                LogManager.WriteLog(LogLevel.Error, "MailProcessor.SafeDeleteDirectory", $"删除目录时遇到未预期的异常：{ex.Message}")
                Thread.Sleep(retryDelay)
            End Try
        Next

        ' 如果所有重试都失败，记录错误
        LogManager.WriteLog(LogLevel.Error, "MailProcessor.SafeDeleteDirectory", $"在 {maxRetries} 次尝试后仍无法删除目录：{directoryPath}")
    End Sub

    ' 关闭目录中可能被打开的文件
    Private Sub CloseFilesInDirectory(directoryPath As String)
        Try
            ' 强制进行垃圾回收，关闭任何可能被CLR持有的文件句柄
            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' 第二次垃圾回收，确保所有临时对象都被释放
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.CloseFilesInDirectory", $"强制垃圾回收过程中发生错误：{ex.Message}")
        End Try
    End Sub

    ' 将文件和目录的属性设置为normal
    Private Sub SetAttributesNormal(path As String)
        ' 处理目录本身
        Try
            Dim di As New DirectoryInfo(path)
            If (di.Attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                di.Attributes = FileAttributes.Normal
            End If
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.SetAttributesNormal", $"修改目录属性失败：{ex.Message}")
        End Try

        ' 处理所有文件
        Try
            For Each filePath As String In Directory.GetFiles(path, "*", SearchOption.AllDirectories)
                Try
                    Dim fi As New FileInfo(filePath)
                    If (fi.Attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                        fi.Attributes = FileAttributes.Normal
                    End If
                Catch fileEx As system.Exception
                    LogManager.WriteLog(LogLevel.Error, "MailProcessor.SetAttributesNormal", $"修改文件属性失败：{filePath}，错误：{fileEx.Message}")
                End Try
            Next
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.SetAttributesNormal", $"获取或处理目录文件列表时出错：{ex.Message}")
        End Try

        ' 递归处理子目录
        Try
            For Each subDirPath As String In Directory.GetDirectories(path)
                SetAttributesNormal(subDirPath)
            Next
        Catch ex As system.Exception
            LogManager.WriteLog(LogLevel.Error, "MailProcessor.SetAttributesNormal", $"获取或处理子目录列表时出错：{ex.Message}")
        End Try
    End Sub

End Class
