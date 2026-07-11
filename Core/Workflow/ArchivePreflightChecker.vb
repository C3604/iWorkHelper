Imports System.IO

''' <summary>
''' 归档前预检查：在真正开始批量处理前，一次性检查配置与运行环境，避免中途才失败。
''' 设计为可离线测试（入参为配置值，不依赖 Outlook）；邮件选择数由调用方（Ribbon）传入。
'''
''' 注意：本类【不再】管理“归档任务是否正在运行”。运行锁由 <see cref="ArchiveRunGuard"/> 统一负责，
''' 预检查只检查配置、目录、选中邮件、权限等静态条件，避免与外层运行锁产生“自我阻断”。
''' （历史缺陷：外层先 TryBeginRun 置位，预检查又检查 IsRunning，导致每次点击都误报“已有任务正在运行”。）
''' </summary>
Public Class ArchivePreflightChecker

    ''' <summary>
    ''' 执行预检查。
    ''' </summary>
    ''' <param name="selectedCount">Outlook 当前选中项数量（由 Ribbon 传入；-1 表示未知/不检查）。</param>
    ''' <param name="archiveFolder">归档目录。</param>
    ''' <param name="ocrOptions">OCR 配置。</param>
    ''' <param name="unifiedTemplate">统一命名模板（可空 → 用默认）。</param>
    Public Function Check(selectedCount As Integer, archiveFolder As String, ocrOptions As BaiduOcrOptions, unifiedTemplate As String) As ArchivePreflightResult
        Dim r As New ArchivePreflightResult()

        ' 说明：明文 Secret Key 的 DPAPI 迁移（依赖 My.Settings）已移至 MainRibbon 在预检查之前执行，
        ' 使本预检查保持“纯配置值入参、不依赖 Outlook / My.Settings”，从而可离线单元测试。

        ' 说明：运行状态检查已移除。是否“已有归档任务正在运行”由 ArchiveRunGuard 在预检查之前判断，
        ' 预检查不再检查运行状态，避免与外层运行锁自我阻断。

        ' 选中邮件
        If selectedCount = 0 Then
            r.AddCode(AppErrorCode.NoMailSelected, blocking:=True)
        End If

        ' 归档目录
        CheckArchiveFolder(archiveFolder, r)

        ' 临时目录 / 日志目录
        Dim tempDir As String = Nothing
        Try
            tempDir = PathHelper.GetTempWorkDirectory()
        Catch
        End Try
        If String.IsNullOrEmpty(tempDir) OrElse Not IsDirWritable(tempDir) Then
            r.AddCode(AppErrorCode.TempDirNotWritable, blocking:=True)
        End If

        Dim logDir As String = Nothing
        Try
            logDir = PathHelper.GetLogDirectory(archiveFolder)
        Catch
        End Try
        If String.IsNullOrEmpty(logDir) OrElse Not IsDirWritable(logDir) Then
            r.AddCode(AppErrorCode.LogDirNotWritable, blocking:=False) ' 日志问题不阻断
        End If

        ' 命名模板
        If String.IsNullOrWhiteSpace(unifiedTemplate) Then
            r.AddCode(AppErrorCode.NamingTemplateEmpty, blocking:=False)
        End If

        ' OCR 配置（仅提示，不阻断——本地可能已足够）
        If ocrOptions IsNot Nothing AndAlso ocrOptions.Enabled AndAlso Not ocrOptions.IsConfigured() Then
            r.AddCode(AppErrorCode.OcrConfigMissing, blocking:=False)
        End If

        ' 汇总 + 逐条明细日志（避免以后只看到“问题数=1”而无法定位是哪一项）。
        AppLogger.Info("归档预检查完成：问题数=" & r.Issues.Count & "，阻断=" & r.HasBlocking)
        For Each issue As ArchivePreflightIssue In r.Issues
            AppLogger.Info(String.Format("  预检查{0}：Code={1}, Severity={2}, Blocking={3}, Message={4}",
                                         If(issue.IsBlocking, "[阻断]", "[提示]"),
                                         issue.Code, issue.Severity, issue.IsBlocking, issue.UserMessage))
        Next
        Return r
    End Function

    Private Sub CheckArchiveFolder(archiveFolder As String, r As ArchivePreflightResult)
        If String.IsNullOrWhiteSpace(archiveFolder) Then
            r.AddCode(AppErrorCode.ArchiveFolderNotConfigured, blocking:=True)
            Return
        End If

        ' 路径格式校验
        Try
            Dim full As String = Path.GetFullPath(archiveFolder)
        Catch
            r.AddCode(AppErrorCode.ArchiveFolderPathInvalid, blocking:=True)
            Return
        End Try

        ' 不存在 → 尝试创建（相当于“自动创建”）
        If Not SafeDirExists(archiveFolder) Then
            If Not PathHelper.EnsureDirectory(archiveFolder) Then
                r.AddCode(AppErrorCode.ArchiveFolderMissing, blocking:=True)
                Return
            End If
        End If

        ' 写权限
        If Not IsDirWritable(archiveFolder) Then
            r.AddCode(AppErrorCode.ArchiveFolderNotWritable, blocking:=True)
        End If
    End Sub

    Private Function SafeDirExists(dir As String) As Boolean
        Try
            Return Directory.Exists(dir)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>写权限检测：尝试创建并删除一个临时文件。</summary>
    Private Function IsDirWritable(dir As String) As Boolean
        Try
            If Not Directory.Exists(dir) Then Return False
            Dim probe As String = Path.Combine(dir, ".iwh_write_test_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ".tmp")
            File.WriteAllText(probe, "x")
            File.Delete(probe)
            Return True
        Catch
            Return False
        End Try
    End Function

End Class
