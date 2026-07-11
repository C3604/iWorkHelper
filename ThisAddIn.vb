Imports System.Diagnostics

Public Class ThisAddIn

    Private _startupStopwatch As Stopwatch = Nothing

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        _startupStopwatch = New System.Diagnostics.Stopwatch()
        _startupStopwatch.Start()

        Try
            ' 极轻量启动：仅初始化日志，标记已加载。
            ' 业务初始化和配置迁移全部延迟到用户操作时执行。

            StartupPerformanceTracker.StartStage("AppLogger.Initialize")
            AppLogger.Initialize(My.Settings.ArchiveFolderPath)
            StartupPerformanceTracker.LogStageDuration("AppLogger.Initialize")

            AppLogger.Info("=== iWorkHelper Startup 开始 ===")
            AppLogger.Debug("进程 ID: " & Process.GetCurrentProcess().Id)
            AppLogger.Debug("Outlook 加载项启动。")

            ' 注意：明文 Secret Key 迁移现在移到 SettingsForm 打开或归档前执行，
            ' 而不是在 Startup 中执行。这样可以避免启动时不必要的磁盘写入。
            ' 详见：ProtectedSettingsProvider.MigratePlaintextIfNeeded() 移至 SettingsForm.vb

            AppLogger.Info("iWorkHelper Startup 完成。")

        Catch ex As Exception
            ' 启动阶段任何异常都不应阻止加载项加载。
            Try
                AppLogger.Error("Startup 发生异常（加载项仍会继续运行）。", ex)
            Catch
            End Try
        Finally
            If _startupStopwatch IsNot Nothing Then
                _startupStopwatch.Stop()
                Try
                    AppLogger.Info("Startup 总耗时：" & _startupStopwatch.ElapsedMilliseconds & "ms")
                Catch
                End Try
            End If
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Dim shutdownStopwatch As New System.Diagnostics.Stopwatch()
        shutdownStopwatch.Start()

        Try
            AppLogger.Info("=== iWorkHelper Shutdown 开始 ===")
            ' Shutdown 阶段仅记日志，不执行耗时清理。
            ' 所有资源释放应在各对象的 Dispose 中处理。
            AppLogger.Info("iWorkHelper Shutdown 完成。")

        Catch ex As Exception
            ' Shutdown 异常必须吞掉，不能抛出。
            Try
                AppLogger.Error("Shutdown 发生异常（已吞掉，不影响 Outlook 关闭）。", ex)
            Catch
            End Try
        Finally
            shutdownStopwatch.Stop()
            Try
                AppLogger.Info("Shutdown 总耗时：" & shutdownStopwatch.ElapsedMilliseconds & "ms")
            Catch
            End Try
        End Try
    End Sub

End Class
