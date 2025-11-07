Imports System
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Collections.Concurrent

Namespace LogManager
    ' 关键路径：日志核心类，负责队列缓冲、异步写入、滚动分割与通知
    Public Class Logger
        Implements IDisposable

        Private ReadOnly _queue As New ConcurrentQueue(Of String)()
        Private ReadOnly _signal As New AutoResetEvent(False)
        Private ReadOnly _cts As New CancellationTokenSource()
        Private ReadOnly _notifier As ILogNotifier
        Private ReadOnly _flushIntervalMs As Integer
        Private ReadOnly _batchSize As Integer
        Private ReadOnly _maxFileSizeBytes As Long

        Private ReadOnly _utf8NoBom As New UTF8Encoding(False)
        Private ReadOnly _syncRoot As New Object()

        Private _basePath As String
        Private _logDir As String
        Private _logFile As String
        Private ReadOnly _fallbackDir As String

        Private _started As Boolean = False

        Public Sub New(Optional basePath As String = Nothing,
                       Optional notifier As ILogNotifier = Nothing,
                       Optional flushIntervalMs As Integer = 300,
                       Optional batchSize As Integer = 64,
                       Optional maxFileSizeBytes As Long = 10L * 1024 * 1024)

            Dim configuredPath As String = Nothing
            Try
                Dim hasSetting As Boolean = (My.Settings.Properties("tmppath") IsNot Nothing)
                If hasSetting Then
                    Dim obj = My.Settings("tmppath")
                    If obj IsNot Nothing Then
                        configuredPath = TryCast(obj, String)
                    End If
                End If
            Catch
                configuredPath = Nothing
            End Try

            _basePath = If(Not String.IsNullOrWhiteSpace(basePath), basePath,
                           If(Not String.IsNullOrWhiteSpace(configuredPath), configuredPath, Path.GetTempPath()))
            _logDir = Path.Combine(_basePath, "log")
            _logFile = Path.Combine(_logDir, "iWorkhelper.log")
            _fallbackDir = Path.Combine(Path.GetTempPath(), "iWorkHelper", "log")

            _notifier = If(notifier, New MessageBoxNotifier())
            _flushIntervalMs = Math.Max(100, flushIntervalMs)
            _batchSize = Math.Max(1, batchSize)
            _maxFileSizeBytes = Math.Max(1024 * 1024, maxFileSizeBytes)
        End Sub

        ' 轻量刷新：仅在 tmppath 发生变化时更新内部路径，不频繁读取设置（关键路径）
        Public Sub RefreshPathIfChanged(Optional basePathOverride As String = Nothing)
            Try
                Dim configuredPath As String = Nothing
                If String.IsNullOrWhiteSpace(basePathOverride) Then
                    Try
                        Dim obj = My.Settings("tmppath")
                        If obj IsNot Nothing Then
                            configuredPath = TryCast(obj, String)
                        End If
                    Catch
                        configuredPath = Nothing
                    End Try
                Else
                    configuredPath = basePathOverride
                End If

                Dim candidate As String = If(Not String.IsNullOrWhiteSpace(configuredPath), configuredPath, _basePath)
                If Not String.IsNullOrWhiteSpace(candidate) AndAlso String.Compare(candidate, _basePath, True) <> 0 Then
                    SyncLock _syncRoot
                        _basePath = candidate
                        _logDir = Path.Combine(_basePath, "log")
                        _logFile = Path.Combine(_logDir, "iWorkhelper.log")
                    End SyncLock
                End If
            Catch
                ' 静默处理刷新失败，避免影响主流程
            End Try
        End Sub

        Public Sub Start()
            If _started Then Return
            _started = True
            Task.Run(AddressOf WorkerLoop, _cts.Token)
        End Sub

        Public Sub [Stop]()
            If Not _started Then Return
            _cts.Cancel()
            _signal.Set()
            _started = False
        End Sub

        Public Sub LogInfo(message As String)
            Enqueue(LogLevel.Info, message)
        End Sub

        Public Sub LogWarn(message As String)
            ' 同步弹窗
            _notifier.Notify(LogLevel.Warn, message)
            Enqueue(LogLevel.Warn, message)
        End Sub

        Public Sub LogError(message As String)
            ' 同步弹窗
            _notifier.Notify(LogLevel.ErrorLevel, message)
            Enqueue(LogLevel.ErrorLevel, message)
        End Sub

        Private Sub Enqueue(level As LogLevel, message As String)
            Dim line As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{LevelToText(level)}] {message}"
            _queue.Enqueue(line)
            _signal.Set()
        End Sub

        Private Function LevelToText(level As LogLevel) As String
            Select Case level
                Case LogLevel.Info
                    Return "Info"
                Case LogLevel.Warn
                    Return "Warn"
                Case LogLevel.ErrorLevel
                    Return "Error"
                Case Else
                    Return "Info"
            End Select
        End Function

        Private Sub WorkerLoop()
            Dim token = _cts.Token
            Do While Not token.IsCancellationRequested
                Try
                    ' 等待信号或定时器
                    _signal.WaitOne(_flushIntervalMs)

                    ' 批量提取
                    Dim batch As New List(Of String)(_batchSize)
                    Dim item As String = Nothing
                    While batch.Count < _batchSize AndAlso _queue.TryDequeue(item)
                        batch.Add(item)
                    End While

                    If batch.Count = 0 Then
                        Continue Do
                    End If

                    ' 写入文件（线程安全）
                    WriteBatch(batch)
                Catch ex As Exception
                    ' 自身异常吞掉，避免影响主线程；尝试写入备用文件
                    Try
                        SafeWriteFallback($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [Error] Logger内部错误: {ex.Message}")
                    Catch
                        ' 备用也失败则忽略，避免死循环
                    End Try
                End Try
            Loop
        End Sub

        Private Sub EnsureDir(path As String)
            If Not Directory.Exists(path) Then
                Directory.CreateDirectory(path)
            End If
        End Sub

        Private Sub WriteBatch(lines As List(Of String))
            SyncLock _syncRoot
                Try
                    EnsureDir(_logDir)

                    ' 滚动分割：超过阈值则重命名现有文件
                    If File.Exists(_logFile) Then
                        Dim fi As New FileInfo(_logFile)
                        If fi.Length >= _maxFileSizeBytes Then
                            Dim rolledName = Path.Combine(_logDir, $"iWorkhelper_{DateTime.Now:yyyyMMdd_HHmmss}.log")
                            File.Move(_logFile, rolledName)
                        End If
                    End If

                    Using fs As New FileStream(_logFile, FileMode.Append, FileAccess.Write, FileShare.Read)
                        Using sw As New StreamWriter(fs, _utf8NoBom)
                            For Each line In lines
                                sw.WriteLine(line)
                            Next
                            sw.Flush()
                        End Using
                    End Using
                Catch ioex As IOException
                    ' IO异常：尝试写入备用目录
                    SafeWriteFallback($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [Error] IO异常写入失败: {ioex.Message}")
                Catch ex As Exception
                    ' 其它异常：同样写入备用目录
                    SafeWriteFallback($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [Error] 写入失败: {ex.Message}")
                End Try
            End SyncLock
        End Sub

        Private Sub SafeWriteFallback(line As String)
            Try
                EnsureDir(_fallbackDir)
                Dim fbFile = Path.Combine(_fallbackDir, "iWorkhelper_fallback.log")
                Using fs As New FileStream(fbFile, FileMode.Append, FileAccess.Write, FileShare.Read)
                    Using sw As New StreamWriter(fs, _utf8NoBom)
                        sw.WriteLine(line)
                        sw.Flush()
                    End Using
                End Using
            Catch
                ' 备用也失败则静默
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                [Stop]()
                _signal.Dispose()
                _cts.Dispose()
            Catch
            End Try
        End Sub
    End Class
End Namespace