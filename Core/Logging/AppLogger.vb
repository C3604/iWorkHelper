Imports System.IO
Imports System.Text

''' <summary>
''' 日志级别。
''' </summary>
Public Enum LogLevel
    Debug_ = 0
    Info = 1
    Warn = 2
    [Error] = 3
End Enum

''' <summary>
''' 轻量级本地文件日志。要求：
'''  - 所有核心流程可调用，写入本地文本文件（默认 %AppData%\oWorkHelper\logs\yyyy-MM-dd.log）；
'''  - 任何写日志失败都不得导致主流程崩溃（内部全部吞掉异常）。
''' 非高并发场景，使用简单锁保证线程安全。
''' </summary>
Public Module AppLogger

    Private ReadOnly SyncRoot As New Object()
    Private _logDirectory As String = Nothing
    Private _initialized As Boolean = False

    ''' <summary>
    ''' 初始化日志目录。可传入归档目录，优先在其下建 logs 子目录。
    ''' 允许重复调用；失败时回退到 AppData。
    ''' </summary>
    Public Sub Initialize(Optional archiveFolder As String = Nothing)
        SyncLock SyncRoot
            Try
                _logDirectory = PathHelper.GetLogDirectory(archiveFolder)
            Catch
                _logDirectory = Nothing
            End Try
            _initialized = True
        End SyncLock
    End Sub

    ''' <summary>当前日志文件完整路径（供 UI 提示定位）。目录不可用时返回目录名或占位。</summary>
    Public Function CurrentLogPath() As String
        Try
            SyncLock SyncRoot
                If Not _initialized Then Initialize()
                If String.IsNullOrEmpty(_logDirectory) Then Return "(日志目录不可用)"
                Return Path.Combine(_logDirectory, DateTime.Now.ToString("yyyy-MM-dd") & ".log")
            End SyncLock
        Catch
            Return "(日志目录不可用)"
        End Try
    End Function

    Public Sub Info(message As String)
        Write(LogLevel.Info, message, Nothing)
    End Sub

    Public Sub Warn(message As String)
        Write(LogLevel.Warn, message, Nothing)
    End Sub

    Public Sub [Error](message As String, Optional ex As Exception = Nothing)
        Write(LogLevel.Error, message, ex)
    End Sub

    Public Sub Debug(message As String)
        Write(LogLevel.Debug_, message, Nothing)
    End Sub

    Private Sub Write(level As LogLevel, message As String, ex As Exception)
        Try
            SyncLock SyncRoot
                If Not _initialized Then
                    Initialize()
                End If

                Dim dir As String = _logDirectory
                If String.IsNullOrEmpty(dir) Then
                    ' 目录不可用时静默放弃，绝不抛出。
                    Return
                End If

                Dim fileName As String = DateTime.Now.ToString("yyyy-MM-dd") & ".log"
                Dim fullPath As String = Path.Combine(dir, fileName)

                Dim sb As New StringBuilder()
                sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                sb.Append(" [")
                sb.Append(LevelText(level))
                sb.Append("] ")
                sb.Append(If(message, String.Empty))
                If ex IsNot Nothing Then
                    sb.AppendLine()
                    sb.Append(ExceptionFormatter.ToLogText(ex))
                End If
                sb.AppendLine()

                File.AppendAllText(fullPath, sb.ToString(), Encoding.UTF8)
            End SyncLock
        Catch
            ' 日志绝不影响主流程。
        End Try
    End Sub

    Private Function LevelText(level As LogLevel) As String
        Select Case level
            Case LogLevel.Debug_
                Return "DEBUG"
            Case LogLevel.Info
                Return "INFO "
            Case LogLevel.Warn
                Return "WARN "
            Case LogLevel.Error
                Return "ERROR"
            Case Else
                Return "INFO "
        End Select
    End Function

End Module
