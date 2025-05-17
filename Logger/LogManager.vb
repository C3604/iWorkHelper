Imports System.IO
Imports iWorkHelper.My
Imports System.IO.Compression
Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Threading.Tasks

Public Enum LogLevel
    INFO
    [Error]
End Enum

Public Class LogManager
    ' 获取日志文件路径
    Private Shared ReadOnly logFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp", "iWorkHelper", "iWorkHelperLog.log")
    Private Shared ReadOnly logFolder As String = Path.GetDirectoryName(logFilePath)

    ' 最大日志文件大小，单位为字节（例如：10MB）
    Private Shared ReadOnly MaxLogSize As Long = 10 * 1024 * 1024 ' 10MB

    ' 日志队列、定时器和锁对象
    Private Shared ReadOnly logQueue As New ConcurrentQueue(Of String)
    Private Shared logTimer As Timer = Nothing
    Private Shared ReadOnly lockObject As New Object()
    Private Shared isInitialized As Boolean = False
    
    ' 控制日志刷新的间隔时间（毫秒）
    Private Shared ReadOnly LogFlushInterval As Integer = 30000 ' 30秒
    
    ' 是否禁用文件系统操作，用于应急情况
    Private Shared disableFileOperations As Boolean = False

    ' 延迟初始化日志系统 - 该方法将被惰性调用
    Private Shared Sub InitializeLogger()
        ' 双重检查锁定模式，确保线程安全
        If Not isInitialized Then
            SyncLock lockObject
                If Not isInitialized Then
                    Try
                        ' 设置定时刷新日志的定时器
                        logTimer = New Timer(AddressOf FlushLogQueue, Nothing, LogFlushInterval, LogFlushInterval)
                        
                        ' 创建日志文件夹（如果不存在）
                        If Not Directory.Exists(logFolder) Then
                            Directory.CreateDirectory(logFolder)
                        End If
                        
                        isInitialized = True
                    Catch ex As Exception
                        ' 初始化失败时禁用文件操作
                        disableFileOperations = True
                        Console.WriteLine("日志系统初始化失败: " & ex.Message)
                    End Try
                End If
            End SyncLock
        End If
    End Sub

    ' 刷新日志队列，将日志写入文件
    Private Shared Sub FlushLogQueue(state As Object)
        If logQueue.IsEmpty OrElse disableFileOperations Then
            Return
        End If

        Try
            ' 批量处理日志队列中的所有条目
            Dim logEntries As New List(Of String)(100) ' 预分配容量
            Dim entry As String = Nothing

            ' 从队列中取出最多100条日志条目，避免一次处理太多
            For i As Integer = 0 To 99
                If Not logQueue.TryDequeue(entry) Then
                    Exit For
                End If
                
                If entry IsNot Nothing Then
                    logEntries.Add(entry)
                End If
            Next

            ' 如果有日志条目，则批量写入文件
            If logEntries.Count > 0 Then
                ' 检查并执行日志滚动
                CheckLogFileSize()
                
                ' 写入日志文件，使用追加模式
                Using writer As New StreamWriter(logFilePath, True)
                    For Each logEntry In logEntries
                        writer.WriteLine(logEntry)
                    Next
                End Using
            End If
        Catch ex As Exception
            ' 如果写日志过程中发生错误，禁用文件操作以避免频繁出错
            disableFileOperations = True
            Console.WriteLine("写入日志失败，已禁用文件操作: " & ex.Message)
        End Try
    End Sub

    ' 检查日志文件大小并在需要时执行滚动
    Private Shared Sub CheckLogFileSize()
        If disableFileOperations Then
            Return
        End If
        
        Try
            ' 确保日志文件夹存在
            If Not Directory.Exists(logFolder) Then
                Directory.CreateDirectory(logFolder)
                Return
            End If

            ' 如果日志文件不存在，则创建一个空的日志文件
            If Not File.Exists(logFilePath) Then
                Using fs As FileStream = File.Create(logFilePath)
                    ' 空文件创建成功后关闭流
                End Using
                Return
            End If

            ' 检查文件大小，如果超过限制则执行滚动
            Dim fileInfo As New FileInfo(logFilePath)
            If fileInfo.Length > MaxLogSize Then
                ' 执行日志滚动
                RollOverLogFile()
            End If
        Catch ex As Exception
            disableFileOperations = True
            Console.WriteLine("检查日志文件大小失败: " & ex.Message)
        End Try
    End Sub

    ' 执行日志文件滚动操作
    Private Shared Sub RollOverLogFile()
        If disableFileOperations Then
            Return
        End If
        
        Try
            ' 获取当前时间戳，作为文件名的一部分
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim archiveLogPath As String = Path.Combine(logFolder, "iWorkHelperLog_" & timestamp & ".log")

            ' 将当前日志文件重命名为带时间戳的文件
            File.Move(logFilePath, archiveLogPath)

            ' 启动新线程进行压缩，避免阻塞主线程
            Task.Run(Sub()
                Try
                    ' 使用 GZip 压缩日志文件，生成一个以 .gz 结尾的文件
                    Using fileStream As New FileStream(archiveLogPath, FileMode.Open)
                        Using gzipStream As New GZipStream(New FileStream(archiveLogPath & ".gz", FileMode.Create), CompressionMode.Compress)
                            fileStream.CopyTo(gzipStream)
                        End Using
                    End Using
                    ' 删除原始日志文件（压缩后的文件已保存为 .gz）
                    File.Delete(archiveLogPath)
                Catch ex As Exception
                    ' 压缩失败不影响主要功能
                    Console.WriteLine("日志压缩失败: " & ex.Message)
                End Try
            End Sub)
        Catch ex As Exception
            disableFileOperations = True
            Console.WriteLine("日志滚动失败: " & ex.Message)
        End Try
    End Sub

    ' 写日志 - 公共接口 (超轻量级实现)
    Public Shared Sub WriteLog(logType As LogLevel, currentFunction As String, message As String)
        ' 快速检查: 如果调试模式关闭，立即返回
        If My.Settings.DebugStatus = False Then
            Return
        End If

        ' 获取当前时间
        Dim currentDateTime As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        ' 构造日志条目
        Dim logEntry As String = String.Format("[{0}] {1} {2} {3}", logType.ToString(), currentDateTime, currentFunction, message)

        ' 将日志条目添加到队列（这个操作非常快）
        logQueue.Enqueue(logEntry)
        
        ' 延迟初始化日志系统
        If Not isInitialized Then
            Task.Run(AddressOf InitializeLogger)
        End If
    End Sub

    ' 在应用程序退出时清理资源
    Public Shared Sub Cleanup()
        SyncLock lockObject
            If logTimer IsNot Nothing Then
                logTimer.Dispose()
                logTimer = Nothing
            End If

            ' 只有在已初始化且未禁用文件操作时才刷新日志
            If isInitialized AndAlso Not disableFileOperations Then
                FlushLogQueue(Nothing)
            End If
            
            isInitialized = False
        End SyncLock
    End Sub
End Class
