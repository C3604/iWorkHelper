Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class ThisAddIn
    ' 标记是否已初始化
    Private Shared isPluginInitialized As Boolean = False
    
    ' 启动事件处理
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' 不做任何耗时操作，立即返回，完全不影响Outlook启动
        ' 将所有初始化工作推迟到用户第一次交互或使用加载项功能时
    End Sub

    ' 获取此加载项实例的单例方法
    Public Shared Function GetInstance() As ThisAddIn
        Return Globals.ThisAddIn
    End Function

    ' 公共方法，供需要时调用初始化
    Public Sub EnsureInitialized()
        If isPluginInitialized Then
            Return ' 已初始化，直接返回
        End If
        
        ' 使用最轻量的方式初始化，确保不阻塞UI线程
        Try
            ' 在工作线程上按需执行初始化
            InitializePluginCore()
            isPluginInitialized = True
            
            If My.Settings.DebugStatus Then
                ' 延迟记录日志，避免影响性能
                Task.Run(Sub()
                    LogManager.WriteLog(LogLevel.INFO, "ThisAddIn.EnsureInitialized", "插件初始化完成")
                End Sub)
            End If
        Catch ex As Exception
            If My.Settings.DebugStatus Then
                Task.Run(Sub()
                    LogManager.WriteLog(LogLevel.Error, "ThisAddIn.EnsureInitialized", $"插件初始化失败: {ex.Message}")
                End Sub)
            End If
        End Try
    End Sub

    ' 核心初始化逻辑
    Private Sub InitializePluginCore()
        ' 这里只放必要的最小初始化代码
        ' 其他功能模块可以推迟到实际使用时再加载
    End Sub

    ' 关闭时处理
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 所有清理工作放在单独的线程中执行，避免延迟Outlook关闭
        Task.Run(Sub()
            Try
                If My.Settings.DebugStatus Then
                    LogManager.WriteLog(LogLevel.INFO, "ThisAddIn.ThisAddIn_Shutdown", "插件关闭")
                End If
                
                ' 清理日志系统资源
                LogManager.Cleanup()
            Catch ex As Exception
                ' 忽略关闭时的错误
            End Try
        End Sub)
    End Sub
End Class
