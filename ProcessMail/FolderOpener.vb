Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.IO

Public Class FolderOpener

    ' Windows API，用于激活窗口
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    ' Windows API，用于找到窗口句柄
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(className As String, windowName As String) As IntPtr
    End Function

    ' 构造函数，接收文件夹路径
    Public Sub New(folderPath As String)
        ' 记录日志：开始处理文件夹
        LogManager.WriteLog(LogLevel.INFO, "FolderOpener.New", "开始处理文件夹路径：" & folderPath)

        ' 1. 检查文件夹路径是否为空
        If String.IsNullOrWhiteSpace(folderPath) Then
            LogManager.WriteLog(LogLevel.Error, "FolderOpener.New", "文件夹路径为空！")
            Throw New ArgumentException("文件夹路径不能为空！")
        End If

        ' 2. 检查路径是否存在
        If Not IO.Directory.Exists(folderPath) Then
            ' 检查路径是否是文件而不是文件夹
            If IO.File.Exists(folderPath) Then
                LogManager.WriteLog(LogLevel.Error, "FolderOpener.New", "路径指向的是文件而非文件夹：" & folderPath)
                Throw New ArgumentException("路径指向的不是文件夹，而是一个文件！")
            Else
                LogManager.WriteLog(LogLevel.Error, "FolderOpener.New", "指定的文件夹路径不存在：" & folderPath)
                Throw New ArgumentException("指定的文件夹路径不存在！")
            End If
        End If

        ' 3. 检查传入的路径是否是一个文件夹
        If Not (New DirectoryInfo(folderPath).Attributes And FileAttributes.Directory) = FileAttributes.Directory Then
            LogManager.WriteLog(LogLevel.Error, "FolderOpener.New", "路径指向的不是文件夹：" & folderPath)
            Throw New ArgumentException("路径指向的不是文件夹！")
        End If

        ' 4. 检查是否已经打开该文件夹
        If Not IsFolderOpened(folderPath) Then
            ' 如果没有打开，则打开该文件夹
            LogManager.WriteLog(LogLevel.INFO, "FolderOpener.New", "文件夹未打开，准备打开文件夹：" & folderPath)
            OpenFolder(folderPath)
        Else
            ' 如果已打开，则激活该文件夹
            LogManager.WriteLog(LogLevel.INFO, "FolderOpener.New", "文件夹已经打开，准备激活：" & folderPath)
            ActivateFolder(folderPath)
        End If
    End Sub

    ' 检查文件夹是否已经打开
    Private Function IsFolderOpened(folderPath As String) As Boolean
        LogManager.WriteLog(LogLevel.INFO, "FolderOpener.IsFolderOpened", "检查文件夹是否已打开：" & folderPath)

        ' 获取已打开的所有窗口
        Dim processList As Process() = Process.GetProcessesByName("explorer")

        ' 遍历所有窗口
        For Each process As Process In processList
            ' 如果文件夹路径在窗口标题中，则认为已打开
            If process.MainWindowTitle.Contains(folderPath) Then
                LogManager.WriteLog(LogLevel.INFO, "FolderOpener.IsFolderOpened", "已打开文件夹窗口：" & folderPath)
                Return True
            End If
        Next

        ' 没有找到对应的窗口，返回False
        LogManager.WriteLog(LogLevel.INFO, "FolderOpener.IsFolderOpened", "未找到文件夹窗口，文件夹未打开：" & folderPath)
        Return False
    End Function

    ' 打开文件夹
    Private Sub OpenFolder(folderPath As String)
        Try
            LogManager.WriteLog(LogLevel.INFO, "FolderOpener.OpenFolder", "尝试打开文件夹：" & folderPath)
            ' 启动资源管理器打开文件夹
            Process.Start("explorer", folderPath)
            LogManager.WriteLog(LogLevel.INFO, "FolderOpener.OpenFolder", "文件夹已成功打开：" & folderPath)
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "FolderOpener.OpenFolder", "无法打开文件夹：" & folderPath & " 错误：" & ex.Message)
            Throw New InvalidOperationException("无法打开文件夹！", ex)
        End Try
    End Sub

    ' 激活已打开的文件夹窗口
    Private Sub ActivateFolder(folderPath As String)
        Try
            LogManager.WriteLog(LogLevel.INFO, "FolderOpener.ActivateFolder", "尝试激活文件夹：" & folderPath)
            ' 查找窗口句柄
            Dim hWnd As IntPtr = FindWindow(Nothing, folderPath)

            ' 如果找到了句柄，则激活窗口
            If hWnd <> IntPtr.Zero Then
                SetForegroundWindow(hWnd)
                LogManager.WriteLog(LogLevel.INFO, "FolderOpener.ActivateFolder", "文件夹已成功激活：" & folderPath)
            Else
                LogManager.WriteLog(LogLevel.Error, "FolderOpener.ActivateFolder", "未找到文件夹窗口，激活失败：" & folderPath)
                Throw New InvalidOperationException("无法找到已打开的文件夹窗口！")
            End If
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "FolderOpener.ActivateFolder", "激活文件夹失败：" & folderPath & " 错误：" & ex.Message)
            Throw New InvalidOperationException("激活文件夹失败！", ex)
        End Try
    End Sub

End Class
