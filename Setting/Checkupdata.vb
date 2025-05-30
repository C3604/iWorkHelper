Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks

Public Class Checkupdata
    ' 更新信息类
    Public Class UpdateInfo
        Public Property LatestVersion As String
        Public Property LastUpdated As String
        Public Property ChangelogUrl As String
        Public Property UpdateUrl As String
        Public Property ForceUpdate As Boolean
        Public Property MinForceUpdateVersion As String
    End Class

    ' 更新检查结果类
    Public Class UpdateCheckResult
        Public Property HasUpdate As Boolean
        Public Property IsForceUpdate As Boolean
        Public Property UpdateInfo As UpdateInfo
        Public Property CurrentVersion As String
        Public Property ErrorMessage As String
    End Class

    ' 检查更新的入口函数
    Public Async Function CheckForUpdates(currentVersion As String) As Task(Of UpdateCheckResult)
        ' 记录检查更新开始
        LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", "开始检查更新，当前版本: " & currentVersion)
        
        Dim result As New UpdateCheckResult With {
            .HasUpdate = False,
            .IsForceUpdate = False,
            .CurrentVersion = currentVersion
        }

        Try
            ' 从网络获取更新信息
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", "正在获取更新信息...")
            Dim updateInfo = Await GetUpdateInfoAsync()
            result.UpdateInfo = updateInfo
            
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", 
                "成功获取更新信息，最新版本: " & updateInfo.LatestVersion)

            ' 检查是否有新版本
            Dim compareResult = CompareVersions(updateInfo.LatestVersion, currentVersion)
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", 
                "版本比较结果: " & compareResult & " (>0表示有新版本)")
                
            If compareResult > 0 Then
                result.HasUpdate = True
                LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", 
                    "发现新版本: " & updateInfo.LatestVersion)
                
                ' 检查是否需要强制更新
                If updateInfo.ForceUpdate AndAlso CompareVersions(updateInfo.MinForceUpdateVersion, currentVersion) > 0 Then
                    result.IsForceUpdate = True
                    LogManager.WriteLog(LogLevel.Error, "Checkupdata.CheckForUpdates", 
                        "需要强制更新，当前版本低于最低要求版本: " & updateInfo.MinForceUpdateVersion)
                End If
            Else
                LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", "当前已是最新版本")
            End If

        Catch ex As Exception
            result.ErrorMessage = "检查更新时发生错误: " & ex.Message
            LogManager.WriteLog(LogLevel.Error, "Checkupdata.CheckForUpdates", result.ErrorMessage)
        End Try
        
        LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CheckForUpdates", 
            "检查更新完成，HasUpdate: " & result.HasUpdate & ", IsForceUpdate: " & result.IsForceUpdate)

        Return result
    End Function

    ' 从服务器获取更新信息
    Private Async Function GetUpdateInfoAsync() As Task(Of UpdateInfo)
        Dim url As String = "https://gitee.com/Xoin-Yang/master-updater/raw/master/iWorkHelper/updatacheck.json"
        
        LogManager.WriteLog(LogLevel.INFO, "Checkupdata.GetUpdateInfoAsync", "开始从服务器获取更新信息: " & url)
        
        Try
            Using client As New WebClient()
                client.Encoding = Encoding.UTF8
                LogManager.WriteLog(LogLevel.INFO, "Checkupdata.GetUpdateInfoAsync", "开始下载更新信息...")
                Dim jsonData As String = Await client.DownloadStringTaskAsync(url)
                
                LogManager.WriteLog(LogLevel.INFO, "Checkupdata.GetUpdateInfoAsync", 
                    "成功下载更新信息，数据长度: " & jsonData.Length & " 字符")
                
                LogManager.WriteLog(LogLevel.INFO, "Checkupdata.GetUpdateInfoAsync", "开始解析JSON数据...")
                Return ParseJson(jsonData)
            End Using
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "Checkupdata.GetUpdateInfoAsync", 
                "获取更新信息失败: " & ex.Message)
            Throw
        End Try
    End Function
    
    ' 手动解析JSON字符串
    Private Function ParseJson(jsonString As String) As UpdateInfo
        LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "开始解析JSON数据")
        
        Dim info As New UpdateInfo()
        
        Try
            ' 清理JSON字符串，去除多余的空格
            jsonString = jsonString.Trim()
            
            ' 确保是JSON对象
            If Not (jsonString.StartsWith("{") AndAlso jsonString.EndsWith("}")) Then
                Dim errorMsg As String = "无效的JSON格式"
                LogManager.WriteLog(LogLevel.Error, "Checkupdata.ParseJson", errorMsg)
                Throw New Exception(errorMsg)
            End If
            
            ' 去除首尾的花括号
            jsonString = jsonString.Substring(1, jsonString.Length - 2)
            
            ' 解析每个字段
            Dim fields() As String = SplitJsonFields(jsonString)
            
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "JSON字段数量: " & fields.Length)
            
            For Each field In fields
                Dim parts() As String = field.Split(New Char() {":"c}, 2)
                If parts.Length = 2 Then
                    Dim name As String = parts(0).Trim().Trim(""""c)
                    Dim value As String = parts(1).Trim()
                    
                    Select Case name.ToLower()
                        Case "latestversion"
                            info.LatestVersion = value.Trim(""""c)
                            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "解析到字段: latestVersion = " & info.LatestVersion)
                        Case "lastupdated"
                            info.LastUpdated = value.Trim(""""c)
                            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "解析到字段: lastUpdated = " & info.LastUpdated)
                        Case "changelogurl"
                            info.ChangelogUrl = value.Trim(""""c)
                            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "解析到字段: changelogUrl = " & info.ChangelogUrl)
                        Case "updateurl"
                            info.UpdateUrl = value.Trim(""""c)
                            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "解析到字段: updateUrl = " & info.UpdateUrl)
                        Case "forceupdate"
                            info.ForceUpdate = value.ToLower() = "true"
                            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "解析到字段: forceUpdate = " & info.ForceUpdate)
                        Case "minforceupdateversion"
                            info.MinForceUpdateVersion = value.Trim(""""c)
                            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "解析到字段: minForceUpdateVersion = " & info.MinForceUpdateVersion)
                    End Select
                End If
            Next
            
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.ParseJson", "JSON解析完成")
            Return info
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "Checkupdata.ParseJson", "解析JSON数据失败: " & ex.Message)
            Throw
        End Try
    End Function
    
    ' 辅助函数：拆分JSON字段
    Private Function SplitJsonFields(jsonContent As String) As String()
        Dim result As New List(Of String)()
        Dim currentField As New StringBuilder()
        Dim inQuotes As Boolean = False
        Dim bracketCount As Integer = 0
        
        For Each c As Char In jsonContent
            ' 处理引号
            If c = """"c Then
                inQuotes = Not inQuotes
            End If
            
            ' 处理嵌套对象
            If Not inQuotes Then
                If c = "{"c Then bracketCount += 1
                If c = "}"c Then bracketCount -= 1
            End If
            
            ' 处理字段分隔符
            If c = ","c AndAlso Not inQuotes AndAlso bracketCount = 0 Then
                result.Add(currentField.ToString().Trim())
                currentField.Clear()
            Else
                currentField.Append(c)
            End If
        Next
        
        ' 添加最后一个字段
        If currentField.Length > 0 Then
            result.Add(currentField.ToString().Trim())
        End If
        
        Return result.ToArray()
    End Function

    ' 比较版本号，返回值 > 0 表示 version1 比 version2 更新
    Private Function CompareVersions(version1 As String, version2 As String) As Integer
        Try
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CompareVersions", 
                "比较版本号: " & version1 & " vs " & version2)
                
            Dim v1Parts = version1.Split("."c).Select(Function(p) Integer.Parse(p)).ToArray()
            Dim v2Parts = version2.Split("."c).Select(Function(p) Integer.Parse(p)).ToArray()
            
            Dim length As Integer = Math.Max(v1Parts.Length, v2Parts.Length)
            
            For i As Integer = 0 To length - 1
                Dim v1Part As Integer = If(i < v1Parts.Length, v1Parts(i), 0)
                Dim v2Part As Integer = If(i < v2Parts.Length, v2Parts(i), 0)
                
                If v1Part > v2Part Then
                    LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CompareVersions", "版本1更新")
                    Return 1
                ElseIf v1Part < v2Part Then
                    LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CompareVersions", "版本2更新")
                    Return -1
                End If
            Next
            
            LogManager.WriteLog(LogLevel.INFO, "Checkupdata.CompareVersions", "版本相同")
            Return 0
        Catch ex As Exception
            LogManager.WriteLog(LogLevel.Error, "Checkupdata.CompareVersions", "比较版本号时出错: " & ex.Message)
            Throw
        End Try
    End Function
End Class
