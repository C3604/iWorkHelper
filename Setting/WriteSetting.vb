Imports System.Windows.Forms

Public Class WriteSetting
    ' 存储设置的方法
    Public Shared Function SaveSetting(settingName As String, settingType As String, settingValue As String) As Integer
        ' 开始记录日志
        LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"开始保存设置：{settingName}, 类型：{settingType}, 值：{settingValue}")

        Try
            ' 判断类型并进行相应的处理
            Select Case settingType.ToLower()
                Case "string"
                    ' 将字符串值存入 My.Settings
                    My.Settings(settingName) = settingValue
                    LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"设置项 {settingName} 已更新为字符串值：{settingValue}")
                Case "boolean"
                    ' 将布尔值存入 My.Settings（假设传入的是"True"或"False"字符串）
                    If settingValue.ToLower() = "true" Then
                        My.Settings(settingName) = True
                        LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"设置项 {settingName} 已更新为布尔值：True")
                    ElseIf settingValue.ToLower() = "false" Then
                        My.Settings(settingName) = False
                        LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"设置项 {settingName} 已更新为布尔值：False")
                    Else
                        ' 如果传入值不是True或False，则返回失败
                        LogManager.WriteLog(LogLevel.Error, "SettingsManager.SaveSetting", $"设置项 {settingName} 的布尔值无效：{settingValue}")
                        Return 0
                    End If
                Case "integer"
                    ' 将整数值存入 My.Settings（假设传入的是数字字符串）
                    Dim intValue As Integer
                    If Integer.TryParse(settingValue, intValue) Then
                        My.Settings(settingName) = intValue
                        LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"设置项 {settingName} 已更新为整数值：{intValue}")
                    Else
                        ' 如果不能转换为整数，则返回失败
                        LogManager.WriteLog(LogLevel.Error, "SettingsManager.SaveSetting", $"设置项 {settingName} 的整数值无效：{settingValue}")
                        Return 0
                    End If
                Case "datetime"
                    ' 将日期时间值存入 My.Settings（假设传入的是日期格式的字符串）
                    Dim dateValue As DateTime
                    If DateTime.TryParse(settingValue, dateValue) Then
                        My.Settings(settingName) = dateValue
                        LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"设置项 {settingName} 已更新为日期时间值：{dateValue}")
                    Else
                        ' 如果不能转换为日期时间，则返回失败
                        LogManager.WriteLog(LogLevel.Error, "SettingsManager.SaveSetting", $"设置项 {settingName} 的日期时间值无效：{settingValue}")
                        Return 0
                    End If
                Case Else
                    ' 如果不支持的类型，则返回失败
                    LogManager.WriteLog(LogLevel.Error, "SettingsManager.SaveSetting", $"不支持的设置类型：{settingType}")
                    Return 0
            End Select

            ' 如果成功设置了值，保存设置并返回成功标识
            My.Settings.Save()
            LogManager.WriteLog(LogLevel.INFO, "SettingsManager.SaveSetting", $"设置项 {settingName} 已成功保存。")
            Return 1

        Catch ex As Exception
            ' 捕捉任何异常，返回失败并打印错误信息
            LogManager.WriteLog(LogLevel.Error, "SettingsManager.SaveSetting", $"保存设置时发生错误: {ex.Message}")
            MessageBox.Show("错误信息: " & ex.Message, "保存设置失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
    End Function

End Class
