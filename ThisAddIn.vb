Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'MsgBox("加载项已启动")
        Dim ribbon = Globals.Ribbons.MainRibbon

        ' 在启动时校验 My.Settings 中的 tmppath 是否为有效文件夹
        Dim curTmpPath As String = String.Empty
        Try
            Dim obj = My.Settings("tmppath")
            curTmpPath = If(TryCast(obj, String), String.Empty)
        Catch
            curTmpPath = String.Empty
        End Try

        Dim isValid As Boolean = False
        If Not String.IsNullOrWhiteSpace(curTmpPath) Then
            Try
                ' 必须为绝对路径并且存在且为目录
                If System.IO.Path.IsPathRooted(curTmpPath) AndAlso System.IO.Directory.Exists(curTmpPath) Then
                    isValid = True
                End If
            Catch
                isValid = False
            End Try
        End If

        If Not isValid Then
            ' 无效时重置为 当前用户临时文件夹({temp})/iWorkhelper
            Dim fallback As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "iWorkhelper")
            Try
                ' 确保目录存在
                System.IO.Directory.CreateDirectory(fallback)
            Catch
                ' 若创建失败，仍尝试写入回退路径
            End Try

            Try
                SettingsWriter.WriteSetting("tmppath", "String", fallback)
            Catch
                ' 忽略异常以保证启动流程不中断
            End Try
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
