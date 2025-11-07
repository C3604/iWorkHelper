Imports System.Windows.Forms

Namespace LogManager
    Public Class MessageBoxNotifier
        Implements ILogNotifier

        Public Sub Notify(level As LogLevel, message As String) Implements ILogNotifier.Notify
            Select Case level
                Case LogLevel.Warn
                    MessageBox.Show(message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Case LogLevel.ErrorLevel
                    MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Case Else
                    ' Info级别不弹窗
            End Select
        End Sub
    End Class
End Namespace