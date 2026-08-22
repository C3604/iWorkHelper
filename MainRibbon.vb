Imports Microsoft.Office.Tools.Ribbon

Public Class MainRibbon

    Private Sub MainRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ButtonArchive_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonArchive.Click
        Dim parseMode = If(My.Settings.ParseMode, String.Empty)

        If String.Equals(parseMode, "Online", StringComparison.OrdinalIgnoreCase) Then
            MsgBox("在线模式")
        Else
            MsgBox("本地模式")
        End If
    End Sub

    Private Sub ButtonSettings_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSettings.Click
        Using form As New SettingsForm()
            form.ShowDialog()
        End Using
    End Sub

End Class
