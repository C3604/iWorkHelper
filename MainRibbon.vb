Imports Microsoft.Office.Tools.Ribbon

Public Class MainRibbon

    Private Sub MainRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btn_setting_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_setting.Click
        ' 实例化并显示 SettingForm
        Dim settingForm As New SettingForm()
        settingForm.ShowDialog() ' 以模态窗口显示
    End Sub
End Class
