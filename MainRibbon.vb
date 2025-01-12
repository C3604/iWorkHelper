Imports Microsoft.Office.Tools.Ribbon

Public Class MainRibbon

    Private Sub MainRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ButtonArchive_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonArchive.Click
        MainInvoicearchive.CheckLicense()
    End Sub

    Private Sub ButtonSetting_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSetting.Click
        Dim settingForm As New FormSetting()
        settingForm.ShowDialog()
    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonHelp.Click
        ShowHelpDocument.ShowHelpDocument()
    End Sub
End Class
