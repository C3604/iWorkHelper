Imports System.Windows.Forms
Imports System.Reflection
Public Class SettingsForm

    Private isLoading As Boolean

    Private Sub SettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isLoading = True
        Try
            txtFolderPath.Text = My.Settings.ArchiveFolderPath

            If String.Equals(My.Settings.ParseMode, "Online", StringComparison.OrdinalIgnoreCase) Then
                rdoOnlineParse.Checked = True
            Else
                rdoLocalParse.Checked = True
            End If

            lblVersion.Text = "版本：" & GetApplicationVersion()
        Finally
            isLoading = False
        End Try
    End Sub

    Private Sub btnBrowseFolder_Click(sender As Object, e As EventArgs) Handles btnBrowseFolder.Click
        folderBrowserDialog1.SelectedPath = txtFolderPath.Text

        If folderBrowserDialog1.ShowDialog(Me) = DialogResult.OK Then
            txtFolderPath.Text = folderBrowserDialog1.SelectedPath
            SaveFolderPath(txtFolderPath.Text)
        End If
    End Sub

    Private Sub txtFolderPath_Leave(sender As Object, e As EventArgs) Handles txtFolderPath.Leave
        SaveFolderPath(txtFolderPath.Text)
    End Sub

    Private Sub rdoLocalParse_CheckedChanged(sender As Object, e As EventArgs) Handles rdoLocalParse.CheckedChanged
        If isLoading OrElse Not rdoLocalParse.Checked Then
            Return
        End If

        My.Settings.ParseMode = "Local"
        My.Settings.Save()
    End Sub

    Private Sub rdoOnlineParse_CheckedChanged(sender As Object, e As EventArgs) Handles rdoOnlineParse.CheckedChanged
        If isLoading OrElse Not rdoOnlineParse.Checked Then
            Return
        End If

        My.Settings.ParseMode = "Online"
        My.Settings.Save()
    End Sub

    Private Sub SaveFolderPath(folderPath As String)
        If isLoading Then
            Return
        End If

        My.Settings.ArchiveFolderPath = If(folderPath, String.Empty).Trim()
        My.Settings.Save()
    End Sub

    ''' <summary>获取用户展示版本，优先读取 InformationalVersion。</summary>
    Private Function GetApplicationVersion() As String
        Try
            Dim asm = Assembly.GetExecutingAssembly()
            Dim attr = CType(Attribute.GetCustomAttribute(asm, GetType(AssemblyInformationalVersionAttribute)), AssemblyInformationalVersionAttribute)
            If attr IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(attr.InformationalVersion) Then
                Return attr.InformationalVersion
            End If

            Dim version = asm.GetName().Version
            Return If(version Is Nothing, "未知", version.ToString())
        Catch
            Return "未知"
        End Try
    End Function

End Class
