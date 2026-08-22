<<<<<<< HEAD
Imports System.Windows.Forms
Imports System.Reflection

=======
>>>>>>> 1b5cb2d7788e08ddb78646384ca7b0118c66dc37
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

            Dim version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            lblVersion.Text = "版本：" & If(version Is Nothing, String.Empty, version.ToString())
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

<<<<<<< HEAD
    Private Function IsValidHttpUrl(url As String) As Boolean
        Dim result As Uri = Nothing
        Return Uri.TryCreate(url, UriKind.Absolute, result) _
            AndAlso (result.Scheme = Uri.UriSchemeHttp OrElse result.Scheme = Uri.UriSchemeHttps)
    End Function

    Private Function ClampDecimal(value As Integer, min As Decimal, max As Decimal, fallback As Integer) As Decimal
        Dim v As Decimal = value
        If value <= 0 Then v = fallback
        If v < min Then v = min
        If v > max Then v = max
        Return v
    End Function

    Private Sub lblTplHint_Click(sender As Object, e As EventArgs)

    End Sub

    ''' <summary>获取用户展示版本，优先级：InformationalVersion > FileVersion > AssemblyVersion。</summary>
    Private Function GetApplicationVersion() As String
        Try
            ' 第一优先级：产品展示版本。它保留 a.b.YYMMDD.revision 语义，
            ' 不受 ClickOnce/VSTO 内部部署版本的字段上限约束。
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim infoAttr = CType(Attribute.GetCustomAttribute(asm, GetType(AssemblyInformationalVersionAttribute)), AssemblyInformationalVersionAttribute)
                If infoAttr IsNot Nothing AndAlso Not String.IsNullOrEmpty(infoAttr.InformationalVersion) Then
                    AppLogger.Debug("使用 AssemblyInformationalVersion：" & infoAttr.InformationalVersion)
                    Return infoAttr.InformationalVersion
                End If
            Catch ex As Exception
                AppLogger.Debug("读取 InformationalVersion 失败：" & ex.Message)
            End Try

            ' 第二优先级：AssemblyFileVersion
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim fileAttr = CType(Attribute.GetCustomAttribute(asm, GetType(AssemblyFileVersionAttribute)), AssemblyFileVersionAttribute)
                If fileAttr IsNot Nothing AndAlso Not String.IsNullOrEmpty(fileAttr.Version) Then
                    AppLogger.Debug("使用 AssemblyFileVersion：" & fileAttr.Version)
                    Return fileAttr.Version
                End If
            Catch ex As Exception
                AppLogger.Debug("读取 FileVersion 失败：" & ex.Message)
            End Try

            ' 第三优先级：AssemblyVersion（程序集版本）
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim ver = asm.GetName().Version
                If ver IsNot Nothing Then
                    Dim version = String.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision)
                    AppLogger.Debug("使用 AssemblyVersion：" & version)
                    Return version
                End If
            Catch ex As Exception
                AppLogger.Debug("读取 AssemblyVersion 失败：" & ex.Message)
            End Try

            ' 第四优先级：未知
            AppLogger.Warn("无法读取任何有效版本号")
            Return "未知"

        Catch ex As Exception
            AppLogger.Warn("获取应用版本号异常：" & ex.Message)
            Return "未知"
        End Try
    End Function

    ''' <summary>
    ''' 应用内外网版本的布局调整。
    ''' 设计器原始布局适配外网版本（包含在线 OCR）。
    ''' 内网版本隐藏解析模式选择和在线 OCR 配置，将命名规则分组框上移，缩短窗体高度。
    ''' </summary>
    Private Sub ApplyEditionLayout()
        If BuildFeatures.OnlineParserEnabled Then
            ' 外网版本：保持设计器原始布局
            Return
        End If

        ' 内网版本：应用紧凑布局
        Const GrpNamingHeight As Integer = 155
        Const GrpNamingNewY As Integer = 40
        Const LblVersionNewY As Integer = GrpNamingNewY + GrpNamingHeight + 8
        Const BtnSaveNewY As Integer = LblVersionNewY + 20
        Const FormNewHeight As Integer = BtnSaveNewY + 35

        ' 隐藏解析模式选择（内网版只能用本地）
        rdoLocalParse.Visible = False
        rdoOnlineParse.Visible = False

        ' 隐藏在线 OCR 分组框和相关说明
        grpOcr.Visible = False

        ' 上移命名规则分组框
        grpNaming.Location = New System.Drawing.Point(grpNaming.Location.X, GrpNamingNewY)

        ' 上移版本标签
        lblVersion.Location = New System.Drawing.Point(lblVersion.Location.X, LblVersionNewY)

        ' 上移保存按钮
        btnSave.Location = New System.Drawing.Point(btnSave.Location.X, BtnSaveNewY)

        ' 缩短窗体高度
        Me.ClientSize = New System.Drawing.Size(Me.ClientSize.Width, FormNewHeight)

        AppLogger.Debug("已应用内网版紧凑布局（隐藏在线解析、命名规则上移、窗体高度缩短）")
    End Sub

=======
>>>>>>> 1b5cb2d7788e08ddb78646384ca7b0118c66dc37
End Class
