<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingsForm
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    Friend WithEvents txtFolderPath As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseFolder As System.Windows.Forms.Button
    Friend WithEvents rdoLocalParse As System.Windows.Forms.RadioButton
    Friend WithEvents rdoOnlineParse As System.Windows.Forms.RadioButton
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents folderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtFolderPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseFolder = New System.Windows.Forms.Button()
        Me.rdoLocalParse = New System.Windows.Forms.RadioButton()
        Me.rdoOnlineParse = New System.Windows.Forms.RadioButton()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.folderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SuspendLayout()
        '
        'txtFolderPath
        '
        Me.txtFolderPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFolderPath.Location = New System.Drawing.Point(9, 10)
        Me.txtFolderPath.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtFolderPath.Name = "txtFolderPath"
        Me.txtFolderPath.Size = New System.Drawing.Size(328, 21)
        Me.txtFolderPath.TabIndex = 0
        '
        'btnBrowseFolder
        '
        Me.btnBrowseFolder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseFolder.Location = New System.Drawing.Point(340, 9)
        Me.btnBrowseFolder.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnBrowseFolder.Name = "btnBrowseFolder"
        Me.btnBrowseFolder.Size = New System.Drawing.Size(70, 22)
        Me.btnBrowseFolder.TabIndex = 1
        Me.btnBrowseFolder.Text = "选择文件夹"
        Me.btnBrowseFolder.UseVisualStyleBackColor = True
        '
        'rdoLocalParse
        '
        Me.rdoLocalParse.AutoSize = True
        Me.rdoLocalParse.Location = New System.Drawing.Point(9, 44)
        Me.rdoLocalParse.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rdoLocalParse.Name = "rdoLocalParse"
        Me.rdoLocalParse.Size = New System.Drawing.Size(71, 16)
        Me.rdoLocalParse.TabIndex = 2
        Me.rdoLocalParse.TabStop = True
        Me.rdoLocalParse.Text = "本地解析"
        Me.rdoLocalParse.UseVisualStyleBackColor = True
        '
        'rdoOnlineParse
        '
        Me.rdoOnlineParse.AutoSize = True
        Me.rdoOnlineParse.Location = New System.Drawing.Point(92, 44)
        Me.rdoOnlineParse.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rdoOnlineParse.Name = "rdoOnlineParse"
        Me.rdoOnlineParse.Size = New System.Drawing.Size(71, 16)
        Me.rdoOnlineParse.TabIndex = 3
        Me.rdoOnlineParse.TabStop = True
        Me.rdoOnlineParse.Text = "在线解析"
        Me.rdoOnlineParse.UseVisualStyleBackColor = True
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(9, 75)
        Me.lblVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(41, 12)
        Me.lblVersion.TabIndex = 4
        Me.lblVersion.Text = "版本："
        '
        'SettingsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 104)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.rdoOnlineParse)
        Me.Controls.Add(Me.rdoLocalParse)
        Me.Controls.Add(Me.btnBrowseFolder)
        Me.Controls.Add(Me.txtFolderPath)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "设置"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
End Class
