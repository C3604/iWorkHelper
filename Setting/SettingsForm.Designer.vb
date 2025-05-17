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

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_ArchivePath = New System.Windows.Forms.TextBox()
        Me.btn_ArchivePath = New System.Windows.Forms.Button()
        Me.ckb_Debug = New System.Windows.Forms.CheckBox()
        Me.btn_SaveSettings = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.Lab_PathError = New System.Windows.Forms.Label()
        Me.ckb_MergeDidiFiles = New System.Windows.Forms.CheckBox()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 26)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "归档位置"
        '
        'txt_ArchivePath
        '
        Me.txt_ArchivePath.Location = New System.Drawing.Point(64, 22)
        Me.txt_ArchivePath.Margin = New System.Windows.Forms.Padding(2)
        Me.txt_ArchivePath.Name = "txt_ArchivePath"
        Me.txt_ArchivePath.Size = New System.Drawing.Size(260, 21)
        Me.txt_ArchivePath.TabIndex = 1
        '
        'btn_ArchivePath
        '
        Me.btn_ArchivePath.Location = New System.Drawing.Point(334, 21)
        Me.btn_ArchivePath.Margin = New System.Windows.Forms.Padding(2)
        Me.btn_ArchivePath.Name = "btn_ArchivePath"
        Me.btn_ArchivePath.Size = New System.Drawing.Size(64, 23)
        Me.btn_ArchivePath.TabIndex = 2
        Me.btn_ArchivePath.Text = "浏览"
        Me.btn_ArchivePath.UseVisualStyleBackColor = True
        '
        'ckb_Debug
        '
        Me.ckb_Debug.AutoSize = True
        Me.ckb_Debug.Location = New System.Drawing.Point(9, 97)
        Me.ckb_Debug.Margin = New System.Windows.Forms.Padding(2)
        Me.ckb_Debug.Name = "ckb_Debug"
        Me.ckb_Debug.Size = New System.Drawing.Size(72, 16)
        Me.ckb_Debug.TabIndex = 3
        Me.ckb_Debug.Text = "调试模式"
        Me.ckb_Debug.UseVisualStyleBackColor = True
        '
        'btn_SaveSettings
        '
        Me.btn_SaveSettings.Location = New System.Drawing.Point(101, 111)
        Me.btn_SaveSettings.Margin = New System.Windows.Forms.Padding(2)
        Me.btn_SaveSettings.Name = "btn_SaveSettings"
        Me.btn_SaveSettings.Size = New System.Drawing.Size(158, 37)
        Me.btn_SaveSettings.TabIndex = 4
        Me.btn_SaveSettings.Text = "确认"
        Me.btn_SaveSettings.UseVisualStyleBackColor = True
        '
        'FolderBrowserDialog
        '
        Me.FolderBrowserDialog.Description = "选择一个文件夹"
        Me.FolderBrowserDialog.Tag = ""
        '
        'Lab_PathError
        '
        Me.Lab_PathError.AutoSize = True
        Me.Lab_PathError.Enabled = False
        Me.Lab_PathError.ForeColor = System.Drawing.Color.Red
        Me.Lab_PathError.Location = New System.Drawing.Point(62, 50)
        Me.Lab_PathError.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lab_PathError.Name = "Lab_PathError"
        Me.Lab_PathError.Size = New System.Drawing.Size(239, 12)
        Me.Lab_PathError.TabIndex = 6
        Me.Lab_PathError.Text = "*归档位置有误，请确认文件夹路径是否合理"
        Me.Lab_PathError.Visible = False
        '
        'ckb_MergeDidiFiles
        '
        Me.ckb_MergeDidiFiles.AutoSize = True
        Me.ckb_MergeDidiFiles.Location = New System.Drawing.Point(9, 77)
        Me.ckb_MergeDidiFiles.Margin = New System.Windows.Forms.Padding(2)
        Me.ckb_MergeDidiFiles.Name = "ckb_MergeDidiFiles"
        Me.ckb_MergeDidiFiles.Size = New System.Drawing.Size(144, 16)
        Me.ckb_MergeDidiFiles.TabIndex = 8
        Me.ckb_MergeDidiFiles.Text = "合并滴滴发票和行程单"
        Me.ckb_MergeDidiFiles.UseVisualStyleBackColor = True
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(284, 123)
        Me.lblVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(101, 12)
        Me.lblVersion.TabIndex = 7
        Me.lblVersion.Text = "未获取到版本信息"
        '
        'SettingsForm
        '
        Me.AcceptButton = Me.btn_SaveSettings
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 153)
        Me.Controls.Add(Me.ckb_MergeDidiFiles)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.Lab_PathError)
        Me.Controls.Add(Me.btn_SaveSettings)
        Me.Controls.Add(Me.ckb_Debug)
        Me.Controls.Add(Me.btn_ArchivePath)
        Me.Controls.Add(Me.txt_ArchivePath)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.HelpButton = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "设置"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txt_ArchivePath As Windows.Forms.TextBox
    Friend WithEvents btn_ArchivePath As Windows.Forms.Button
    Friend WithEvents ckb_Debug As Windows.Forms.CheckBox
    Friend WithEvents btn_SaveSettings As Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog As Windows.Forms.FolderBrowserDialog
    Friend WithEvents Lab_PathError As Windows.Forms.Label
    Friend WithEvents ckb_MergeDidiFiles As Windows.Forms.CheckBox
    Friend WithEvents lblVersion As Windows.Forms.Label
End Class
