<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_Settings
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.chkShowPlainText = New System.Windows.Forms.CheckBox()
        Me.txtSecretKey = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtApiKey = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtRequestUrl = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.radBaiduOCR = New System.Windows.Forms.RadioButton()
        Me.radOfflineMode = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSelectPath = New System.Windows.Forms.Button()
        Me.txtArchivePath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkDebugMode = New System.Windows.Forms.CheckBox()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.btnViewLog = New System.Windows.Forms.Button()
        Me.chkLog = New System.Windows.Forms.CheckBox()
        Me.btnConfirm = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(666, 443)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.chkShowPlainText)
        Me.TabPage1.Controls.Add(Me.txtSecretKey)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.txtApiKey)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.txtRequestUrl)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.radBaiduOCR)
        Me.TabPage1.Controls.Add(Me.radOfflineMode)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.btnSelectPath)
        Me.TabPage1.Controls.Add(Me.txtArchivePath)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(0)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(658, 414)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "发票设置"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'chkShowPlainText
        '
        Me.chkShowPlainText.AutoSize = True
        Me.chkShowPlainText.Location = New System.Drawing.Point(44, 371)
        Me.chkShowPlainText.Name = "chkShowPlainText"
        Me.chkShowPlainText.Size = New System.Drawing.Size(89, 19)
        Me.chkShowPlainText.TabIndex = 12
        Me.chkShowPlainText.Text = "明文显示"
        Me.chkShowPlainText.UseVisualStyleBackColor = True
        '
        'txtSecretKey
        '
        Me.txtSecretKey.Location = New System.Drawing.Point(145, 315)
        Me.txtSecretKey.Name = "txtSecretKey"
        Me.txtSecretKey.Size = New System.Drawing.Size(431, 25)
        Me.txtSecretKey.TabIndex = 11
        Me.txtSecretKey.UseSystemPasswordChar = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(41, 318)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(87, 15)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Secret Key"
        '
        'txtApiKey
        '
        Me.txtApiKey.Location = New System.Drawing.Point(145, 258)
        Me.txtApiKey.Name = "txtApiKey"
        Me.txtApiKey.Size = New System.Drawing.Size(431, 25)
        Me.txtApiKey.TabIndex = 9
        Me.txtApiKey.UseSystemPasswordChar = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(41, 261)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 15)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "API KEY"
        '
        'txtRequestUrl
        '
        Me.txtRequestUrl.Enabled = False
        Me.txtRequestUrl.Location = New System.Drawing.Point(145, 197)
        Me.txtRequestUrl.Name = "txtRequestUrl"
        Me.txtRequestUrl.Size = New System.Drawing.Size(431, 25)
        Me.txtRequestUrl.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(41, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 15)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "请求地址"
        '
        'radBaiduOCR
        '
        Me.radBaiduOCR.AutoSize = True
        Me.radBaiduOCR.Location = New System.Drawing.Point(346, 135)
        Me.radBaiduOCR.Name = "radBaiduOCR"
        Me.radBaiduOCR.Size = New System.Drawing.Size(92, 19)
        Me.radBaiduOCR.TabIndex = 5
        Me.radBaiduOCR.Text = "BaiduOCR"
        Me.radBaiduOCR.UseVisualStyleBackColor = True
        '
        'radOfflineMode
        '
        Me.radOfflineMode.AutoSize = True
        Me.radOfflineMode.Checked = True
        Me.radOfflineMode.Location = New System.Drawing.Point(155, 133)
        Me.radOfflineMode.Name = "radOfflineMode"
        Me.radOfflineMode.Size = New System.Drawing.Size(88, 19)
        Me.radOfflineMode.TabIndex = 4
        Me.radOfflineMode.TabStop = True
        Me.radOfflineMode.Text = "本地模式"
        Me.radOfflineMode.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(39, 135)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 15)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "OCR模式"
        '
        'btnSelectPath
        '
        Me.btnSelectPath.Location = New System.Drawing.Point(549, 48)
        Me.btnSelectPath.Name = "btnSelectPath"
        Me.btnSelectPath.Size = New System.Drawing.Size(85, 27)
        Me.btnSelectPath.TabIndex = 2
        Me.btnSelectPath.Text = "选择位置"
        Me.btnSelectPath.UseVisualStyleBackColor = True
        '
        'txtArchivePath
        '
        Me.txtArchivePath.Location = New System.Drawing.Point(101, 48)
        Me.txtArchivePath.Name = "txtArchivePath"
        Me.txtArchivePath.Size = New System.Drawing.Size(430, 25)
        Me.txtArchivePath.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "归档位置"
        '
        'chkDebugMode
        '
        Me.chkDebugMode.AutoSize = True
        Me.chkDebugMode.Location = New System.Drawing.Point(48, 450)
        Me.chkDebugMode.Name = "chkDebugMode"
        Me.chkDebugMode.Size = New System.Drawing.Size(89, 19)
        Me.chkDebugMode.TabIndex = 13
        Me.chkDebugMode.Text = "调试模式"
        Me.chkDebugMode.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(48, 476)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(85, 27)
        Me.btnReset.TabIndex = 14
        Me.btnReset.Text = "重置"
        Me.btnReset.UseVisualStyleBackColor = True
        Me.btnReset.Visible = False
        '
        'btnViewLog
        '
        Me.btnViewLog.Location = New System.Drawing.Point(172, 476)
        Me.btnViewLog.Name = "btnViewLog"
        Me.btnViewLog.Size = New System.Drawing.Size(85, 27)
        Me.btnViewLog.TabIndex = 16
        Me.btnViewLog.Text = "打开日志"
        Me.btnViewLog.UseVisualStyleBackColor = True
        Me.btnViewLog.Visible = False
        '
        'chkLog
        '
        Me.chkLog.AutoSize = True
        Me.chkLog.Location = New System.Drawing.Point(172, 450)
        Me.chkLog.Name = "chkLog"
        Me.chkLog.Size = New System.Drawing.Size(89, 19)
        Me.chkLog.TabIndex = 15
        Me.chkLog.Text = "记录日志"
        Me.chkLog.UseVisualStyleBackColor = True
        '
        'btnConfirm
        '
        Me.btnConfirm.Image = Global.XYOutlookPlugin.My.Resources.Resources.Submit
        Me.btnConfirm.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnConfirm.Location = New System.Drawing.Point(431, 455)
        Me.btnConfirm.Name = "btnConfirm"
        Me.btnConfirm.Size = New System.Drawing.Size(75, 40)
        Me.btnConfirm.TabIndex = 17
        Me.btnConfirm.Text = "保存"
        Me.btnConfirm.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnConfirm.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Image = Global.XYOutlookPlugin.My.Resources.Resources.Cancel
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(553, 455)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 40)
        Me.btnCancel.TabIndex = 18
        Me.btnCancel.Text = "取消"
        Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Form_Settings
        '
        Me.AcceptButton = Me.btnConfirm
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(668, 508)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnConfirm)
        Me.Controls.Add(Me.btnViewLog)
        Me.Controls.Add(Me.chkLog)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.chkDebugMode)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.HelpButton = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_Settings"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "设置"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TabControl1 As Windows.Forms.TabControl
    Friend WithEvents TabPage1 As Windows.Forms.TabPage
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnSelectPath As Windows.Forms.Button
    Friend WithEvents txtArchivePath As Windows.Forms.TextBox
    Friend WithEvents radBaiduOCR As Windows.Forms.RadioButton
    Friend WithEvents radOfflineMode As Windows.Forms.RadioButton
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents txtSecretKey As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents txtApiKey As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents txtRequestUrl As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents chkShowPlainText As Windows.Forms.CheckBox
    Friend WithEvents chkDebugMode As Windows.Forms.CheckBox
    Friend WithEvents btnReset As Windows.Forms.Button
    Friend WithEvents btnViewLog As Windows.Forms.Button
    Friend WithEvents chkLog As Windows.Forms.CheckBox
    Friend WithEvents btnConfirm As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
End Class
