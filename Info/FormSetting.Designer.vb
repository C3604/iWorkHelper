<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormSetting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormSetting))
        Me.TabSetting = New System.Windows.Forms.TabControl()
        Me.TabPageInvoice = New System.Windows.Forms.TabPage()
        Me.RadioButtonOCROffline = New System.Windows.Forms.RadioButton()
        Me.LabelOCRMode = New System.Windows.Forms.Label()
        Me.RadioButtonOCROnline = New System.Windows.Forms.RadioButton()
        Me.ButtonHelp = New System.Windows.Forms.Button()
        Me.CheckBoxShowToken = New System.Windows.Forms.CheckBox()
        Me.TextBoxClientSecret = New System.Windows.Forms.TextBox()
        Me.TextBoxClientId = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBoxOCRRequestAddress = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ButtonBrowserpath = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBoxArchivePath = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.ButtonSave = New System.Windows.Forms.Button()
        Me.ButtonSubmit = New System.Windows.Forms.Button()
        Me.CheckBoxdebug = New System.Windows.Forms.CheckBox()
        Me.CheckBoxlog = New System.Windows.Forms.CheckBox()
        Me.Buttonlog = New System.Windows.Forms.Button()
        Me.ButtonEmptySetting = New System.Windows.Forms.Button()
        Me.TabSetting.SuspendLayout()
        Me.TabPageInvoice.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabSetting
        '
        Me.TabSetting.Alignment = System.Windows.Forms.TabAlignment.Left
        Me.TabSetting.Controls.Add(Me.TabPageInvoice)
        Me.TabSetting.Location = New System.Drawing.Point(0, 0)
        Me.TabSetting.Multiline = True
        Me.TabSetting.Name = "TabSetting"
        Me.TabSetting.Padding = New System.Drawing.Point(0, 0)
        Me.TabSetting.SelectedIndex = 0
        Me.TabSetting.Size = New System.Drawing.Size(644, 398)
        Me.TabSetting.TabIndex = 0
        '
        'TabPageInvoice
        '
        Me.TabPageInvoice.Controls.Add(Me.RadioButtonOCROffline)
        Me.TabPageInvoice.Controls.Add(Me.LabelOCRMode)
        Me.TabPageInvoice.Controls.Add(Me.RadioButtonOCROnline)
        Me.TabPageInvoice.Controls.Add(Me.ButtonHelp)
        Me.TabPageInvoice.Controls.Add(Me.CheckBoxShowToken)
        Me.TabPageInvoice.Controls.Add(Me.TextBoxClientSecret)
        Me.TabPageInvoice.Controls.Add(Me.TextBoxClientId)
        Me.TabPageInvoice.Controls.Add(Me.Label8)
        Me.TabPageInvoice.Controls.Add(Me.TextBoxOCRRequestAddress)
        Me.TabPageInvoice.Controls.Add(Me.Label7)
        Me.TabPageInvoice.Controls.Add(Me.ButtonBrowserpath)
        Me.TabPageInvoice.Controls.Add(Me.Label6)
        Me.TabPageInvoice.Controls.Add(Me.TextBoxArchivePath)
        Me.TabPageInvoice.Controls.Add(Me.Label3)
        Me.TabPageInvoice.Location = New System.Drawing.Point(26, 4)
        Me.TabPageInvoice.Name = "TabPageInvoice"
        Me.TabPageInvoice.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageInvoice.Size = New System.Drawing.Size(614, 390)
        Me.TabPageInvoice.TabIndex = 0
        Me.TabPageInvoice.Text = "发票"
        Me.TabPageInvoice.UseVisualStyleBackColor = True
        '
        'RadioButtonOCROffline
        '
        Me.RadioButtonOCROffline.AutoSize = True
        Me.RadioButtonOCROffline.Enabled = False
        Me.RadioButtonOCROffline.Location = New System.Drawing.Point(259, 132)
        Me.RadioButtonOCROffline.Name = "RadioButtonOCROffline"
        Me.RadioButtonOCROffline.Size = New System.Drawing.Size(178, 19)
        Me.RadioButtonOCROffline.TabIndex = 16
        Me.RadioButtonOCROffline.TabStop = True
        Me.RadioButtonOCROffline.Text = "离线模式（暂不可用）"
        Me.RadioButtonOCROffline.UseVisualStyleBackColor = True
        '
        'LabelOCRMode
        '
        Me.LabelOCRMode.AutoSize = True
        Me.LabelOCRMode.Location = New System.Drawing.Point(27, 132)
        Me.LabelOCRMode.Name = "LabelOCRMode"
        Me.LabelOCRMode.Size = New System.Drawing.Size(61, 15)
        Me.LabelOCRMode.TabIndex = 15
        Me.LabelOCRMode.Text = "OCR模式"
        '
        'RadioButtonOCROnline
        '
        Me.RadioButtonOCROnline.AutoSize = True
        Me.RadioButtonOCROnline.Checked = True
        Me.RadioButtonOCROnline.Enabled = False
        Me.RadioButtonOCROnline.Location = New System.Drawing.Point(111, 130)
        Me.RadioButtonOCROnline.Name = "RadioButtonOCROnline"
        Me.RadioButtonOCROnline.Size = New System.Drawing.Size(88, 19)
        Me.RadioButtonOCROnline.TabIndex = 14
        Me.RadioButtonOCROnline.TabStop = True
        Me.RadioButtonOCROnline.Text = "在线模式"
        Me.RadioButtonOCROnline.UseVisualStyleBackColor = True
        '
        'ButtonHelp
        '
        Me.ButtonHelp.Image = Global.XYOutlookPlugin.My.Resources.Resources.Help
        Me.ButtonHelp.Location = New System.Drawing.Point(566, 13)
        Me.ButtonHelp.Name = "ButtonHelp"
        Me.ButtonHelp.Size = New System.Drawing.Size(35, 35)
        Me.ButtonHelp.TabIndex = 13
        Me.ButtonHelp.UseVisualStyleBackColor = True
        '
        'CheckBoxShowToken
        '
        Me.CheckBoxShowToken.AutoSize = True
        Me.CheckBoxShowToken.Location = New System.Drawing.Point(30, 342)
        Me.CheckBoxShowToken.Name = "CheckBoxShowToken"
        Me.CheckBoxShowToken.Size = New System.Drawing.Size(89, 19)
        Me.CheckBoxShowToken.TabIndex = 3
        Me.CheckBoxShowToken.Text = "明文显示"
        Me.CheckBoxShowToken.UseVisualStyleBackColor = True
        '
        'TextBoxClientSecret
        '
        Me.TextBoxClientSecret.Location = New System.Drawing.Point(111, 296)
        Me.TextBoxClientSecret.Name = "TextBoxClientSecret"
        Me.TextBoxClientSecret.Size = New System.Drawing.Size(439, 25)
        Me.TextBoxClientSecret.TabIndex = 1
        Me.TextBoxClientSecret.UseSystemPasswordChar = True
        '
        'TextBoxClientId
        '
        Me.TextBoxClientId.Location = New System.Drawing.Point(111, 239)
        Me.TextBoxClientId.Name = "TextBoxClientId"
        Me.TextBoxClientId.Size = New System.Drawing.Size(439, 25)
        Me.TextBoxClientId.TabIndex = 1
        Me.TextBoxClientId.UseSystemPasswordChar = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(24, 299)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(87, 15)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Secret Key"
        '
        'TextBoxOCRRequestAddress
        '
        Me.TextBoxOCRRequestAddress.Enabled = False
        Me.TextBoxOCRRequestAddress.Location = New System.Drawing.Point(111, 182)
        Me.TextBoxOCRRequestAddress.Name = "TextBoxOCRRequestAddress"
        Me.TextBoxOCRRequestAddress.Size = New System.Drawing.Size(439, 25)
        Me.TextBoxOCRRequestAddress.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(24, 242)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(63, 15)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "API Key"
        '
        'ButtonBrowserpath
        '
        Me.ButtonBrowserpath.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonBrowserpath.Location = New System.Drawing.Point(477, 48)
        Me.ButtonBrowserpath.Name = "ButtonBrowserpath"
        Me.ButtonBrowserpath.Size = New System.Drawing.Size(73, 26)
        Me.ButtonBrowserpath.TabIndex = 2
        Me.ButtonBrowserpath.Text = "浏览"
        Me.ButtonBrowserpath.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(24, 185)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 15)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "请求地址"
        '
        'TextBoxArchivePath
        '
        Me.TextBoxArchivePath.Location = New System.Drawing.Point(111, 49)
        Me.TextBoxArchivePath.Name = "TextBoxArchivePath"
        Me.TextBoxArchivePath.Size = New System.Drawing.Size(350, 25)
        Me.TextBoxArchivePath.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(24, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 15)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "归档路径"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Image = Global.XYOutlookPlugin.My.Resources.Resources.Cancel
        Me.ButtonCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonCancel.Location = New System.Drawing.Point(527, 416)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(100, 40)
        Me.ButtonCancel.TabIndex = 1
        Me.ButtonCancel.Text = "取消"
        Me.ButtonCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'ButtonSave
        '
        Me.ButtonSave.Image = Global.XYOutlookPlugin.My.Resources.Resources.Save
        Me.ButtonSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonSave.Location = New System.Drawing.Point(401, 416)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(100, 40)
        Me.ButtonSave.TabIndex = 1
        Me.ButtonSave.Text = "保存"
        Me.ButtonSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonSave.UseVisualStyleBackColor = True
        '
        'ButtonSubmit
        '
        Me.ButtonSubmit.Image = CType(resources.GetObject("ButtonSubmit.Image"), System.Drawing.Image)
        Me.ButtonSubmit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonSubmit.Location = New System.Drawing.Point(273, 417)
        Me.ButtonSubmit.Name = "ButtonSubmit"
        Me.ButtonSubmit.Size = New System.Drawing.Size(100, 40)
        Me.ButtonSubmit.TabIndex = 1
        Me.ButtonSubmit.Text = "确认"
        Me.ButtonSubmit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonSubmit.UseVisualStyleBackColor = True
        '
        'CheckBoxdebug
        '
        Me.CheckBoxdebug.AutoSize = True
        Me.CheckBoxdebug.Location = New System.Drawing.Point(31, 404)
        Me.CheckBoxdebug.Name = "CheckBoxdebug"
        Me.CheckBoxdebug.Size = New System.Drawing.Size(89, 19)
        Me.CheckBoxdebug.TabIndex = 2
        Me.CheckBoxdebug.Text = "调试模式"
        Me.CheckBoxdebug.UseVisualStyleBackColor = True
        '
        'CheckBoxlog
        '
        Me.CheckBoxlog.AutoSize = True
        Me.CheckBoxlog.Location = New System.Drawing.Point(137, 405)
        Me.CheckBoxlog.Name = "CheckBoxlog"
        Me.CheckBoxlog.Size = New System.Drawing.Size(89, 19)
        Me.CheckBoxlog.TabIndex = 2
        Me.CheckBoxlog.Text = "记录日志"
        Me.CheckBoxlog.UseVisualStyleBackColor = True
        '
        'Buttonlog
        '
        Me.Buttonlog.Location = New System.Drawing.Point(137, 430)
        Me.Buttonlog.Name = "Buttonlog"
        Me.Buttonlog.Size = New System.Drawing.Size(97, 27)
        Me.Buttonlog.TabIndex = 3
        Me.Buttonlog.Text = "查看日志"
        Me.Buttonlog.UseVisualStyleBackColor = True
        '
        'ButtonEmptySetting
        '
        Me.ButtonEmptySetting.Location = New System.Drawing.Point(31, 429)
        Me.ButtonEmptySetting.Name = "ButtonEmptySetting"
        Me.ButtonEmptySetting.Size = New System.Drawing.Size(97, 27)
        Me.ButtonEmptySetting.TabIndex = 4
        Me.ButtonEmptySetting.Text = "重置"
        Me.ButtonEmptySetting.UseVisualStyleBackColor = True
        Me.ButtonEmptySetting.Visible = False
        '
        'FormSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(654, 464)
        Me.Controls.Add(Me.ButtonEmptySetting)
        Me.Controls.Add(Me.Buttonlog)
        Me.Controls.Add(Me.CheckBoxlog)
        Me.Controls.Add(Me.CheckBoxdebug)
        Me.Controls.Add(Me.ButtonSave)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonSubmit)
        Me.Controls.Add(Me.TabSetting)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormSetting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置"
        Me.TabSetting.ResumeLayout(False)
        Me.TabPageInvoice.ResumeLayout(False)
        Me.TabPageInvoice.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TabSetting As Windows.Forms.TabControl
    Friend WithEvents ButtonSubmit As Windows.Forms.Button
    Friend WithEvents ButtonCancel As Windows.Forms.Button
    Friend WithEvents ButtonSave As Windows.Forms.Button
    Friend WithEvents CheckBoxdebug As Windows.Forms.CheckBox
    Friend WithEvents CheckBoxlog As Windows.Forms.CheckBox
    Friend WithEvents Buttonlog As Windows.Forms.Button
    Friend WithEvents ButtonEmptySetting As Windows.Forms.Button
    Friend WithEvents TabPageInvoice As Windows.Forms.TabPage
    Friend WithEvents ButtonHelp As Windows.Forms.Button
    Friend WithEvents CheckBoxShowToken As Windows.Forms.CheckBox
    Friend WithEvents TextBoxClientSecret As Windows.Forms.TextBox
    Friend WithEvents TextBoxClientId As Windows.Forms.TextBox
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents TextBoxOCRRequestAddress As Windows.Forms.TextBox
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents ButtonBrowserpath As Windows.Forms.Button
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents TextBoxArchivePath As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents LabelOCRMode As Windows.Forms.Label
    Friend WithEvents RadioButtonOCROnline As Windows.Forms.RadioButton
    Friend WithEvents RadioButtonOCROffline As Windows.Forms.RadioButton
End Class
