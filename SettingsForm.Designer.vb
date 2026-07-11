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
    Friend WithEvents lblFolder As System.Windows.Forms.Label
    Friend WithEvents grpOcr As System.Windows.Forms.GroupBox
    Friend WithEvents chkOcrEnabled As System.Windows.Forms.CheckBox
    Friend WithEvents lblApiKey As System.Windows.Forms.Label
    Friend WithEvents txtApiKey As System.Windows.Forms.TextBox
    Friend WithEvents lblSecretKey As System.Windows.Forms.Label
    Friend WithEvents txtSecretKey As System.Windows.Forms.TextBox
    Friend WithEvents lblApiUrl As System.Windows.Forms.Label
    Friend WithEvents txtApiUrl As System.Windows.Forms.TextBox
    Friend WithEvents lblTokenUrl As System.Windows.Forms.Label
    Friend WithEvents txtTokenUrl As System.Windows.Forms.TextBox
    Friend WithEvents lblTimeout As System.Windows.Forms.Label
    Friend WithEvents numTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblMaxPages As System.Windows.Forms.Label
    Friend WithEvents numMaxPages As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkProbability As System.Windows.Forms.CheckBox
    Friend WithEvents chkLocation As System.Windows.Forms.CheckBox
    Friend WithEvents chkVerify As System.Windows.Forms.CheckBox
    Friend WithEvents chkPreferLocal As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutoFallback As System.Windows.Forms.CheckBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents grpNaming As System.Windows.Forms.GroupBox
    Friend WithEvents lblInvoiceTpl As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceTemplate As System.Windows.Forms.TextBox
    Friend WithEvents lblTripTpl As System.Windows.Forms.Label
    Friend WithEvents txtTripTemplate As System.Windows.Forms.TextBox
    Friend WithEvents lblUnknownTpl As System.Windows.Forms.Label
    Friend WithEvents txtUnknownTemplate As System.Windows.Forms.TextBox
    Friend WithEvents btnTplVars As System.Windows.Forms.Button
    Friend WithEvents btnTplPreview As System.Windows.Forms.Button
    Friend WithEvents btnTestOcr As System.Windows.Forms.Button

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtFolderPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseFolder = New System.Windows.Forms.Button()
        Me.rdoLocalParse = New System.Windows.Forms.RadioButton()
        Me.rdoOnlineParse = New System.Windows.Forms.RadioButton()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.folderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.lblFolder = New System.Windows.Forms.Label()
        Me.grpOcr = New System.Windows.Forms.GroupBox()
        Me.chkOcrEnabled = New System.Windows.Forms.CheckBox()
        Me.lblApiKey = New System.Windows.Forms.Label()
        Me.txtApiKey = New System.Windows.Forms.TextBox()
        Me.lblSecretKey = New System.Windows.Forms.Label()
        Me.txtSecretKey = New System.Windows.Forms.TextBox()
        Me.lblApiUrl = New System.Windows.Forms.Label()
        Me.txtApiUrl = New System.Windows.Forms.TextBox()
        Me.lblTokenUrl = New System.Windows.Forms.Label()
        Me.txtTokenUrl = New System.Windows.Forms.TextBox()
        Me.lblTimeout = New System.Windows.Forms.Label()
        Me.numTimeout = New System.Windows.Forms.NumericUpDown()
        Me.lblMaxPages = New System.Windows.Forms.Label()
        Me.numMaxPages = New System.Windows.Forms.NumericUpDown()
        Me.chkProbability = New System.Windows.Forms.CheckBox()
        Me.chkLocation = New System.Windows.Forms.CheckBox()
        Me.chkVerify = New System.Windows.Forms.CheckBox()
        Me.chkPreferLocal = New System.Windows.Forms.CheckBox()
        Me.chkAutoFallback = New System.Windows.Forms.CheckBox()
        Me.btnTestOcr = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.grpNaming = New System.Windows.Forms.GroupBox()
        Me.lblInvoiceTpl = New System.Windows.Forms.Label()
        Me.txtInvoiceTemplate = New System.Windows.Forms.TextBox()
        Me.lblTripTpl = New System.Windows.Forms.Label()
        Me.txtTripTemplate = New System.Windows.Forms.TextBox()
        Me.lblUnknownTpl = New System.Windows.Forms.Label()
        Me.txtUnknownTemplate = New System.Windows.Forms.TextBox()
        Me.btnTplVars = New System.Windows.Forms.Button()
        Me.btnTplPreview = New System.Windows.Forms.Button()
        Me.grpOcr.SuspendLayout()
        CType(Me.numTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numMaxPages, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNaming.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtFolderPath
        '
        Me.txtFolderPath.Location = New System.Drawing.Point(75, 9)
        Me.txtFolderPath.Name = "txtFolderPath"
        Me.txtFolderPath.Size = New System.Drawing.Size(268, 21)
        Me.txtFolderPath.TabIndex = 0
        '
        'btnBrowseFolder
        '
        Me.btnBrowseFolder.Location = New System.Drawing.Point(349, 8)
        Me.btnBrowseFolder.Name = "btnBrowseFolder"
        Me.btnBrowseFolder.Size = New System.Drawing.Size(90, 23)
        Me.btnBrowseFolder.TabIndex = 1
        Me.btnBrowseFolder.Text = "选择文件夹"
        Me.btnBrowseFolder.UseVisualStyleBackColor = True
        '
        'rdoLocalParse
        '
        Me.rdoLocalParse.AutoSize = True
        Me.rdoLocalParse.Location = New System.Drawing.Point(75, 40)
        Me.rdoLocalParse.Name = "rdoLocalParse"
        Me.rdoLocalParse.Size = New System.Drawing.Size(71, 16)
        Me.rdoLocalParse.TabIndex = 2
        Me.rdoLocalParse.Text = "本地解析"
        Me.rdoLocalParse.UseVisualStyleBackColor = True
        '
        'rdoOnlineParse
        '
        Me.rdoOnlineParse.AutoSize = True
        Me.rdoOnlineParse.Location = New System.Drawing.Point(160, 40)
        Me.rdoOnlineParse.Name = "rdoOnlineParse"
        Me.rdoOnlineParse.Size = New System.Drawing.Size(71, 16)
        Me.rdoOnlineParse.TabIndex = 3
        Me.rdoOnlineParse.Text = "在线解析"
        Me.rdoOnlineParse.UseVisualStyleBackColor = True
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(10, 530)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(41, 12)
        Me.lblVersion.TabIndex = 9
        Me.lblVersion.Text = "版本："
        '
        'lblFolder
        '
        Me.lblFolder.AutoSize = True
        Me.lblFolder.Location = New System.Drawing.Point(9, 12)
        Me.lblFolder.Name = "lblFolder"
        Me.lblFolder.Size = New System.Drawing.Size(65, 12)
        Me.lblFolder.TabIndex = 8
        Me.lblFolder.Text = "归档目录："
        '
        'grpOcr
        '
        Me.grpOcr.Controls.Add(Me.chkOcrEnabled)
        Me.grpOcr.Controls.Add(Me.lblApiKey)
        Me.grpOcr.Controls.Add(Me.txtApiKey)
        Me.grpOcr.Controls.Add(Me.lblSecretKey)
        Me.grpOcr.Controls.Add(Me.txtSecretKey)
        Me.grpOcr.Controls.Add(Me.lblApiUrl)
        Me.grpOcr.Controls.Add(Me.txtApiUrl)
        Me.grpOcr.Controls.Add(Me.lblTokenUrl)
        Me.grpOcr.Controls.Add(Me.txtTokenUrl)
        Me.grpOcr.Controls.Add(Me.lblTimeout)
        Me.grpOcr.Controls.Add(Me.numTimeout)
        Me.grpOcr.Controls.Add(Me.lblMaxPages)
        Me.grpOcr.Controls.Add(Me.numMaxPages)
        Me.grpOcr.Controls.Add(Me.chkProbability)
        Me.grpOcr.Controls.Add(Me.chkLocation)
        Me.grpOcr.Controls.Add(Me.chkVerify)
        Me.grpOcr.Controls.Add(Me.chkPreferLocal)
        Me.grpOcr.Controls.Add(Me.chkAutoFallback)
        Me.grpOcr.Controls.Add(Me.btnTestOcr)
        Me.grpOcr.Location = New System.Drawing.Point(9, 65)
        Me.grpOcr.Name = "grpOcr"
        Me.grpOcr.Size = New System.Drawing.Size(430, 290)
        Me.grpOcr.TabIndex = 4
        Me.grpOcr.TabStop = False
        Me.grpOcr.Text = "在线 OCR（百度智能财务票据识别）"
        '
        'chkOcrEnabled
        '
        Me.chkOcrEnabled.AutoSize = True
        Me.chkOcrEnabled.Location = New System.Drawing.Point(15, 24)
        Me.chkOcrEnabled.Name = "chkOcrEnabled"
        Me.chkOcrEnabled.Size = New System.Drawing.Size(96, 16)
        Me.chkOcrEnabled.TabIndex = 0
        Me.chkOcrEnabled.Text = "启用在线 OCR"
        Me.chkOcrEnabled.UseVisualStyleBackColor = True
        '
        'lblApiKey
        '
        Me.lblApiKey.AutoSize = True
        Me.lblApiKey.Location = New System.Drawing.Point(15, 54)
        Me.lblApiKey.Name = "lblApiKey"
        Me.lblApiKey.Size = New System.Drawing.Size(59, 12)
        Me.lblApiKey.TabIndex = 1
        Me.lblApiKey.Text = "API Key："
        '
        'txtApiKey
        '
        Me.txtApiKey.Location = New System.Drawing.Point(110, 51)
        Me.txtApiKey.Name = "txtApiKey"
        Me.txtApiKey.Size = New System.Drawing.Size(300, 21)
        Me.txtApiKey.TabIndex = 1
        '
        'lblSecretKey
        '
        Me.lblSecretKey.AutoSize = True
        Me.lblSecretKey.Location = New System.Drawing.Point(15, 84)
        Me.lblSecretKey.Name = "lblSecretKey"
        Me.lblSecretKey.Size = New System.Drawing.Size(77, 12)
        Me.lblSecretKey.TabIndex = 2
        Me.lblSecretKey.Text = "Secret Key："
        '
        'txtSecretKey
        '
        Me.txtSecretKey.Location = New System.Drawing.Point(110, 81)
        Me.txtSecretKey.Name = "txtSecretKey"
        Me.txtSecretKey.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSecretKey.Size = New System.Drawing.Size(300, 21)
        Me.txtSecretKey.TabIndex = 2
        '
        'lblApiUrl
        '
        Me.lblApiUrl.AutoSize = True
        Me.lblApiUrl.Location = New System.Drawing.Point(15, 114)
        Me.lblApiUrl.Name = "lblApiUrl"
        Me.lblApiUrl.Size = New System.Drawing.Size(65, 12)
        Me.lblApiUrl.TabIndex = 3
        Me.lblApiUrl.Text = "接口地址："
        '
        'txtApiUrl
        '
        Me.txtApiUrl.Location = New System.Drawing.Point(110, 111)
        Me.txtApiUrl.Name = "txtApiUrl"
        Me.txtApiUrl.Size = New System.Drawing.Size(300, 21)
        Me.txtApiUrl.TabIndex = 3
        '
        'lblTokenUrl
        '
        Me.lblTokenUrl.AutoSize = True
        Me.lblTokenUrl.Location = New System.Drawing.Point(15, 144)
        Me.lblTokenUrl.Name = "lblTokenUrl"
        Me.lblTokenUrl.Size = New System.Drawing.Size(77, 12)
        Me.lblTokenUrl.TabIndex = 4
        Me.lblTokenUrl.Text = "Token 地址："
        '
        'txtTokenUrl
        '
        Me.txtTokenUrl.Location = New System.Drawing.Point(110, 141)
        Me.txtTokenUrl.Name = "txtTokenUrl"
        Me.txtTokenUrl.Size = New System.Drawing.Size(300, 21)
        Me.txtTokenUrl.TabIndex = 4
        '
        'lblTimeout
        '
        Me.lblTimeout.AutoSize = True
        Me.lblTimeout.Location = New System.Drawing.Point(15, 176)
        Me.lblTimeout.Name = "lblTimeout"
        Me.lblTimeout.Size = New System.Drawing.Size(77, 12)
        Me.lblTimeout.TabIndex = 5
        Me.lblTimeout.Text = "超时(毫秒)："
        '
        'numTimeout
        '
        Me.numTimeout.Increment = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.numTimeout.Location = New System.Drawing.Point(110, 174)
        Me.numTimeout.Maximum = New Decimal(New Integer() {120000, 0, 0, 0})
        Me.numTimeout.Minimum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.numTimeout.Name = "numTimeout"
        Me.numTimeout.Size = New System.Drawing.Size(90, 21)
        Me.numTimeout.TabIndex = 5
        Me.numTimeout.Value = New Decimal(New Integer() {30000, 0, 0, 0})
        '
        'lblMaxPages
        '
        Me.lblMaxPages.AutoSize = True
        Me.lblMaxPages.Location = New System.Drawing.Point(230, 176)
        Me.lblMaxPages.Name = "lblMaxPages"
        Me.lblMaxPages.Size = New System.Drawing.Size(89, 12)
        Me.lblMaxPages.TabIndex = 6
        Me.lblMaxPages.Text = "最大识别页数："
        '
        'numMaxPages
        '
        Me.numMaxPages.Location = New System.Drawing.Point(330, 174)
        Me.numMaxPages.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.numMaxPages.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numMaxPages.Name = "numMaxPages"
        Me.numMaxPages.Size = New System.Drawing.Size(60, 21)
        Me.numMaxPages.TabIndex = 6
        Me.numMaxPages.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'chkProbability
        '
        Me.chkProbability.AutoSize = True
        Me.chkProbability.Location = New System.Drawing.Point(15, 208)
        Me.chkProbability.Name = "chkProbability"
        Me.chkProbability.Size = New System.Drawing.Size(84, 16)
        Me.chkProbability.TabIndex = 7
        Me.chkProbability.Text = "返回置信度"
        Me.chkProbability.UseVisualStyleBackColor = True
        '
        'chkLocation
        '
        Me.chkLocation.AutoSize = True
        Me.chkLocation.Location = New System.Drawing.Point(150, 208)
        Me.chkLocation.Name = "chkLocation"
        Me.chkLocation.Size = New System.Drawing.Size(96, 16)
        Me.chkLocation.TabIndex = 8
        Me.chkLocation.Text = "返回字段位置"
        Me.chkLocation.UseVisualStyleBackColor = True
        '
        'chkVerify
        '
        Me.chkVerify.AutoSize = True
        Me.chkVerify.Location = New System.Drawing.Point(290, 208)
        Me.chkVerify.Name = "chkVerify"
        Me.chkVerify.Size = New System.Drawing.Size(96, 16)
        Me.chkVerify.TabIndex = 9
        Me.chkVerify.Text = "开启验真参数"
        Me.chkVerify.UseVisualStyleBackColor = True
        '
        'chkPreferLocal
        '
        Me.chkPreferLocal.AutoSize = True
        Me.chkPreferLocal.Location = New System.Drawing.Point(15, 238)
        Me.chkPreferLocal.Name = "chkPreferLocal"
        Me.chkPreferLocal.Size = New System.Drawing.Size(144, 16)
        Me.chkPreferLocal.TabIndex = 10
        Me.chkPreferLocal.Text = "优先使用本地文本解析"
        Me.chkPreferLocal.UseVisualStyleBackColor = True
        '
        'chkAutoFallback
        '
        Me.chkAutoFallback.AutoSize = True
        Me.chkAutoFallback.Location = New System.Drawing.Point(200, 238)
        Me.chkAutoFallback.Name = "chkAutoFallback"
        Me.chkAutoFallback.Size = New System.Drawing.Size(180, 16)
        Me.chkAutoFallback.TabIndex = 11
        Me.chkAutoFallback.Text = "本地不足时自动调用在线 OCR"
        Me.chkAutoFallback.UseVisualStyleBackColor = True
        '
        'btnTestOcr
        '
        Me.btnTestOcr.Location = New System.Drawing.Point(15, 260)
        Me.btnTestOcr.Name = "btnTestOcr"
        Me.btnTestOcr.Size = New System.Drawing.Size(150, 24)
        Me.btnTestOcr.TabIndex = 12
        Me.btnTestOcr.Text = "测试 OCR 配置"
        Me.btnTestOcr.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(349, 527)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(90, 25)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "保存"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'grpNaming
        '
        Me.grpNaming.Controls.Add(Me.lblInvoiceTpl)
        Me.grpNaming.Controls.Add(Me.txtInvoiceTemplate)
        Me.grpNaming.Controls.Add(Me.lblTripTpl)
        Me.grpNaming.Controls.Add(Me.txtTripTemplate)
        Me.grpNaming.Controls.Add(Me.lblUnknownTpl)
        Me.grpNaming.Controls.Add(Me.txtUnknownTemplate)
        Me.grpNaming.Controls.Add(Me.btnTplVars)
        Me.grpNaming.Controls.Add(Me.btnTplPreview)
        Me.grpNaming.Location = New System.Drawing.Point(9, 360)
        Me.grpNaming.Name = "grpNaming"
        Me.grpNaming.Size = New System.Drawing.Size(430, 155)
        Me.grpNaming.TabIndex = 6
        Me.grpNaming.TabStop = False
        Me.grpNaming.Text = "最终归档文件命名规则（占位符用 {变量名}）"
        '
        'lblInvoiceTpl
        '
        Me.lblInvoiceTpl.AutoSize = True
        Me.lblInvoiceTpl.Location = New System.Drawing.Point(15, 27)
        Me.lblInvoiceTpl.Name = "lblInvoiceTpl"
        Me.lblInvoiceTpl.Size = New System.Drawing.Size(65, 12)
        Me.lblInvoiceTpl.TabIndex = 0
        Me.lblInvoiceTpl.Text = "滴滴发票："
        '
        'txtInvoiceTemplate
        '
        Me.txtInvoiceTemplate.Location = New System.Drawing.Point(75, 24)
        Me.txtInvoiceTemplate.Name = "txtInvoiceTemplate"
        Me.txtInvoiceTemplate.Size = New System.Drawing.Size(335, 21)
        Me.txtInvoiceTemplate.TabIndex = 0
        '
        'lblTripTpl
        '
        Me.lblTripTpl.AutoSize = True
        Me.lblTripTpl.Location = New System.Drawing.Point(15, 57)
        Me.lblTripTpl.Name = "lblTripTpl"
        Me.lblTripTpl.Size = New System.Drawing.Size(65, 12)
        Me.lblTripTpl.TabIndex = 1
        Me.lblTripTpl.Text = "常规发票："
        '
        'txtTripTemplate
        '
        Me.txtTripTemplate.Location = New System.Drawing.Point(75, 54)
        Me.txtTripTemplate.Name = "txtTripTemplate"
        Me.txtTripTemplate.Size = New System.Drawing.Size(335, 21)
        Me.txtTripTemplate.TabIndex = 1
        '
        'lblUnknownTpl
        '
        Me.lblUnknownTpl.AutoSize = True
        Me.lblUnknownTpl.Location = New System.Drawing.Point(15, 87)
        Me.lblUnknownTpl.Name = "lblUnknownTpl"
        Me.lblUnknownTpl.Size = New System.Drawing.Size(53, 12)
        Me.lblUnknownTpl.TabIndex = 2
        Me.lblUnknownTpl.Text = "未识别："
        '
        'txtUnknownTemplate
        '
        Me.txtUnknownTemplate.Location = New System.Drawing.Point(75, 84)
        Me.txtUnknownTemplate.Name = "txtUnknownTemplate"
        Me.txtUnknownTemplate.Size = New System.Drawing.Size(335, 21)
        Me.txtUnknownTemplate.TabIndex = 2
        '
        'btnTplVars
        '
        Me.btnTplVars.Location = New System.Drawing.Point(249, 119)
        Me.btnTplVars.Name = "btnTplVars"
        Me.btnTplVars.Size = New System.Drawing.Size(75, 22)
        Me.btnTplVars.TabIndex = 3
        Me.btnTplVars.Text = "变量说明"
        Me.btnTplVars.UseVisualStyleBackColor = True
        '
        'btnTplPreview
        '
        Me.btnTplPreview.Location = New System.Drawing.Point(330, 119)
        Me.btnTplPreview.Name = "btnTplPreview"
        Me.btnTplPreview.Size = New System.Drawing.Size(80, 22)
        Me.btnTplPreview.TabIndex = 4
        Me.btnTplPreview.Text = "预览"
        Me.btnTplPreview.UseVisualStyleBackColor = True
        '
        'SettingsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(451, 560)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.grpNaming)
        Me.Controls.Add(Me.grpOcr)
        Me.Controls.Add(Me.lblFolder)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.rdoOnlineParse)
        Me.Controls.Add(Me.rdoLocalParse)
        Me.Controls.Add(Me.btnBrowseFolder)
        Me.Controls.Add(Me.txtFolderPath)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "设置"
        Me.grpOcr.ResumeLayout(False)
        Me.grpOcr.PerformLayout()
        CType(Me.numTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numMaxPages, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNaming.ResumeLayout(False)
        Me.grpNaming.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
End Class
