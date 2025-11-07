<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SettingForm
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
        Me.txt_archivepath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_archive = New System.Windows.Forms.Button()
        Me.FolderBrowser_setting = New System.Windows.Forms.FolderBrowserDialog()
        Me.btn_accept = New System.Windows.Forms.Button()
        Me.txt_tmppath = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btn_temp = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'txt_archivepath
        '
        Me.txt_archivepath.Location = New System.Drawing.Point(20, 65)
        Me.txt_archivepath.Name = "txt_archivepath"
        Me.txt_archivepath.Size = New System.Drawing.Size(410, 21)
        Me.txt_archivepath.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(20, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "归档路径"
        '
        'btn_archive
        '
        Me.btn_archive.Location = New System.Drawing.Point(440, 65)
        Me.btn_archive.Name = "btn_archive"
        Me.btn_archive.Size = New System.Drawing.Size(75, 25)
        Me.btn_archive.TabIndex = 2
        Me.btn_archive.Text = "浏览"
        Me.btn_archive.UseVisualStyleBackColor = True
        '
        'btn_accept
        '
        Me.btn_accept.Location = New System.Drawing.Point(22, 390)
        Me.btn_accept.Name = "btn_accept"
        Me.btn_accept.Size = New System.Drawing.Size(75, 23)
        Me.btn_accept.TabIndex = 3
        Me.btn_accept.Text = "保存"
        Me.btn_accept.UseVisualStyleBackColor = True
        '
        'txt_tmppath
        '
        Me.txt_tmppath.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txt_tmppath.Location = New System.Drawing.Point(20, 145)
        Me.txt_tmppath.Name = "txt_tmppath"
        Me.txt_tmppath.Size = New System.Drawing.Size(410, 21)
        Me.txt_tmppath.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(20, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "缓存路径"
        '
        'btn_temp
        '
        Me.btn_temp.Location = New System.Drawing.Point(440, 145)
        Me.btn_temp.Name = "btn_temp"
        Me.btn_temp.Size = New System.Drawing.Size(75, 25)
        Me.btn_temp.TabIndex = 2
        Me.btn_temp.Text = "浏览"
        Me.btn_temp.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(20, 346)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(144, 16)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Text = "合并滴滴发票与行程单"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"开票日期", "发票金额", "销售方名称", "购买方名称"})
        Me.ComboBox1.Location = New System.Drawing.Point(22, 239)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(85, 20)
        Me.ComboBox1.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(20, 207)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 12)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "重命名规则"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(20, 277)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(65, 12)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "文件名示例"
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(515, 179)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        '
        'SettingForm
        '
        Me.AcceptButton = Me.btn_accept
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(546, 461)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.btn_accept)
        Me.Controls.Add(Me.btn_temp)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn_archive)
        Me.Controls.Add(Me.txt_tmppath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_archivepath)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.HelpButton = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "设置"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txt_archivepath As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btn_archive As Windows.Forms.Button
    Friend WithEvents FolderBrowser_setting As Windows.Forms.FolderBrowserDialog
    Friend WithEvents btn_accept As Windows.Forms.Button
    Friend WithEvents txt_tmppath As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents btn_temp As Windows.Forms.Button
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
End Class
