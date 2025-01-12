<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MailProgressBarForm
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
        Me.LabelProgressBar = New System.Windows.Forms.Label()
        Me.MailProgressBar = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'LabelProgressBar
        '
        Me.LabelProgressBar.AutoSize = True
        Me.LabelProgressBar.Location = New System.Drawing.Point(12, 49)
        Me.LabelProgressBar.Name = "LabelProgressBar"
        Me.LabelProgressBar.Size = New System.Drawing.Size(97, 15)
        Me.LabelProgressBar.TabIndex = 0
        Me.LabelProgressBar.Text = "正在处理……"
        '
        'MailProgressBar
        '
        Me.MailProgressBar.Location = New System.Drawing.Point(15, 15)
        Me.MailProgressBar.Name = "MailProgressBar"
        Me.MailProgressBar.Size = New System.Drawing.Size(555, 25)
        Me.MailProgressBar.TabIndex = 1
        '
        'MailProgressBarForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(582, 73)
        Me.ControlBox = False
        Me.Controls.Add(Me.MailProgressBar)
        Me.Controls.Add(Me.LabelProgressBar)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MailProgressBarForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "正在处理……"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelProgressBar As Windows.Forms.Label
    Friend WithEvents MailProgressBar As Windows.Forms.ProgressBar
End Class
