<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressForm
    Inherits System.Windows.Forms.Form

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

    Private components As System.ComponentModel.IContainer

    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents lblCurrent As System.Windows.Forms.Label
    Friend WithEvents lblStage As System.Windows.Forms.Label
    Friend WithEvents lblNote As System.Windows.Forms.Label
    Friend WithEvents progressBar1 As System.Windows.Forms.ProgressBar

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblCurrent = New System.Windows.Forms.Label()
        Me.lblStage = New System.Windows.Forms.Label()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.progressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Location = New System.Drawing.Point(12, 12)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(125, 12)
        Me.lblTitle.TabIndex = 5
        Me.lblTitle.Text = "正在归档，请稍候……"
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(189, 12)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(149, 12)
        Me.lblTotal.TabIndex = 4
        Me.lblTotal.Text = "共 0 封邮件，已处理 0 封"
        '
        'lblCurrent
        '
        Me.lblCurrent.AutoSize = True
        Me.lblCurrent.Location = New System.Drawing.Point(12, 36)
        Me.lblCurrent.Name = "lblCurrent"
        Me.lblCurrent.Size = New System.Drawing.Size(47, 12)
        Me.lblCurrent.TabIndex = 3
        Me.lblCurrent.Text = "当前：-"
        '
        'lblStage
        '
        Me.lblStage.AutoSize = True
        Me.lblStage.Location = New System.Drawing.Point(12, 62)
        Me.lblStage.Name = "lblStage"
        Me.lblStage.Size = New System.Drawing.Size(47, 12)
        Me.lblStage.TabIndex = 2
        Me.lblStage.Text = "阶段：-"
        '
        'lblNote
        '
        Me.lblNote.AutoSize = True
        Me.lblNote.ForeColor = System.Drawing.Color.FromArgb(CType(CType(180, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblNote.Location = New System.Drawing.Point(12, 135)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(0, 12)
        Me.lblNote.TabIndex = 1
        '
        'progressBar1
        '
        Me.progressBar1.Location = New System.Drawing.Point(12, 88)
        Me.progressBar1.Name = "progressBar1"
        Me.progressBar1.Size = New System.Drawing.Size(400, 22)
        Me.progressBar1.TabIndex = 0
        '
        'ProgressForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(424, 128)
        Me.ControlBox = False
        Me.Controls.Add(Me.progressBar1)
        Me.Controls.Add(Me.lblNote)
        Me.Controls.Add(Me.lblStage)
        Me.Controls.Add(Me.lblCurrent)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.lblTitle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ProgressForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "工作助手 - 归档进度"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
End Class
