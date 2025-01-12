Partial Class MainRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabMain = Me.Factory.CreateRibbonTab
        Me.GroupInvoice = Me.Factory.CreateRibbonGroup
        Me.ButtonArchive = Me.Factory.CreateRibbonButton
        Me.ButtonSetting = Me.Factory.CreateRibbonButton
        Me.ButtonHelp = Me.Factory.CreateRibbonButton
        Me.TabMain.SuspendLayout()
        Me.GroupInvoice.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabMain
        '
        Me.TabMain.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabMain.Groups.Add(Me.GroupInvoice)
        Me.TabMain.Label = "工具"
        Me.TabMain.Name = "TabMain"
        '
        'GroupInvoice
        '
        Me.GroupInvoice.Items.Add(Me.ButtonArchive)
        Me.GroupInvoice.Items.Add(Me.ButtonSetting)
        Me.GroupInvoice.Items.Add(Me.ButtonHelp)
        Me.GroupInvoice.Label = "发票工具"
        Me.GroupInvoice.Name = "GroupInvoice"
        '
        'ButtonArchive
        '
        Me.ButtonArchive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonArchive.Image = Global.XYOutlookPlugin.My.Resources.Resources.Archive
        Me.ButtonArchive.Label = "归档"
        Me.ButtonArchive.Name = "ButtonArchive"
        Me.ButtonArchive.ShowImage = True
        '
        'ButtonSetting
        '
        Me.ButtonSetting.Image = Global.XYOutlookPlugin.My.Resources.Resources.Setting
        Me.ButtonSetting.Label = "设置"
        Me.ButtonSetting.Name = "ButtonSetting"
        Me.ButtonSetting.ShowImage = True
        '
        'ButtonHelp
        '
        Me.ButtonHelp.Image = Global.XYOutlookPlugin.My.Resources.Resources.Help
        Me.ButtonHelp.Label = "帮助"
        Me.ButtonHelp.Name = "ButtonHelp"
        Me.ButtonHelp.ShowImage = True
        '
        'MainRibbon
        '
        Me.Name = "MainRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.TabMain)
        Me.TabMain.ResumeLayout(False)
        Me.TabMain.PerformLayout()
        Me.GroupInvoice.ResumeLayout(False)
        Me.GroupInvoice.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabMain As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupInvoice As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonArchive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSetting As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonHelp As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property MainRibbon() As MainRibbon
        Get
            Return Me.GetRibbon(Of MainRibbon)()
        End Get
    End Property
End Class
