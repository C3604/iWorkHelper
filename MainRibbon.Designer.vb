Partial Class iWorkHelperRibbon
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
        Me.MainTab = Me.Factory.CreateRibbonTab
        Me.Group_InvoiceHelper = Me.Factory.CreateRibbonGroup
        Me.btn_Archive = Me.Factory.CreateRibbonButton
        Me.btn_Settings = Me.Factory.CreateRibbonButton
        Me.MainTab.SuspendLayout()
        Me.Group_InvoiceHelper.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainTab
        '
        Me.MainTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.MainTab.Groups.Add(Me.Group_InvoiceHelper)
        Me.MainTab.Label = "工作助手"
        Me.MainTab.Name = "MainTab"
        '
        'Group_InvoiceHelper
        '
        Me.Group_InvoiceHelper.Items.Add(Me.btn_Archive)
        Me.Group_InvoiceHelper.Items.Add(Me.btn_Settings)
        Me.Group_InvoiceHelper.Label = "发票助手"
        Me.Group_InvoiceHelper.Name = "Group_InvoiceHelper"
        '
        'btn_Archive
        '
        Me.btn_Archive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_Archive.Image = Global.iWorkHelper.My.Resources.Resources.Archive
        Me.btn_Archive.Label = "发票归档"
        Me.btn_Archive.Name = "btn_Archive"
        Me.btn_Archive.ShowImage = True
        '
        'btn_Settings
        '
        Me.btn_Settings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_Settings.Image = Global.iWorkHelper.My.Resources.Resources.Settings
        Me.btn_Settings.Label = "设置"
        Me.btn_Settings.Name = "btn_Settings"
        Me.btn_Settings.ShowImage = True
        '
        'iWorkHelperRibbon
        '
        Me.Name = "iWorkHelperRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.MainTab)
        Me.MainTab.ResumeLayout(False)
        Me.MainTab.PerformLayout()
        Me.Group_InvoiceHelper.ResumeLayout(False)
        Me.Group_InvoiceHelper.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents MainTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group_InvoiceHelper As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_Archive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_Settings As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property MainRibbon() As iWorkHelperRibbon
        Get
            Return Me.GetRibbon(Of iWorkHelperRibbon)()
        End Get
    End Property
End Class
