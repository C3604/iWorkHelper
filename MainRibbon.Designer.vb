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
        Me.TabWorkHelper = Me.Factory.CreateRibbonTab
        Me.GroupInvoiceFiling = Me.Factory.CreateRibbonGroup
        Me.ButtonArchive = Me.Factory.CreateRibbonButton
        Me.ButtonSettings = Me.Factory.CreateRibbonButton
        Me.TabWorkHelper.SuspendLayout()
        Me.GroupInvoiceFiling.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabWorkHelper
        '
        Me.TabWorkHelper.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabWorkHelper.Groups.Add(Me.GroupInvoiceFiling)
        Me.TabWorkHelper.Label = "工作助手"
        Me.TabWorkHelper.Name = "TabWorkHelper"
        '
        'GroupInvoiceFiling
        '
        Me.GroupInvoiceFiling.Items.Add(Me.ButtonArchive)
        Me.GroupInvoiceFiling.Items.Add(Me.ButtonSettings)
        Me.GroupInvoiceFiling.Label = "发票归档"
        Me.GroupInvoiceFiling.Name = "GroupInvoiceFiling"
        '
        'ButtonArchive
        '
        Me.ButtonArchive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonArchive.Image = Global.iWorkhelper.My.Resources.Resources.Archive
        Me.ButtonArchive.Label = "归档"
        Me.ButtonArchive.Name = "ButtonArchive"
        Me.ButtonArchive.ShowImage = True
        '
        'ButtonSettings
        '
        Me.ButtonSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonSettings.Image = Global.iWorkhelper.My.Resources.Resources.Setting
        Me.ButtonSettings.Label = "设置"
        Me.ButtonSettings.Name = "ButtonSettings"
        Me.ButtonSettings.ShowImage = True
        '
        'MainRibbon
        '
        Me.Name = "MainRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.TabWorkHelper)
        Me.TabWorkHelper.ResumeLayout(False)
        Me.TabWorkHelper.PerformLayout()
        Me.GroupInvoiceFiling.ResumeLayout(False)
        Me.GroupInvoiceFiling.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabWorkHelper As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupInvoiceFiling As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonArchive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property MainRibbon() As MainRibbon
        Get
            Return Me.GetRibbon(Of MainRibbon)()
        End Get
    End Property
End Class
