''' <summary>
''' 编译期功能开关。根据编译配置选择，控制内网版和外网版的功能差异。
'''
''' 编译配置映射：
'''   Release-Intranet -> INTRANET_BUILD
'''   Release-Internet -> INTERNET_BUILD
'''
''' 默认安全策略：如果未定义 INTERNET_BUILD，则在线解析被禁用。
''' </summary>
Public NotInheritable Class BuildFeatures

    Private Sub New()
    End Sub

#If INTERNET_BUILD Then
    ''' <summary>在线解析是否启用（外网版本）。</summary>
    Public Const OnlineParserEnabled As Boolean = True

    ''' <summary>版本类型显示文本（外网版本）。</summary>
    Public Const EditionName As String = ""  ' 外网版本不显示版本类型

    ''' <summary>版本完整显示格式（外网版本）。</summary>
    Public Const EditionDisplay As String = ""

#Else
    ''' <summary>在线解析是否启用（内网版本，默认）。</summary>
    Public Const OnlineParserEnabled As Boolean = False

    ''' <summary>版本类型显示文本（内网版本）。</summary>
    Public Const EditionName As String = "Intranet"

    ''' <summary>版本完整显示格式（内网版本）。</summary>
    Public Const EditionDisplay As String = "（内网版）"

#End If

End Class
