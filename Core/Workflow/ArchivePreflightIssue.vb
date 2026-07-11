''' <summary>
''' 归档前预检查发现的单个问题（面向用户）。IsBlocking=True 表示阻止进入批量处理。
''' </summary>
Public Class ArchivePreflightIssue
    Public Property Code As AppErrorCode
    Public Property Severity As ErrorSeverity
    Public Property UserMessage As String
    Public Property Suggestion As String
    Public Property IsBlocking As Boolean

    Public Shared Function FromCode(code As AppErrorCode, blocking As Boolean) As ArchivePreflightIssue
        Dim d As AppError = UserFriendlyMessageProvider.Describe(code)
        Return New ArchivePreflightIssue With {
            .Code = code, .Severity = d.Severity, .UserMessage = d.UserMessage,
            .Suggestion = d.Suggestion, .IsBlocking = blocking}
    End Function
End Class
