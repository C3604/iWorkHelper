Imports System.Collections.Generic
Imports System.Text

''' <summary>
''' 归档前预检查结果：问题清单（可一次列出多个）+ 是否存在阻断项。
''' </summary>
Public Class ArchivePreflightResult

    Public Sub New()
        Issues = New List(Of ArchivePreflightIssue)()
    End Sub

    Public Property Issues As List(Of ArchivePreflightIssue)

    Public Sub Add(issue As ArchivePreflightIssue)
        If issue IsNot Nothing Then Issues.Add(issue)
    End Sub

    Public Sub AddCode(code As AppErrorCode, blocking As Boolean)
        Issues.Add(ArchivePreflightIssue.FromCode(code, blocking))
    End Sub

    ''' <summary>是否存在阻断项（存在则不进入批量处理）。</summary>
    Public ReadOnly Property HasBlocking As Boolean
        Get
            For Each i As ArchivePreflightIssue In Issues
                If i.IsBlocking Then Return True
            Next
            Return False
        End Get
    End Property

    Public ReadOnly Property HasAny As Boolean
        Get
            Return Issues.Count > 0
        End Get
    End Property

    ''' <summary>面向用户的合并文案（一次列出全部问题 + 建议）。</summary>
    Public Function BuildUserText() As String
        If Issues.Count = 0 Then Return "预检查通过。"
        Dim sb As New StringBuilder()
        If HasBlocking Then
            sb.AppendLine("无法开始归档，请先解决以下问题：")
        Else
            sb.AppendLine("以下提示不影响归档，可继续：")
        End If
        sb.AppendLine()
        Dim n As Integer = 0
        For Each i As ArchivePreflightIssue In Issues
            n += 1
            sb.Append(n.ToString()).Append(". ").Append(i.UserMessage)
            If Not String.IsNullOrEmpty(i.Suggestion) Then
                sb.Append("（").Append(i.Suggestion).Append("）")
            End If
            sb.AppendLine()
        Next
        Return sb.ToString()
    End Function

End Class
