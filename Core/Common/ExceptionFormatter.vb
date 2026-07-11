Imports System.Text

''' <summary>
''' 异常格式化工具：把异常转换为对用户/日志友好的可读文本，
''' 包含内层异常链。用于把错误反馈到 UI，而不是只弹一个空洞的 MessageBox。
''' </summary>
Public Module ExceptionFormatter

    ''' <summary>
    ''' 面向用户的简短描述：仅顶层与内层异常的消息，不含堆栈。
    ''' </summary>
    Public Function ToUserMessage(ex As Exception) As String
        If ex Is Nothing Then
            Return "未知错误。"
        End If

        Dim sb As New StringBuilder()
        Dim current As Exception = ex
        Dim depth As Integer = 0
        While current IsNot Nothing AndAlso depth < 5
            If depth > 0 Then
                sb.Append(" -> ")
            End If
            sb.Append(current.Message)
            current = current.InnerException
            depth += 1
        End While
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 面向日志的完整描述：类型、消息、堆栈及内层异常链。
    ''' </summary>
    Public Function ToLogText(ex As Exception) As String
        If ex Is Nothing Then
            Return "(无异常对象)"
        End If

        Dim sb As New StringBuilder()
        Dim current As Exception = ex
        Dim depth As Integer = 0
        While current IsNot Nothing AndAlso depth < 8
            If depth > 0 Then
                sb.AppendLine("--- 内层异常 ---")
            End If
            sb.AppendLine("类型: " & current.GetType().FullName)
            sb.AppendLine("消息: " & current.Message)
            If Not String.IsNullOrEmpty(current.StackTrace) Then
                sb.AppendLine("堆栈:")
                sb.AppendLine(current.StackTrace)
            End If
            current = current.InnerException
            depth += 1
        End While
        Return sb.ToString()
    End Function

End Module
