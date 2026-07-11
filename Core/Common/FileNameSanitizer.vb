Imports System.IO
Imports System.Text

''' <summary>
''' 文件名清理工具：去除 Windows 非法字符、处理保留名、限制长度。
''' 用于把从发票中提取的字段安全地拼接为文件名。
''' </summary>
Public Module FileNameSanitizer

    ''' <summary>Windows 文件名最大安全长度（保守值，避免路径整体超限）。</summary>
    Private Const MaxBaseNameLength As Integer = 120

    ''' <summary>Windows 保留设备名，不能作为文件名主体。</summary>
    Private ReadOnly ReservedNames As String() = New String() {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"}

    ''' <summary>
    ''' 清理文件名主体（不含扩展名）。非法字符替换为下划线；
    ''' 空白折叠；保留名加前缀；超长截断。
    ''' </summary>
    ''' <param name="rawName">原始名称片段。</param>
    ''' <param name="fallback">清理后为空时使用的回退名。</param>
    Public Function SanitizeBaseName(rawName As String, Optional fallback As String = "未命名") As String
        If String.IsNullOrWhiteSpace(rawName) Then
            Return fallback
        End If

        Dim invalid As Char() = Path.GetInvalidFileNameChars()
        Dim sb As New StringBuilder(rawName.Length)
        For Each c As Char In rawName
            If Array.IndexOf(invalid, c) >= 0 Then
                sb.Append("_"c)
            ElseIf c = ControlChars.Tab OrElse c = ControlChars.Lf OrElse c = ControlChars.Cr Then
                sb.Append(" "c)
            Else
                sb.Append(c)
            End If
        Next

        Dim cleaned As String = sb.ToString()

        ' 折叠连续空白为单个空格，并去除首尾空白与点（Windows 不允许结尾为点/空格）。
        While cleaned.IndexOf("  ") >= 0
            cleaned = cleaned.Replace("  ", " ")
        End While
        cleaned = cleaned.Trim().Trim("."c).Trim()

        If cleaned.Length = 0 Then
            Return fallback
        End If

        ' 保留名处理。
        Dim upper As String = cleaned.ToUpperInvariant()
        For Each reserved As String In ReservedNames
            If upper = reserved Then
                cleaned = "_" & cleaned
                Exit For
            End If
        Next

        ' 超长截断。
        If cleaned.Length > MaxBaseNameLength Then
            cleaned = cleaned.Substring(0, MaxBaseNameLength).Trim()
        End If

        If cleaned.Length = 0 Then
            Return fallback
        End If

        Return cleaned
    End Function

    ''' <summary>
    ''' 生成完整文件名（自动补全扩展名，默认 .pdf）。
    ''' </summary>
    Public Function BuildFileName(baseName As String, Optional extension As String = ".pdf", Optional fallback As String = "未命名") As String
        Dim safeBase As String = SanitizeBaseName(baseName, fallback)
        Dim ext As String = If(extension, String.Empty)
        If ext.Length > 0 AndAlso Not ext.StartsWith(".") Then
            ext = "." & ext
        End If
        Return safeBase & ext
    End Function

End Module
