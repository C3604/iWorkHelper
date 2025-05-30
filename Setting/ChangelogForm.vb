Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Public Partial Class ChangelogForm
    Private Class VersionInfo
        Public Property Version As String
        Public Property Date As String
        Public Property Changes As List(Of String)
    End Class

    Private Class ChangelogData
        Public Property Versions As List(Of VersionInfo)
    End Class

    ' 构造函数，接收JSON字符串作为参数
    Public Sub New(jsonData As String)
        InitializeComponent()
        
        ' 解析JSON并填充控件
        Try
            PopulateChangelog(jsonData)
        Catch ex As Exception
            MessageBox.Show("解析更新日志时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' 解析JSON并填充日志信息
    Private Sub PopulateChangelog(jsonData As String)
        If String.IsNullOrEmpty(jsonData) Then
            txtChangelog.Text = "无法获取更新日志信息"
            Return
        End If
        
        ' 手动解析JSON (简化处理)
        Try
            Dim changelog As New StringBuilder()
            Dim data As New ChangelogData()
            
            ' 解析JSON得到版本列表
            data.Versions = ParseVersionsFromJson(jsonData)
            
            ' 生成格式化的更新日志文本
            changelog.AppendLine("软件更新日志")
            changelog.AppendLine("====================")
            changelog.AppendLine()
            
            For Each version As VersionInfo In data.Versions
                changelog.AppendLine($"版本 {version.Version} - {version.Date}")
                changelog.AppendLine("--------------------")
                
                For Each change As String In version.Changes
                    changelog.AppendLine($"• {change}")
                Next
                
                changelog.AppendLine()
            Next
            
            txtChangelog.Text = changelog.ToString()
        Catch ex As Exception
            txtChangelog.Text = "解析更新日志时出错: " & ex.Message
        End Try
    End Sub
    
    ' 简单的JSON解析方法
    Private Function ParseVersionsFromJson(jsonString As String) As List(Of VersionInfo)
        Dim versions As New List(Of VersionInfo)()
        
        Try
            ' 查找versions数组
            Dim versionsStart As Integer = jsonString.IndexOf("""versions""")
            If versionsStart < 0 Then Return versions
            
            ' 提取版本数组内容
            Dim arrayStart As Integer = jsonString.IndexOf("[", versionsStart)
            Dim arrayEnd As Integer = FindMatchingBracket(jsonString, arrayStart)
            
            If arrayStart < 0 OrElse arrayEnd < 0 Then Return versions
            
            Dim versionsJson As String = jsonString.Substring(arrayStart + 1, arrayEnd - arrayStart - 1)
            
            ' 分割各个版本对象
            Dim versionObjects As List(Of String) = SplitJsonObjects(versionsJson)
            
            ' 解析每个版本对象
            For Each versionJson As String In versionObjects
                Dim version As New VersionInfo()
                
                ' 提取版本号
                version.Version = ExtractJsonValue(versionJson, "version")
                
                ' 提取日期
                version.Date = ExtractJsonValue(versionJson, "date")
                
                ' 提取更改内容
                version.Changes = New List(Of String)()
                Dim changesStart As Integer = versionJson.IndexOf("""changes""")
                If changesStart >= 0 Then
                    Dim changesArrayStart As Integer = versionJson.IndexOf("[", changesStart)
                    Dim changesArrayEnd As Integer = FindMatchingBracket(versionJson, changesArrayStart)
                    
                    If changesArrayStart >= 0 AndAlso changesArrayEnd >= 0 Then
                        Dim changesJson As String = versionJson.Substring(changesArrayStart + 1, changesArrayEnd - changesArrayStart - 1)
                        Dim changeItems As List(Of String) = SplitJsonArray(changesJson)
                        
                        For Each changeItem As String In changeItems
                            ' 清理引号和空格
                            Dim change As String = changeItem.Trim().Trim(""""c)
                            If Not String.IsNullOrEmpty(change) Then
                                version.Changes.Add(change)
                            End If
                        Next
                    End If
                End If
                
                versions.Add(version)
            Next
        Catch ex As Exception
            ' 处理解析错误
        End Try
        
        Return versions
    End Function
    
    ' 从JSON对象中提取指定键的值
    Private Function ExtractJsonValue(jsonObject As String, key As String) As String
        Dim keyPattern As String = """" & key & """\s*:"
        Dim keyIndex As Integer = jsonObject.IndexOf(keyPattern)
        
        If keyIndex < 0 Then Return ""
        
        Dim valueStart As Integer = jsonObject.IndexOf(":", keyIndex) + 1
        Dim valueEnd As Integer = -1
        
        ' 跳过空白字符
        While valueStart < jsonObject.Length AndAlso Char.IsWhiteSpace(jsonObject(valueStart))
            valueStart += 1
        End While
        
        If valueStart >= jsonObject.Length Then Return ""
        
        ' 检查值的类型
        If jsonObject(valueStart) = """"c Then
            ' 字符串值
            valueStart += 1
            valueEnd = jsonObject.IndexOf("""", valueStart)
            If valueEnd < 0 Then Return ""
            
            Return jsonObject.Substring(valueStart, valueEnd - valueStart)
        ElseIf jsonObject(valueStart) = "["c OrElse jsonObject(valueStart) = "{"c Then
            ' 数组或对象，找到匹配的括号
            valueEnd = FindMatchingBracket(jsonObject, valueStart)
            If valueEnd < 0 Then Return ""
            
            Return jsonObject.Substring(valueStart, valueEnd - valueStart + 1)
        Else
            ' 数字、布尔值等
            valueEnd = jsonObject.IndexOf(",", valueStart)
            If valueEnd < 0 Then
                valueEnd = jsonObject.IndexOf("}", valueStart)
            End If
            
            If valueEnd < 0 Then Return ""
            
            Return jsonObject.Substring(valueStart, valueEnd - valueStart).Trim()
        End If
    End Function
    
    ' 查找匹配的括号
    Private Function FindMatchingBracket(text As String, openIndex As Integer) As Integer
        If openIndex < 0 OrElse openIndex >= text.Length Then Return -1
        
        Dim openChar As Char = text(openIndex)
        Dim closeChar As Char
        
        If openChar = "["c Then
            closeChar = "]"c
        ElseIf openChar = "{"c Then
            closeChar = "}"c
        Else
            Return -1
        End If
        
        Dim depth As Integer = 1
        Dim i As Integer = openIndex + 1
        
        While i < text.Length AndAlso depth > 0
            If text(i) = openChar Then
                depth += 1
            ElseIf text(i) = closeChar Then
                depth -= 1
            End If
            
            i += 1
        End While
        
        If depth = 0 Then
            Return i - 1
        Else
            Return -1
        End If
    End Function
    
    ' 分割JSON对象
    Private Function SplitJsonObjects(json As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim i As Integer = 0
        
        While i < json.Length
            ' 跳过空白字符
            While i < json.Length AndAlso Char.IsWhiteSpace(json(i))
                i += 1
            End While
            
            If i >= json.Length Then Exit While
            
            If json(i) = "{"c Then
                Dim objEnd As Integer = FindMatchingBracket(json, i)
                If objEnd < 0 Then Exit While
                
                result.Add(json.Substring(i, objEnd - i + 1))
                i = objEnd + 1
            Else
                i += 1
            End If
            
            ' 跳过逗号和空白字符
            While i < json.Length AndAlso (json(i) = ","c OrElse Char.IsWhiteSpace(json(i)))
                i += 1
            End While
        End While
        
        Return result
    End Function
    
    ' 分割JSON数组元素
    Private Function SplitJsonArray(json As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim i As Integer = 0
        Dim itemStart As Integer = 0
        Dim inQuotes As Boolean = False
        Dim depth As Integer = 0
        
        While i < json.Length
            Dim c As Char = json(i)
            
            If c = """"c AndAlso (i = 0 OrElse json(i - 1) <> "\"c) Then
                inQuotes = Not inQuotes
            ElseIf Not inQuotes Then
                If c = "{"c OrElse c = "["c Then
                    depth += 1
                ElseIf c = "}"c OrElse c = "]"c Then
                    depth -= 1
                ElseIf c = ","c AndAlso depth = 0 Then
                    result.Add(json.Substring(itemStart, i - itemStart).Trim())
                    itemStart = i + 1
                End If
            End If
            
            i += 1
        End While
        
        ' 添加最后一个元素
        If itemStart < json.Length Then
            result.Add(json.Substring(itemStart).Trim())
        End If
        
        Return result
    End Function

    ' 关闭按钮点击事件
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class 