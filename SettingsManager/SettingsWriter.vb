Imports System.Configuration
Imports System.Text.RegularExpressions
Imports iWorkHelper.LogManager

' 线程安全的 My.Settings 写入工具类
' 职责：按照给定的变量名与类型（Boolean/String）将值写入到 My.Settings。
' 特性：
' - 强类型检查：仅接受 Boolean 或 String，且值必须与类型匹配。
' - 变量存在则更新；不存在则创建新设置项。
' - 详细错误信息：返回失败原因（类型不匹配、变量名无效、提供程序缺失、写入失败等）。
' - 线程安全：通过 SyncLock 保证多线程场景下的原子性。
' - 高效：避免不必要的保存操作（仅在值变化或新增时保存）。

Public NotInheritable Class SettingsWriter

    Private Sub New()
        ' 禁止实例化
    End Sub

    ' 全局锁对象，确保多线程环境下的读写安全
    Private Shared ReadOnly _syncRoot As New Object()
    ' 日志实例（共享）供该模块记录输入问题与异常
    Private Shared ReadOnly _logger As New Logger()

    Shared Sub New()
        Try
            _logger.Start()
        Catch
        End Try
    End Sub

    ' 变量名校验：以字母或下划线开头，仅包含字母、数字、下划线
    Private Shared ReadOnly _namePattern As New Regex("^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.Compiled)

    ' 对外主方法：写设置项
    ' 返回：Tuple(Of Boolean, String) => (是否成功, 错误信息；成功时为 Nothing)
    Public Shared Function WriteSetting(ByVal name As String, ByVal typeName As String, ByVal value As Object) As Tuple(Of Boolean, String)
        ' 入参基本校验
        If String.IsNullOrWhiteSpace(name) Then
            Try : _logger.LogWarn("SettingsWriter：变量名无效") : Catch : End Try
            Return Tuple.Create(False, "变量名无效")
        End If

        If Not _namePattern.IsMatch(name) Then
            Try : _logger.LogWarn("SettingsWriter：变量名无效") : Catch : End Try
            Return Tuple.Create(False, "变量名无效")
        End If

        If String.IsNullOrWhiteSpace(typeName) Then
            Try : _logger.LogWarn("SettingsWriter：类型无效") : Catch : End Try
            Return Tuple.Create(False, "类型无效")
        End If

        Dim normalizedType As String = typeName.Trim()
        Dim targetType As Type = Nothing
        Select Case normalizedType
            Case "Boolean"
                targetType = GetType(Boolean)
                If value Is Nothing OrElse Not TypeOf value Is Boolean Then
                    Try : _logger.LogWarn("SettingsWriter：类型不匹配(Boolean)") : Catch : End Try
                    Return Tuple.Create(False, "类型不匹配")
                End If
            Case "String"
                targetType = GetType(String)
                If value Is Nothing OrElse Not TypeOf value Is String Then
                    Try : _logger.LogWarn("SettingsWriter：类型不匹配(String)") : Catch : End Try
                    Return Tuple.Create(False, "类型不匹配")
                End If
            Case Else
                Try : _logger.LogWarn("SettingsWriter：类型无效") : Catch : End Try
                Return Tuple.Create(False, "类型无效")
        End Select

        ' 线程安全写入
        SyncLock _syncRoot
            Try
                ' 确认设置提供程序（增强回退）：优先使用 LocalFileSettingsProvider；若集合为空则创建并添加
                Dim provider As SettingsProvider = My.Settings.Providers("LocalFileSettingsProvider")
                If provider Is Nothing Then
                    ' 尝试从集合中取任一可用提供程序
                    For Each p As SettingsProvider In My.Settings.Providers
                        provider = p
                        Exit For
                    Next
                End If
                If provider Is Nothing Then
                    ' 集合为空，创建本地文件提供程序并加入集合
                    Dim localProvider As New LocalFileSettingsProvider()
                    localProvider.Initialize("LocalFileSettingsProvider", Nothing)
                    My.Settings.Providers.Add(localProvider)
                    provider = localProvider
                End If

                Dim exists As Boolean = (My.Settings.Properties(name) IsNot Nothing)

                If Not exists Then
                    ' 创建新设置项
                    Dim prop As New SettingsProperty(name)
                    prop.PropertyType = targetType
                    prop.Provider = provider
                    ' 设置作用域为用户级，允许保存到 user.config
                    prop.Attributes.Add(GetType(UserScopedSettingAttribute), New UserScopedSettingAttribute())
                    ' 统一使用字符串序列化，框架会正确反序列化类型
                    prop.SerializeAs = SettingsSerializeAs.String
                    ' 与序列化格式一致的默认值
                    If targetType Is GetType(Boolean) Then
                        prop.DefaultValue = "False"
                    Else
                        prop.DefaultValue = String.Empty
                    End If

                    My.Settings.Properties.Add(prop)
                End If

                ' 仅在数值变化时写入与保存，提高性能
                Dim shouldSave As Boolean = Not exists
                If exists Then
                    Dim current As Object = My.Settings(name)
                    If targetType Is GetType(Boolean) Then
                        Dim curBool As Boolean = False
                        If current IsNot Nothing AndAlso TypeOf current Is Boolean Then
                            curBool = DirectCast(current, Boolean)
                        End If
                        Dim newBool As Boolean = DirectCast(value, Boolean)
                        shouldSave = (curBool <> newBool)
                    Else
                        Dim curStr As String = If(TryCast(current, String), String.Empty)
                        Dim newStr As String = DirectCast(value, String)
                        shouldSave = (curStr <> newStr)
                    End If
                End If

                ' 设置值
                My.Settings(name) = value

                If shouldSave Then
                    My.Settings.Save()
                End If

                ' 若写入的是 tmppath，立即刷新 Logger 的内部路径到新值
                Try
                    If name IsNot Nothing AndAlso name.ToLowerInvariant() = "tmppath" Then
                        _logger.RefreshPathIfChanged(DirectCast(value, String))
                    End If
                Catch
                End Try

                Try : _logger.LogInfo($"SettingsWriter：写入成功 {name}") : Catch : End Try
                Return Tuple.Create(True, CType(Nothing, String))
            Catch ex As Exception
                ' 捕获所有异常并返回错误原因
                Try : _logger.LogError($"SettingsWriter：写入失败：{ex.Message}") : Catch : End Try
                Return Tuple.Create(False, $"写入失败：{ex.Message}")
            End Try
        End SyncLock
    End Function

End Class