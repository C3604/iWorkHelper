''' <summary>
''' 敏感设置读写门面：对 My.Settings 中的 Secret Key 做透明加解密与明文迁移。
''' 只加密 Secret Key 这一敏感项；ApiKey/URL 等普通配置不加密。
''' 解密失败不抛出、不阻止插件加载；日志不输出明文。
''' </summary>
Public Module ProtectedSettingsProvider

    ''' <summary>
    ''' 读取 Secret Key 明文（供 SettingsForm 显示、OcrConfigProvider 使用）。
    ''' 解密失败返回空串，并通过 outDecryptFailed 告知调用方以便提示。
    ''' </summary>
    Public Function GetSecretKey(ByRef outDecryptFailed As Boolean) As String
        outDecryptFailed = False
        Try
            Dim stored As String = My.Settings.BaiduSecretKey
            If String.IsNullOrEmpty(stored) Then
                Return String.Empty
            End If
            Dim ok As Boolean
            Dim plain As String = SecretProtector.TryUnprotect(stored, ok)
            If Not ok Then
                outDecryptFailed = True
                Return String.Empty
            End If
            Return plain
        Catch ex As Exception
            AppLogger.Error("读取 Secret Key 异常（不影响加载）。", ex)
            outDecryptFailed = True
            Return String.Empty
        End Try
    End Function

    ''' <summary>简化重载。</summary>
    Public Function GetSecretKey() As String
        Dim failed As Boolean
        Return GetSecretKey(failed)
    End Function

    ''' <summary>
    ''' 加密保存 Secret Key。调用方负责 My.Settings.Save()。
    ''' </summary>
    Public Sub SetSecretKey(plainText As String)
        Try
            My.Settings.BaiduSecretKey = SecretProtector.Protect(If(plainText, String.Empty).Trim())
        Catch ex As Exception
            AppLogger.Error("保存 Secret Key 异常。", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 启动时迁移：若 Secret Key 为明文遗留值，则加密回写。返回是否发生迁移。
    ''' </summary>
    Public Function MigratePlaintextIfNeeded() As Boolean
        Try
            Dim stored As String = My.Settings.BaiduSecretKey
            If String.IsNullOrEmpty(stored) OrElse SecretProtector.IsProtected(stored) Then
                Return False
            End If
            ' 明文遗留 → 加密回写
            My.Settings.BaiduSecretKey = SecretProtector.Protect(stored)
            My.Settings.Save()
            AppLogger.Info("已将明文 Secret Key 迁移为 DPAPI 加密存储。")
            Return True
        Catch ex As Exception
            AppLogger.Error("Secret Key 明文迁移失败（不影响加载）。", ex)
            Return False
        End Try
    End Function

End Module
