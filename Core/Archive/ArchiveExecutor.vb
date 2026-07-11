Imports System.IO

''' <summary>
''' 归档执行器：把临时 PDF 复制到目标路径。
''' 采用复制（非移动），保留临时文件供失败排查；临时文件清理由工作流统一处理。
''' 单个文件失败仅影响该文件，不抛出以免中断批次。
''' </summary>
Public Class ArchiveExecutor

    ''' <summary>
    ''' 执行复制。目标路径应为已解决冲突的完整路径（ArchivePlanner 产出）。
    ''' 二次防御：若目标已存在（并发/异常），拒绝覆盖并返回失败。
    ''' </summary>
    Public Function Execute(sourceTempPath As String, targetPath As String) As Result
        Try
            If String.IsNullOrWhiteSpace(sourceTempPath) OrElse Not File.Exists(sourceTempPath) Then
                Return Result.Fail("源临时文件不存在：" & PrivacySafeFormatter.MaskPath(sourceTempPath))
            End If

            If File.Exists(targetPath) Then
                ' 绝不覆盖已有文件。
                Return Result.Fail("目标文件已存在，已跳过以避免覆盖：" & PrivacySafeFormatter.MaskPath(targetPath))
            End If

            Dim targetDir As String = Path.GetDirectoryName(targetPath)
            If Not PathHelper.EnsureDirectory(targetDir) Then
                Return Result.Fail("无法创建目标目录：" & PrivacySafeFormatter.MaskPath(targetDir))
            End If

            File.Copy(sourceTempPath, targetPath, overwrite:=False)
            AppLogger.Info("已归档: " & PrivacySafeFormatter.MaskPath(sourceTempPath) &
                           " -> " & PrivacySafeFormatter.MaskPath(targetPath))
            Return Result.Ok(targetPath)

        Catch ex As Exception
            AppLogger.Error("归档复制失败: " & PrivacySafeFormatter.MaskPath(targetPath), ex)
            Return Result.Fail("归档失败：" & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Function

End Class
