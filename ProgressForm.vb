Imports System.Windows.Forms

''' <summary>
''' 归档进度窗口。以“邮件”为单位显示总数/已处理/当前邮件/当前阶段/百分比。
''' 采用 Application.DoEvents 在关键阶段刷新（VSTO 归档在 UI 线程执行，避免后台线程访问 Outlook COM）。
''' 关闭窗口默认不取消任务（无取消按钮，见 SECRET/PROCESS 文档说明）。
''' </summary>
Public Class ProgressForm

    ''' <summary>用进度信息刷新界面，并让 UI 有机会重绘。</summary>
    Public Sub UpdateProgress(info As ArchiveProgressInfo)
        If info Is Nothing Then Return
        Try
            lblTotal.Text = String.Format("共 {0} 封邮件，已处理 {1} 封", info.TotalEmails, info.ProcessedEmails)
            Dim subject As String = If(String.IsNullOrEmpty(info.CurrentEmailSubject), "-", info.CurrentEmailSubject)
            If subject.Length > 42 Then subject = subject.Substring(0, 42) & "…"
            lblCurrent.Text = String.Format("当前：第 {0}/{1} 封  {2}", info.CurrentEmailIndex, info.TotalEmails, subject)
            lblStage.Text = "阶段：" & info.StageText
            lblNote.Text = If(info.Note, "")
            progressBar1.Value = info.Percent
            Application.DoEvents()
        Catch
            ' 进度刷新失败不影响归档主流程。
        End Try
    End Sub

    Private Sub lblStage_Click(sender As Object, e As EventArgs) Handles lblStage.Click

    End Sub
End Class
