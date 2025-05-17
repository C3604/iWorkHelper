Public Class PdfPigProgressBarForm
    ' 窗体中的进度条和标签控件

    ' 初始化窗体
    Public Sub New(totalMails As Integer)
        InitializeComponent()

        ' 设置进度条最大值和初始值
        MailProgressBar.Maximum = 100 ' 0到100的进度
        MailProgressBar.Value = 0
        LabelProgressBar.Text = "等待处理"

        ' 窗体标题显示邮件处理进度
        Me.Text = $"正在处理第 1 封邮件，共 {totalMails} 封邮件"

        ' 记录构造函数调用的日志
        LogManager.WriteLog(LogLevel.INFO, "PdfPigProgressBarForm.New",
                         $"初始化进度条：最大值=100，当前处理第1封邮件，共{totalMails}封邮件")
    End Sub


    ' 更新进度条的值
    Public Sub UpdateMailProgress(currentProgress As Integer, Optional isRecursiveCall As Boolean = False)
        ' 只在非递归调用时记录日志
        If Not isRecursiveCall Then
            LogManager.WriteLog(LogLevel.INFO, "PdfPigProgressBarForm.UpdateMailProgress",
                            $"更新邮件进度: {currentProgress}%")
        End If

        ' 检查是否需要跨线程调用
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of Integer, Boolean)(AddressOf UpdateMailProgress), 
                     currentProgress, True) ' 传递True表示这是递归调用
            Return
        End If
        
        ' 更新进度条的值
        MailProgressBar.Value = currentProgress
    End Sub


    ' 更新当前步骤状态
    Public Sub UpdateStepStatus(stepName As String, Optional isRecursiveCall As Boolean = False)
        ' 只在非递归调用时记录日志
        If Not isRecursiveCall Then
            LogManager.WriteLog(LogLevel.INFO, "PdfPigProgressBarForm.UpdateStepStatus",
                            $"更新步骤: {stepName}")
        End If

        ' 检查是否需要跨线程调用
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of String, Boolean)(AddressOf UpdateStepStatus), 
                     stepName, True) ' 传递True表示这是递归调用
            Return
        End If
        
        ' 更新控件的内容
        LabelProgressBar.Text = stepName
    End Sub


    ' 更新窗体标题
    Public Sub UpdateWindowTitle(currentMail As Integer, totalMails As Integer, Optional isRecursiveCall As Boolean = False)
        ' 只在非递归调用时记录日志
        If Not isRecursiveCall Then
            LogManager.WriteLog(LogLevel.INFO, "PdfPigProgressBarForm.UpdateWindowTitle",
                            $"更新窗口标题: 当前处理邮件 {currentMail}/{totalMails}")
        End If

        ' 检查是否需要跨线程调用
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of Integer, Integer, Boolean)(AddressOf UpdateWindowTitle), 
                     currentMail, totalMails, True) ' 传递True表示这是递归调用
            Return
        End If
        
        ' 更新窗口标题
        Me.Text = $"正在处理第 {currentMail} 封邮件，共 {totalMails} 封邮件"
    End Sub

End Class
