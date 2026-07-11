''' <summary>
''' 归档进度上报接口。工作流在各阶段调用；UI 实现负责刷新进度窗口。
''' 实现必须容错：上报异常不得中断归档主流程。
''' </summary>
Public Interface IArchiveProgressReporter

    ''' <summary>上报当前进度。</summary>
    Sub Report(info As ArchiveProgressInfo)

    ''' <summary>是否被请求取消（当前默认始终 False；如实现取消按钮再返回 True）。</summary>
    ReadOnly Property IsCancellationRequested As Boolean

End Interface

''' <summary>空实现：用于离线/无 UI 场景，什么都不做。</summary>
Public Class NullArchiveProgressReporter
    Implements IArchiveProgressReporter

    Public Sub Report(info As ArchiveProgressInfo) Implements IArchiveProgressReporter.Report
        ' no-op
    End Sub

    Public ReadOnly Property IsCancellationRequested As Boolean Implements IArchiveProgressReporter.IsCancellationRequested
        Get
            Return False
        End Get
    End Property
End Class
