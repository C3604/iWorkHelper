''' <summary>
''' 归档运行锁令牌：由 <see cref="ArchiveRunGuard.TryAcquire"/> 获取成功后返回。
'''
''' 约定：
'''  - 释放归档运行锁的唯一途径就是 Dispose 本 token；
'''  - 必须用 <c>Using</c> 或 <c>Try...Finally</c> 包裹，保证预检查失败、业务异常、
'''    进度窗口异常关闭等任何路径都能释放；
'''  - Dispose 幂等，多次调用安全。
''' </summary>
Public NotInheritable Class ArchiveRunToken
    Implements IDisposable

    Private _disposed As Boolean = False

    ''' <summary>只允许 <see cref="ArchiveRunGuard"/> 创建。</summary>
    Friend Sub New()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        _disposed = True
        ArchiveRunGuard.Release()
    End Sub

End Class
