''' <summary>错误严重级别（弹窗图标 / 日志级别映射）。</summary>
Public Enum ErrorSeverity
    ''' <summary>提示性信息（如跳过无附件邮件）。</summary>
    Info = 0
    ''' <summary>警告（可继续，但结果不理想）。</summary>
    Warning = 1
    ''' <summary>错误（该项失败）。</summary>
    [Error] = 2
    ''' <summary>严重（阻断整个流程，如归档目录不可用）。</summary>
    Critical = 3
End Enum
