Imports System.Windows.Forms

Namespace LogManager
    Public Interface ILogNotifier
        Sub Notify(level As LogLevel, message As String)
    End Interface
End Namespace