Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports iWorkHelper.LogManager

<TestClass()> Public Class LogManagerTests

    Private Class DummyNotifier
        Implements ILogNotifier

        Public LastLevel As LogLevel = LogLevel.Info
        Public LastMessage As String = Nothing

        Public Sub Notify(level As LogLevel, message As String) Implements ILogNotifier.Notify
            LastLevel = level
            LastMessage = message
        End Sub
    End Class

    Private Function NewTempRoot() As String
        Dim root = Path.Combine(Path.GetTempPath(), "iWorkHelper_Test_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Return root
    End Function

    <TestMethod()> Public Sub Logger_Creates_Dir_And_File_UTF8()
        Dim root = NewTempRoot()
        Dim notifier As New DummyNotifier()
        Using logger As New Logger(root, notifier, flushIntervalMs:=200, batchSize:=16)
            logger.Start()
            logger.LogInfo("测试内容-中文-UTF8")
            Task.Delay(500).Wait()
            logger.Stop()
        End Using

        Dim logDir = Path.Combine(root, "log")
        Dim logFile = Path.Combine(logDir, "iWorkhelper.log")
        Assert.IsTrue(Directory.Exists(logDir))
        Assert.IsTrue(File.Exists(logFile))

        Dim bytes = File.ReadAllBytes(logFile)
        ' 只要能用 UTF8 正常解码且包含中文则认为编码正确
        Dim text = Encoding.UTF8.GetString(bytes)
        StringAssert.Contains(text, "测试内容-中文-UTF8")
        StringAssert.Contains(text, "[Info]")
    End Sub

    <TestMethod()> Public Sub Logger_Warn_Error_Popup_Notified_And_Written()
        Dim root = NewTempRoot()
        Dim notifier As New DummyNotifier()
        Using logger As New Logger(root, notifier, flushIntervalMs:=200, batchSize:=4)
            logger.Start()
            logger.LogWarn("需要关注的情况")
            Assert.AreEqual(LogLevel.Warn, notifier.LastLevel)
            StringAssert.Contains(notifier.LastMessage, "需要关注的情况")

            logger.LogError("严重错误发生")
            Assert.AreEqual(LogLevel.ErrorLevel, notifier.LastLevel)
            StringAssert.Contains(notifier.LastMessage, "严重错误发生")

            Task.Delay(500).Wait()
            logger.Stop()
        End Using

        Dim logText = File.ReadAllText(Path.Combine(root, "log", "iWorkhelper.log"), Encoding.UTF8)
        StringAssert.Contains(logText, "[Warn]")
        StringAssert.Contains(logText, "需要关注的情况")
        StringAssert.Contains(logText, "[Error]")
        StringAssert.Contains(logText, "严重错误发生")
    End Sub

    <TestMethod()> Public Sub Logger_Rotate_When_Exceeds_Max_Size()
        Dim root = NewTempRoot()
        Dim notifier As New DummyNotifier()
        Using logger As New Logger(root, notifier, flushIntervalMs:=100, batchSize:=32, maxFileSizeBytes:=100 * 1024)
            logger.Start()
            Dim bigMsg = New String("A"c, 20 * 1024)
            For i = 0 To 40
                logger.LogInfo($"line {i} {bigMsg}")
            Next
            Task.Delay(1000).Wait()
            logger.Stop()
        End Using

        Dim logDir = Path.Combine(root, "log")
        Dim rotated = Directory.GetFiles(logDir, "iWorkhelper_*.log")
        Assert.IsTrue(rotated.Length >= 1)
        Assert.IsTrue(File.Exists(Path.Combine(logDir, "iWorkhelper.log")))
    End Sub

    <TestMethod()> Public Sub Logger_Fallback_When_IO_Exception()
        Dim root = NewTempRoot()
        Dim logDir = Path.Combine(root, "log")
        Directory.CreateDirectory(logDir)
        Dim logFile = Path.Combine(logDir, "iWorkhelper.log")
        File.WriteAllText(logFile, "", Encoding.UTF8)
        ' 设置只读，触发写入失败
        Dim attr = File.GetAttributes(logFile)
        File.SetAttributes(logFile, attr Or FileAttributes.ReadOnly)

        Dim notifier As New DummyNotifier()
        Using logger As New Logger(root, notifier, flushIntervalMs:=200, batchSize:=8)
            logger.Start()
            logger.LogError("触发IO异常写入失败")
            Task.Delay(600).Wait()
            logger.Stop()
        End Using

        ' 恢复属性以便清理
        File.SetAttributes(logFile, FileAttributes.Normal)

        Dim fbDir = Path.Combine(Path.GetTempPath(), "iWorkHelper", "log")
        Dim fbFile = Path.Combine(fbDir, "iWorkhelper_fallback.log")
        Assert.IsTrue(File.Exists(fbFile))
        Dim fbText = File.ReadAllText(fbFile, Encoding.UTF8)
        StringAssert.Contains(fbText, "IO异常")
    End Sub

End Class