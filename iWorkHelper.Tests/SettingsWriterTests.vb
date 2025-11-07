Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Threading
Imports System.Threading.Tasks

<TestClass()> Public Class SettingsWriterTests

    Private Sub RemoveSetting(name As String)
        SyncLock GetType(SettingsWriter)
            Try
                If My.Settings.Properties(name) IsNot Nothing Then
                    My.Settings.Properties.Remove(name)
                    My.Settings.Save()
                End If
            Catch
                ' 清理失败忽略，不影响测试准确性（重新创建会覆盖）
            End Try
        End SyncLock
    End Sub

    <TestInitialize()> Public Sub Init()
        ' 每个测试前清理可能残留的测试项
        RemoveSetting("Test_String_New")
        RemoveSetting("Test_Boolean_New")
        RemoveSetting("Test_Update_String")
        RemoveSetting("Test_ThreadSafe_Bool")
    End Sub

    <TestMethod()> Public Sub Reject_Invalid_Name_Empty()
        Dim result = SettingsWriter.WriteSetting("", "String", "abc")
        Assert.IsFalse(result.Item1)
        Assert.IsTrue(Not String.IsNullOrWhiteSpace(result.Item2))
    End Sub

    <TestMethod()> Public Sub Reject_Unsupported_Type()
        Dim result = SettingsWriter.WriteSetting("X", "Integer", 1)
        Assert.IsFalse(result.Item1)
        Assert.AreEqual("类型无效", result.Item2)
    End Sub

    <TestMethod()> Public Sub Reject_Type_Mismatch_Boolean()
        Dim result = SettingsWriter.WriteSetting("Flag", "Boolean", "True")
        Assert.IsFalse(result.Item1)
        Assert.AreEqual("类型不匹配", result.Item2)
    End Sub

    <TestMethod()> Public Sub Reject_Type_Mismatch_String()
        Dim result = SettingsWriter.WriteSetting("Path", "String", True)
        Assert.IsFalse(result.Item1)
        Assert.AreEqual("类型不匹配", result.Item2)
    End Sub

    <TestMethod()> Public Sub Create_New_String_Setting_Succeeds()
        Dim name = "Test_String_New"
        Dim value = "C:\\Data"
        Dim result = SettingsWriter.WriteSetting(name, "String", value)
        Assert.IsTrue(result.Item1)
        Assert.IsNull(result.Item2)
        Assert.AreEqual(value, CStr(My.Settings(name)))
    End Sub

    <TestMethod()> Public Sub Create_New_Boolean_Setting_Succeeds()
        Dim name = "Test_Boolean_New"
        Dim value As Boolean = True
        Dim result = SettingsWriter.WriteSetting(name, "Boolean", value)
        Assert.IsTrue(result.Item1)
        Assert.IsNull(result.Item2)
        Assert.AreEqual(value, CBool(My.Settings(name)))
    End Sub

    <TestMethod()> Public Sub Update_Existing_String_Setting_Succeeds()
        Dim name = "Test_Update_String"
        Dim first = SettingsWriter.WriteSetting(name, "String", "A")
        Assert.IsTrue(first.Item1)
        Dim second = SettingsWriter.WriteSetting(name, "String", "B")
        Assert.IsTrue(second.Item1)
        Assert.AreEqual("B", CStr(My.Settings(name)))
    End Sub

    <TestMethod()> Public Sub Thread_Safety_Concurrent_Writes_No_Exception()
        Dim name = "Test_ThreadSafe_Bool"
        Dim exCount As Integer = 0
        Parallel.For(0, 50, Sub(i)
                                Try
                                    Dim val As Boolean = (i Mod 2 = 0)
                                    Dim r = SettingsWriter.WriteSetting(name, "Boolean", val)
                                    If Not r.Item1 Then
                                        Interlocked.Increment(exCount)
                                    End If
                                Catch
                                    Interlocked.Increment(exCount)
                                End Try
                            End Sub)
        Assert.AreEqual(0, exCount)
        Assert.IsInstanceOfType(My.Settings(name), GetType(Boolean))
    End Sub

End Class