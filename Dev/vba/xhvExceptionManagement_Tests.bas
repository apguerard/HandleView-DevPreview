Attribute VB_Name = "xhvExceptionManagement_Tests"
Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("lib.HandleView.Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("ExceptionHandlers")
Private Sub Throw_WithParams_ShouldRaiseExeception()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Throw 999, "TestMethod", "Testing Method expected error"
    
    'Assert:
    

TestExit:
    Exit Sub
TestFail:
    If Err.Number <> 999 Then
        Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    End If
End Sub
'@TestMethod("ExceptionHandlers")
Private Sub ReThrow_WithParams_ShouldRaiseException()                        'TODO Rename test
    On Error Resume Next
    'Arrange:

    'Act:
    Throw 999, "TestModule", "Expected Test Exception"
    ReThrow
    
    'Assert:
    If InStr(Err.Description, "Rethrown") > 0 Then
        Assert.Pass
    Else
        On Error GoTo TestFail:
        Assert.Fail "Exception was not rethrown"
    End If

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

