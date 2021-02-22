Attribute VB_Name = "Navigation_Tests"
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

'@TestMethod("Navigation")
Private Sub Configuration_StartupRoute_ShouldReturnValidPath()
    On Error GoTo TestFail
    
    'Arrange:
    SetConst
    Dim res As String
    
    'Act:
    res = Configuration("App.StartupRoute")
    
    'Assert:
    Assert.AreEqual res, "app||home"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


