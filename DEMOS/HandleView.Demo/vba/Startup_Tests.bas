Attribute VB_Name = "Startup_Tests"
Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("lib.HandleView.Tests")

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
    #End If
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

'@TestMethod("Bootstrap")
Private Sub Configuration_WithValidConfigId_ReturnsValue()
    On Error GoTo TestFail
    
    'Arrange:
    SetConst
    xhvConfigurator.AddLocalDB xhvConst.APP_CONFIG_TABLE_NAME, xhvConst.APP_CONFIG_ID_FIELD, xhvConst.APP_CONFIG_VALUE_FIELD
    Dim res As Variant
    
    'Act:
    res = Configuration("App.EnabledLogging")
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Bootstrap")
Private Sub Configuration_WithINVALIDConfigId_Throws2002()
    On Error GoTo TestFail

    'Arrange:
    SetConst
    xhvConfigurator.AddLocalDB xhvConst.APP_CONFIG_TABLE_NAME, xhvConst.APP_CONFIG_ID_FIELD, xhvConst.APP_CONFIG_VALUE_FIELD
    Dim res As Variant
    
    'Act:
    res = Configuration("INVALID")
    
    'Assert
    
    
TestExit:
    Exit Sub
    
TestFail:
    If Err.Number <> 2002 Then
        Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    End If
End Sub

'@TestMethod("Bootstrap")
Private Sub Config_RootRouterPort_ShouldReturnRoot()
    On Error GoTo TestFail
    
    'Arrange:
    SetConst
    
    'Act:

    'Assert:
    Assert.IsNotNothing Configuration("App.RootRouterPort")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


