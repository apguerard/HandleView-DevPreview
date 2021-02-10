VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvIController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Controllers

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents the Component Interface
'
Option Explicit

Public Property Get AncestorsList() As String
End Property
Public Property Let AncestorsList(value As String)
End Property

Public Property Get componentObject() As Object
End Property


Public Property Get ChildComponents() As Collection
End Property
Public Property Set ChildComponents(value As Collection)
End Property

'ReadOnly
Public Property Get Guid() As String
End Property


Public Property Get IsUsedAsEntryPoint() As Boolean
End Property
Public Property Let IsUsedAsEntryPoint(value As Boolean)
End Property


'ReadOnly
Public Property Get NameType() As String
End Property
Public Property Let NameType(value As String)
End Property

Public Property Get ParentComponent() As xhvIController
End Property
Public Property Set ParentComponent(value As xhvIController)
End Property

Public Property Get routerPort() As xhvRouterPort
End Property
Public Property Set routerPort(value As xhvRouterPort)
End Property

Public Property Get TemplateString() As String
End Property
Public Property Let TemplateString(value As String)
End Property

'ReadOnly
Public Property Get TemplateUrl() As String
End Property

Public Property Get view() As MSHTML.HTMLGenericElement
End Property
Public Property Set view(value As MSHTML.HTMLGenericElement)
End Property

Public Property Get WrapperElementType() As String
End Property
Public Property Let WrapperElementType(value As String)
End Property


Public Sub Init(params As Scripting.dictionary)
End Sub

Public Sub Dispose()
End Sub

Public Function Render() As MSHTML.HTMLGenericElement
End Function

Public Sub OnChildrenPortRendered()
End Sub

Public Sub AfterParentInsertedIntoDOM()
End Sub

Public Sub AfterInsertIntoDOM()
End Sub

Public Sub OnAfterInit(attributes As Scripting.dictionary)
End Sub

Public Function InterComponentCall(ParamArray agrs() As Variant) As Boolean
End Function

Public Sub Refresh()
End Sub

Public Sub SubscribeEventDispatcher()
End Sub

