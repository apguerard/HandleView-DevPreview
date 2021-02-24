VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvRouterPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Routing

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Gu�rard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents a RouterPort - RouterPorts are used to render components in the HTML Document.
'
Option Explicit

Private Type TxhvPort
    ChildPorts As Collection
    Component As xhvIController
    DOMNodeRef As MSHTML.HTMLGenericElement
    Guid As String
    Name As String
    ParentPort As xhvRouterPort
End Type

Private this  As TxhvPort


Private Sub Class_Initialize()
    this.Guid = NewGUID()
    Set this.ChildPorts = New Collection
End Sub

Public Property Get ChildPorts() As Collection
    Set ChildPorts = this.ChildPorts
End Property
Public Property Set ChildPorts(Value As Collection)
    Set this.ChildPorts = Value
End Property

Public Property Get Component() As xhvIController
    Set Component = this.Component
End Property
Public Property Set Component(Value As xhvIController)
    Set this.Component = Value
End Property

Public Property Get DOMNodeRef() As MSHTML.HTMLGenericElement
    Set DOMNodeRef = this.DOMNodeRef
End Property
Public Property Set DOMNodeRef(Value As MSHTML.HTMLGenericElement)
    Set this.DOMNodeRef = Value
End Property


'ReadOnly
Public Property Get Guid() As String
    Guid = this.Guid
End Property

Public Property Get Name() As String
    Name = this.Name
End Property
Public Property Let Name(Value As String)
    this.Name = Value
End Property

Public Property Get ParentPort() As xhvRouterPort
    Set ParentPort = this.ParentPort
End Property
Public Property Set ParentPort(Value As xhvRouterPort)
    Set this.ParentPort = Value
End Property
