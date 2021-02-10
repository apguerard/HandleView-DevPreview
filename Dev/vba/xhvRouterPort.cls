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
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents a RouterPort - RouterPorts are used to render component in the HTML Document.
'
Option Explicit

Private Type TxhvPort
    childPorts As Collection
    component As xhvIController
    DOMNodeRef As MSHTML.HTMLGenericElement
    Guid As String
    name As String
    parentPort As xhvRouterPort
End Type

Private this  As TxhvPort


Private Sub Class_Initialize()
    this.Guid = NewGUID()
    Set this.childPorts = New Collection
End Sub

Public Property Get childPorts() As Collection
    Set childPorts = this.childPorts
End Property
Public Property Set childPorts(value As Collection)
    Set this.childPorts = value
End Property

Public Property Get component() As xhvIController
    Set component = this.component
End Property
Public Property Set component(value As xhvIController)
    Set this.component = value
End Property

Public Property Get DOMNodeRef() As MSHTML.HTMLGenericElement
    Set DOMNodeRef = this.DOMNodeRef
End Property
Public Property Set DOMNodeRef(value As MSHTML.HTMLGenericElement)
    Set this.DOMNodeRef = value
End Property


'ReadOnly
Public Property Get Guid() As String
    Guid = this.Guid
End Property

Public Property Get name() As String
    name = this.name
End Property
Public Property Let name(value As String)
    this.name = value
End Property

Public Property Get parentPort() As xhvRouterPort
    Set parentPort = this.parentPort
End Property
Public Property Set parentPort(value As xhvRouterPort)
    Set this.parentPort = value
End Property
