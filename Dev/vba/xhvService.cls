VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Services

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Represents a service in the DI system.
' This is a model class only.
'
Option Explicit

Private Type TxhvService
    interface As String
    scope As xhvENUM_DependencyInjectionScope
    useClass As String
    useFactory As String
End Type
Private this As TxhvService

Public Property Get interface() As String
    interface = this.interface
End Property
Public Property Let interface(value As String)
    this.interface = interface
End Property

Public Property Get scope() As xhvENUM_DependencyInjectionScope
    scope = this.scope
End Property
Public Property Let scope(value As xhvENUM_DependencyInjectionScope)
    this.scope = value
End Property

Public Property Get useClass() As String
    useClass = this.useClass
End Property
Public Property Let useClass(value As String)
    this.useClass = value
End Property

Public Property Get useFactory() As String
    useFactory = this.useFactory
End Property
Public Property Let useFactory(value As String)
    this.useFactory = value
End Property
