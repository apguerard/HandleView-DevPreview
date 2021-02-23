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
    Interface As String
    Scope As xhvENUM_DependencyInjectionScope
    UseClass As String
    UseFactory As String
End Type
Private this As TxhvService

Public Property Get Interface() As String
    Interface = this.Interface
End Property
Public Property Let Interface(Value As String)
    this.Interface = Interface
End Property

Public Property Get Scope() As xhvENUM_DependencyInjectionScope
    Scope = this.Scope
End Property
Public Property Let Scope(Value As xhvENUM_DependencyInjectionScope)
    this.Scope = Value
End Property

Public Property Get UseClass() As String
    UseClass = this.UseClass
End Property
Public Property Let UseClass(Value As String)
    this.UseClass = Value
End Property

Public Property Get UseFactory() As String
    UseFactory = this.UseFactory
End Property
Public Property Let UseFactory(Value As String)
    this.UseFactory = Value
End Property
