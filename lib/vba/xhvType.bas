Attribute VB_Name = "xhvType"
'@Folder lib.HandleView.Config

' Copyright (C) 2019 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This module contains the type def of global type used in the framework
'
Option Explicit

''
' Contains base properties for a Component, based on xhvIController. Used in all concrete implementation of components.
'
Public Type TxhvIController
    AncestorsList As String
    ComponentObject As Object
    ChildComponents As Collection
    IsUsedAsEntryPoint As Boolean
    Guid As String
    NameType As String
    ParentComponent As xhvIController
    RouterPort As xhvRouterPort
    TemplateString As String
    TemplateUrl As String
    View As MSHTML.HTMLGenericElement
    WrapperElementType As String
End Type

