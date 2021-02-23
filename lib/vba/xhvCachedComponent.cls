VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvCachedComponent"
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

' This class represents a Cached component in memory. xhvCachedComponent is used by the ComponentFactory when the UseComponentsCacheManagement config is True
' This class is only a model class. -->  No functionality
'
Option Explicit

Private Type TxhvCachedComponent
    ComponentName As String
    Css As Scripting.dictionary
    HtmlTemplate As String
    Scripts As Scripting.dictionary
End Type

Private this As TxhvCachedComponent

Private Sub Class_Initialize()
    Set this.Scripts = New Scripting.dictionary
    Set this.Css = New Scripting.dictionary
End Sub

Public Property Get ComponentName() As String
    ComponentName = this.ComponentName
End Property
Public Property Let ComponentName(Value As String)
    this.ComponentName = Value
End Property

Public Property Get Css() As Scripting.dictionary
    Set Css = this.Css
End Property
Public Property Set Css(Value As Scripting.dictionary)
    Set this.Css = Value
End Property

Public Property Get HtmlTemplate() As String
    HtmlTemplate = this.HtmlTemplate
End Property
Public Property Let HtmlTemplate(Value As String)
    this.HtmlTemplate = Value
End Property

Public Property Get Scripts() As Scripting.dictionary
    Set Scripts = this.Scripts
End Property
Public Property Set Scripts(Value As Scripting.dictionary)
    Set this.Scripts = Value
End Property
