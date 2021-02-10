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
    componentName As String
    css As Scripting.dictionary
    htmlTemplate As String
    scripts As Scripting.dictionary
End Type

Private this As TxhvCachedComponent

Private Sub Class_Initialize()
    Set this.scripts = New Scripting.dictionary
    Set this.css = New Scripting.dictionary
End Sub

Public Property Get componentName() As String
    componentName = this.componentName
End Property
Public Property Let componentName(value As String)
    this.componentName = value
End Property

Public Property Get css() As Scripting.dictionary
    Set css = this.css
End Property
Public Property Set css(value As Scripting.dictionary)
    Set this.css = value
End Property

Public Property Get htmlTemplate() As String
    htmlTemplate = this.htmlTemplate
End Property
Public Property Let htmlTemplate(value As String)
    this.htmlTemplate = value
End Property

Public Property Get scripts() As Scripting.dictionary
    Set scripts = this.scripts
End Property
Public Property Set scripts(value As Scripting.dictionary)
    Set this.scripts = value
End Property
