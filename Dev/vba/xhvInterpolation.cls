VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvInterpolation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Helpers

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents the interpolation used in a HTML template
' This  is only a model class.
'

Option Explicit

Private Type TxhvInterpolation
    typeInterpolation As String
    value As String
End Type
Private this As TxhvInterpolation

Public Property Get typeInterpolation() As String
    typeInterpolation = this.typeInterpolation
End Property
Public Property Let typeInterpolation(value As String)
    this.typeInterpolation = value
End Property

Public Property Get value() As String
    value = this.value
End Property
Public Property Let value(value As String)
    this.value = value
End Property
