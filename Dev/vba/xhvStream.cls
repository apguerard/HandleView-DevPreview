VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvStream"
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
' Represents a stream of data as char.
'

Option Explicit

Private position As Integer
Private length As Integer

Private myStream As String


Public Property Get CurrentPosition() As Integer
    CurrentPosition = position
End Property

Public Property Let Stream(Value As String)
    myStream = Value
    position = 0
    length = Len(myStream)
End Property


Public Function peak(Optional PeakSize As Integer = 1) As String

    If position + PeakSize <= length Then
        peak = Mid(myStream, position + 1, PeakSize)
    Else
        peak = vbNullString
    End If

End Function


Public Function MoveNext() As String

    position = position + 1
    If position <= length Then
        MoveNext = Mid(myStream, position, 1)
    Else
        MoveNext = vbNullString
    End If
    
End Function


Public Function CurrentChar() As String
    
    If position <= length Then
        CurrentChar = Mid(myStream, position, 1)
    Else
        CurrentChar = vbNullString
    End If

End Function

