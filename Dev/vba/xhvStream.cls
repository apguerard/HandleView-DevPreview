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

Public Property Let stream(value As String)
    myStream = value
    position = 0
    length = Len(myStream)
End Property


Public Function Peak(Optional PeakSize As Integer = 1) As String

    If position + PeakSize <= length Then
        Peak = Mid(myStream, position + 1, PeakSize)
    Else
        Peak = vbNullString
    End If

End Function


Public Function moveNext() As String

    position = position + 1
    If position <= length Then
        moveNext = Mid(myStream, position, 1)
    Else
        moveNext = vbNullString
    End If
    
End Function


Public Function currentChar() As String
    
    If position <= length Then
        currentChar = Mid(myStream, position, 1)
    Else
        currentChar = vbNullString
    End If

End Function

