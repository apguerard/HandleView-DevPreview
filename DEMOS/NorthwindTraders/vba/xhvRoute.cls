VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvRoute"
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
' This class represents a route - Routes are used to navigate in the application.
'
Option Explicit

Private Type TxhvRoute
    ComponentName As String
    ExitGate As String
    Guid As String
    matchPattern As String
    Params As Scripting.dictionary
    Path As String
    RouterPortName As String
    SecurityGate As String
End Type

Private this As TxhvRoute

Private Sub Class_Initialize()
    this.Guid = NewGUID()
End Sub


Public Property Get ComponentName() As String
    ComponentName = this.ComponentName
End Property
Public Property Let ComponentName(Value As String)
    this.ComponentName = Value
End Property

Public Property Get ExitGate() As String
    ExitGate = this.ExitGate
End Property
Public Property Let ExitGate(Value As String)
    this.ExitGate = Value
End Property

Public Property Get Guid() As String
    Guid = this.Guid
End Property
Public Property Let Guid(Value As String)
    this.Guid = Value
End Property

'ReadOnly
Public Property Get matchPattern() As String
    If this.matchPattern = vbNullString Then
        this.matchPattern = getMatchPattern()
    End If
    matchPattern = this.matchPattern
End Property

Public Property Get Params() As Scripting.dictionary
    Set Params = this.Params
End Property
Public Property Set Params(Value As Scripting.dictionary)
    Set this.Params = Value
End Property

Public Property Get Path() As String
    Path = this.Path
End Property
Public Property Let Path(Value As String)
    this.Path = Value
End Property

Public Property Get RouterPortName() As String
    RouterPortName = this.RouterPortName
End Property
Public Property Let RouterPortName(Value As String)
    this.RouterPortName = Value
End Property

Public Property Get SecurityGate() As String
    SecurityGate = this.SecurityGate
End Property
Public Property Let SecurityGate(Value As String)
    this.SecurityGate = Value
End Property

''
' CloneRoute a route
'
' @return A new xhvRoute, CloneRouted from this object
Public Function CloneRoute() As xhvRoute

    Dim route As New xhvRoute
    Dim varKey As Variant
    
    'CloneRoute base type properties
    route.RouterPortName = this.RouterPortName
    route.ComponentName = this.ComponentName
    route.ExitGate = this.ExitGate
    route.Guid = this.Guid
    route.Path = this.Path
    route.SecurityGate = this.SecurityGate
    'Create a new params dic
    Set route.Params = New Scripting.dictionary

    Set CloneRoute = route

End Function

''
' Generate the pattern from the route defintion to compare it to the asked route in the router
'
' @return A string containing the pattern of the route.
Private Function getMatchPattern() As String

On Error GoTo ERR_

Dim route As String
Dim tempString As String
Dim sChar As String
Dim i As Integer

route = this.Path

For i = 1 To Len(route)
    sChar = Mid(route, i, 1)
    Select Case sChar
        Case "/"
            If peak(route, i + 1) <> "/" Then
                i = i + 1
                Do While (Mid(route, i, 1) <> "/" And Mid(route, i, 1) <> "&") And i <= Len(route)
                    i = i + 1
                Loop
                i = i - 1
            Else
                tempString = tempString & sChar
                If peak(route, i + 1) = "/" Then
                    i = i + 1
                    tempString = tempString & Mid(route, i, 1)
                End If
            End If
        Case "&"
            tempString = tempString & sChar
            If peak(route, i + 1) = "&" Then
                i = i + 1
                tempString = tempString & Mid(route, i, 1)
            End If
        
        Case Else
            tempString = tempString & sChar
    End Select
Next

getMatchPattern = tempString

Exit Function

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        ReThrow
    End If
End Function

Private Function peak(tempString As String, position As Integer) As String
    peak = Mid(tempString, position, 1)
End Function



