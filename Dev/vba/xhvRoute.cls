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
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents a route - Routes are used to navigate in the application
'
Option Explicit

Private Type TxhvRoute
    componentName As String
    exitGate As String
    Guid As String
    matchPattern As String
    params As Scripting.dictionary
    path As String
    routerPortName As String
    securityGate As String
End Type

Private this As TxhvRoute

Private Sub Class_Initialize()
    this.Guid = NewGUID()
End Sub


Public Property Get componentName() As String
    componentName = this.componentName
End Property
Public Property Let componentName(value As String)
    this.componentName = value
End Property

Public Property Get exitGate() As String
    exitGate = this.exitGate
End Property
Public Property Let exitGate(value As String)
    this.exitGate = value
End Property

Public Property Get Guid() As String
    Guid = this.Guid
End Property
Public Property Let Guid(value As String)
    this.Guid = value
End Property

'ReadOnly
Public Property Get matchPattern() As String
    If this.matchPattern = vbNullString Then
        this.matchPattern = getMatchPattern()
    End If
    matchPattern = this.matchPattern
End Property

Public Property Get params() As Scripting.dictionary
    Set params = this.params
End Property
Public Property Set params(value As Scripting.dictionary)
    Set this.params = value
End Property

Public Property Get path() As String
    path = this.path
End Property
Public Property Let path(value As String)
    this.path = value
End Property

Public Property Get routerPortName() As String
    routerPortName = this.routerPortName
End Property
Public Property Let routerPortName(value As String)
    this.routerPortName = value
End Property

Public Property Get securityGate() As String
    securityGate = this.securityGate
End Property
Public Property Let securityGate(value As String)
    this.securityGate = value
End Property

''
' CloneRoute a route
'
' @return A new xhvRoute, CloneRouted from this object
Public Function CloneRoute() As xhvRoute

    Dim route As New xhvRoute
    Dim varKey As Variant
    
    'CloneRoute base type properties
    route.routerPortName = this.routerPortName
    route.componentName = this.componentName
    route.exitGate = this.exitGate
    route.Guid = this.Guid
    route.path = this.path
    route.securityGate = this.securityGate
    'Create a new params dic
    Set route.params = New Scripting.dictionary

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

route = this.path

For i = 1 To Len(route)
    sChar = Mid(route, i, 1)
    Select Case sChar
        Case "/"
            If Peak(route, i + 1) <> "/" Then
                i = i + 1
                Do While (Mid(route, i, 1) <> "/" And Mid(route, i, 1) <> "&") And i <= Len(route)
                    i = i + 1
                Loop
                i = i - 1
            Else
                tempString = tempString & sChar
                If Peak(route, i + 1) = "/" Then
                    i = i + 1
                    tempString = tempString & Mid(route, i, 1)
                End If
            End If
        Case "&"
            tempString = tempString & sChar
            If Peak(route, i + 1) = "&" Then
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

Private Function Peak(tempString As String, position As Integer) As String
    Peak = Mid(tempString, position, 1)
End Function



