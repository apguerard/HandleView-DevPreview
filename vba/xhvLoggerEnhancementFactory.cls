VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvLoggerEnhancementFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder lib.HandleView.Logging

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Gu�rard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
' *********************************************************************

''
' This class is responsible for creating instances of enhancement classes for the HandleView Logger
'
Option Explicit

''
' Creates and initilizes a new enhancement
'
' @param enhancementclassName The name of the enhancement to create.
' @return A concrete xhvLoggerEnhacement through the xhvILoggerEnhancement Interface
'
Public Function CreateLoggerEnhancement(enhancementclassName As String) As xhvILoggerEnhancement

On Error GoTo ERR_

    Select Case UCase(enhancementclassName)
        Case "XHVLOGENHANCEMENTWITHUSERCONTEXT"
            Set CreateLoggerEnhancement = New xhvLogEnhancementWithUserContext
            
        Case Else
            Throw 2020, , "Logger Enhancement Factory cannot find a log enhancement class named '" & enhancementclassName & "'." & vbCrLf & vbCrLf & "Check your class name or add the related class name in xhvLoggerEnhancementFactory.createLoggerEnhancement function."
    End Select

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
