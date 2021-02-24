VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvLoggerFactory"
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
'

''
' This class is responsible for creating and initializing Logger instances in the Logging System
'
Option Explicit

''
' Creates and initilizes a new Logger
'
' @param loggerclassName The name of the Logger to create.
' @param loggerProperties Dictionnary with the specific properties for the Logger
' @param minimumLogLevel The minimal logging level sets in the configuration
' @return A concrete xhvLogger through the xhvILogger Interface
'
Public Function CreateLogger(loggerclassName As String, loggerProperties As Scripting.dictionary, MinimumLogLevel As xhvENUM_LogLevel) As xhvILogger

On Error GoTo ERR_

    Select Case UCase(loggerclassName)
        Case "XHVLOGGERCONSOLE"
            Set CreateLogger = New xhvLoggerConsole
         Case "XHVLOGGERTEXTFILE"
            Set CreateLogger = New xhvLoggerTextFile
         Case "XHVLOGGERJSON"
            Set CreateLogger = New xhvLoggerJson
        Case Else
            Throw 2021, , "Logger Factory cannot find a Logger class named '" & loggerclassName & "'." & vbCrLf & vbCrLf & "Check your class name or add the related class name in xhvLoggerFactory.createLogger function."
    End Select

    CreateLogger.MinimumLogLevel = MinimumLogLevel
    CreateLogger.setProperties loggerProperties


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
