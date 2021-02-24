VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder lib.HandleView.Logging

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents the Log Static Class used for logging
' It's the main part of the Logging System
'
Option Explicit

Private Type TxhvLog
    EnabledLogging  As Boolean
    enhancements As Scripting.dictionary
    MinimumLogLevel As xhvENUM_LogLevel
    registeredLoggers As Scripting.dictionary
End Type
Private this As TxhvLog

Private Sub Class_Initialize()
    Set this.enhancements = New Scripting.dictionary
    Set this.registeredLoggers = New Scripting.dictionary
End Sub

Public Property Get EnabledLogging() As Boolean
    EnabledLogging = this.EnabledLogging
End Property
Public Property Let EnabledLogging(Value As Boolean)
    this.EnabledLogging = Value
End Property
Public Property Get MinimumLogLevel() As xhvENUM_LogLevel
    MinimumLogLevel = this.MinimumLogLevel
End Property
Public Property Let MinimumLogLevel(Value As xhvENUM_LogLevel)
    this.MinimumLogLevel = Value
End Property


''
' Register a concrete logger with the logging system. e.g. log will be processed by this logger.
'
' @param LoggerName Name of the logger. This name must match te name you used in the the xhvLoggerFactory.createLogger() Method Select Case
' @param loggerProperties Optional - If the  logger have specific properties, you can pass them through this dictionnary
' @param minimumLogLevelOverride Optional - Pass the minimun level of logging as xhvENUM_LogLevel if you want to
'            override the global minimum level of logging of the Log system for this specific Logger. Leave empty otherwise.
'
Public Sub UseLogger(LoggerName As String, Optional loggerProperties As Scripting.dictionary, Optional minimumLogLevelOverride As xhvENUM_LogLevel = -1)

On Error GoTo ERR_

    If this.registeredLoggers.Exists(LoggerName) Then
        Throw 2018, , "The logger '" & LoggerName & "' has already been registered in the logging system and cannot be added again. Please check your Logging Configurations."
    End If

    If minimumLogLevelOverride = -1 Then
        this.registeredLoggers.Add LoggerName, xhvLoggerFactory.CreateLogger(LoggerName, loggerProperties, this.MinimumLogLevel)
    Else
        this.registeredLoggers.Add LoggerName, xhvLoggerFactory.CreateLogger(LoggerName, loggerProperties, minimumLogLevelOverride)
    End If
    
Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        ReThrow
    End If
End Sub


''
' Register a logger enhancement with every logger taht has been registered.
'
' @param LogEnhancementName Name of the logger enhancement. This name must match te name you used in the the xhvLoggerEnhancementFactory.createLoggerEnhancement() Method Select Case
'
Public Sub UseLoggerEnhancement(LogEnhancementName As String)

On Error GoTo ERR_

    Dim enhancement As xhvILoggerEnhancement
    Dim logger  As xhvILogger
    Dim vKey As Variant

    Set enhancement = xhvLoggerEnhancementFactory.CreateLoggerEnhancement(LogEnhancementName)

    For Each vKey In this.registeredLoggers.Keys()
        Set logger = this.registeredLoggers(vKey)
        If logger.enhancements.Exists(LogEnhancementName) Then
            Throw 2019, , "The logger enhancement'" & LogEnhancementName & "' has already been registered in the loggin system and cannot be added again. Please check your Logging Configurations."
        Else
            logger.enhancements.Add enhancement.enhancementName, enhancement
        End If
    Next

    'Also registered it in the xhvLog for future use?
    this.enhancements.Add enhancement.enhancementName, enhancement
    
Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        ReThrow
    End If
End Sub


Public Sub LogTrace(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logger  As xhvILogger
    Dim vKey As Variant

    If this.EnabledLogging Then
        For Each vKey In this.registeredLoggers.Keys()
            Set logger = this.registeredLoggers(vKey)
            If logger.MinimumLogLevel <= TRACE_LEVEL Then
                logger.LogTrace template, args
            End If
        Next
    End If

Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        'If not in debug mode, determine if we want to fail silently the loggin error or stop application.
        If xhvConst.FAIL_SILENT_LOG_EXCEPTION Then
            tryWriteLogSystemErrorToFile template, Err.Number, Err.Description
            Resume Next
        Else
            ReThrow
        End If
    End If
End Sub


Public Sub LogDebug(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logger  As xhvILogger
    Dim vKey As Variant

    If this.EnabledLogging Then
        For Each vKey In this.registeredLoggers.Keys()
            Set logger = this.registeredLoggers(vKey)
            If logger.MinimumLogLevel <= DEBUG_LEVEL Then
                logger.LogDebug template, args
            End If
        Next
    End If

Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        'If not in debug mode, determine if we want to fail silently the loggin error or stop application.
        If xhvConst.FAIL_SILENT_LOG_EXCEPTION Then
            tryWriteLogSystemErrorToFile template, Err.Number, Err.Description
            Resume Next
        Else
            ReThrow
        End If
    End If
End Sub


Public Sub LogInfo(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logger  As xhvILogger
    Dim vKey As Variant

    If this.EnabledLogging Then
        For Each vKey In this.registeredLoggers.Keys()
            Set logger = this.registeredLoggers(vKey)
            If logger.MinimumLogLevel <= INFO_LEVEL Then
                logger.LogInfo template, args
            End If
        Next
    End If

Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        'If not in debug mode, determine if we want to fail silently the loggin error or stop application.
        If xhvConst.FAIL_SILENT_LOG_EXCEPTION Then
            tryWriteLogSystemErrorToFile template, Err.Number, Err.Description
            Resume Next
        Else
            ReThrow
        End If
    End If
End Sub


Public Sub LogWarning(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logger  As xhvILogger
    Dim vKey As Variant

    If this.EnabledLogging Then
        For Each vKey In this.registeredLoggers.Keys()
            Set logger = this.registeredLoggers(vKey)
            If logger.MinimumLogLevel <= WARNING_LEVEL Then
                logger.LogWarning template, args
            End If
        Next
    End If

Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        'If not in debug mode, determine if we want to fail silently the loggin error or stop application.
        If xhvConst.FAIL_SILENT_LOG_EXCEPTION Then
            tryWriteLogSystemErrorToFile template, Err.Number, Err.Description
            Resume Next
        Else
            ReThrow
        End If
    End If
End Sub


Public Sub LogError(template As String, errNumber As Long, errDescription As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logger  As xhvILogger
    Dim vKey As Variant

    If this.EnabledLogging Then
        For Each vKey In this.registeredLoggers.Keys()
            Set logger = this.registeredLoggers(vKey)
            If logger.MinimumLogLevel <= ERROR_LEVEL Then
                logger.LogError template, errNumber, errDescription, args
            End If
        Next
    End If

Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        'If not in debug mode, determine if we want to fail silently the loggin error or stop application.
        If xhvConst.FAIL_SILENT_LOG_EXCEPTION Then
            tryWriteLogSystemErrorToFile template, errNumber, errDescription
            Resume Next
        Else
            ReThrow
        End If
    End If
End Sub


Public Sub LogCritical(template As String, errNumber As Long, errDescription As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logger  As xhvILogger
    Dim vKey As Variant

    If this.EnabledLogging Then
        For Each vKey In this.registeredLoggers.Keys()
            Set logger = this.registeredLoggers(vKey)
            If logger.MinimumLogLevel <= CRITICAL_LEVEL Then
                logger.LogCritical template, errNumber, errDescription, args
            End If
        Next
    End If

Exit Sub

ERR_:
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        'If not in debug mode, determine if we want to fail silently the loggin error or stop application.
        If xhvConst.FAIL_SILENT_LOG_EXCEPTION Then
            tryWriteLogSystemErrorToFile template, errNumber, errDescription
            Resume Next
        Else
            ReThrow
        End If
    End If
End Sub



''
' This function try to log an unhandle silent fail logging error in case something went wrong when logging.
'
' @param template Original template
' @param errNumber Original error number
' @param errDesc Original error description
'
Private Sub tryWriteLogSystemErrorToFile(template As String, errNumber As Long, errDesc As String)

On Error Resume Next
    Dim logEntry As String
    Dim noFile As Integer
    
    noFile = FreeFile()
    
    logEntry = "An silent fail exception happened in the Logging System." & vbCrLf
    logEntry = logEntry & "The template string was: '" & template & "'" & vbCrLf
    logEntry = logEntry & "The original error was: [" & errNumber & "] " & errDesc

    Open CurrentProject.Path & "\xhvLoggerError.log" For Append As #noFile
    Print #noFile, logEntry
    Close #noFile ' close the file
    
On Error GoTo 0

End Sub

