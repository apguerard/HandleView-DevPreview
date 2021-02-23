VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvLoggerTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
' This class represents the Logger for a text file
'
Option Explicit

Private Const MODULE_NAME As String = "xhvLoggerTextFile"

Implements xhvILogger

Private Type TxhvLoggerTextFile
    MinimumLogLevel As xhvENUM_LogLevel
    FileName As String
    overwriteOnAppOpen As Boolean
    firstWrite As Boolean
    enhancements As Scripting.dictionary
End Type
Private this As TxhvLoggerTextFile

Private Sub Class_Initialize()
    this.firstWrite = True
    Set this.enhancements = New Scripting.dictionary
End Sub

Private Property Let xhvILogger_minimumLogLevel(Value As xhvENUM_LogLevel)
    this.MinimumLogLevel = Value
End Property
Private Property Get xhvILogger_minimumLogLevel() As xhvENUM_LogLevel
    xhvILogger_minimumLogLevel = this.MinimumLogLevel
End Property
Private Property Set xhvILogger_enhancements(Value As Scripting.dictionary)
    Set this.enhancements = Value
End Property
Private Property Get xhvILogger_enhancements() As Scripting.dictionary
    Set xhvILogger_enhancements = this.enhancements
End Property


''
' Set properties of this copncerte implementation of xhvILogger
'
' @param loggerProperties Dictionary containing the properties. These properties must be set when configuring the Logger
'
Private Sub xhvILogger_setProperties(loggerProperties As Scripting.IDictionary)

On Error GoTo ERR_

    If loggerProperties.Exists("fileName") Then
        this.FileName = loggerProperties("fileName")
    Else
        Throw 3006, , "The Logger 'xhvLoggerTextFile' needs a property named 'fileName', but it was not provided on initialization. Add this property when initializing the Logger."
    End If

    If loggerProperties.Exists("overwriteOnAppOpen") Then
        this.overwriteOnAppOpen = loggerProperties("overwriteOnAppOpen")
    Else
        Throw 3007, , "The Logger 'xhvLoggerTextFile' needs a property named 'overwriteOnAppOpen', but it was not provided on initialization. Add this property when initializing the Logger."
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


Private Sub xhvILogger_logTrace(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logEntry As String
    
    logEntry = fillTemplate(template, args)
    formatWithTimeStamp logEntry, TRACE_LEVEL
    formatWithEnhancements logEntry
    
    writeLogToFile logEntry
    
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

Private Sub xhvILogger_logInfo(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logEntry As String
    
    logEntry = fillTemplate(template, args)
    formatWithTimeStamp logEntry, INFO_LEVEL
    formatWithEnhancements logEntry
    
    writeLogToFile logEntry
    
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

Private Sub xhvILogger_logDebug(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logEntry As String
    
    logEntry = fillTemplate(template, args)
    formatWithTimeStamp logEntry, DEBUG_LEVEL
    formatWithEnhancements logEntry
    
    writeLogToFile logEntry
    
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

Private Sub xhvILogger_logWarning(template As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logEntry As String
    
    logEntry = fillTemplate(template, args)
    formatWithTimeStamp logEntry, WARNING_LEVEL
    formatWithEnhancements logEntry
    
    writeLogToFile logEntry
    
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

Private Sub xhvILogger_logError(template As String, errNumber As Long, errDescription As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logEntry As String
    
    logEntry = fillTemplate(template, args)
    formatWithErrNo logEntry, errNumber, errDescription
    formatWithTimeStamp logEntry, ERROR_LEVEL
    formatWithEnhancements logEntry
    
    writeLogToFile logEntry
    
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

Private Sub xhvILogger_logCritical(template As String, errNumber As Long, errDescription As String, ParamArray args() As Variant)

On Error GoTo ERR_

    Dim logEntry As String
    
    logEntry = fillTemplate(template, args)
    formatWithErrNo logEntry, errNumber, errDescription
    formatWithTimeStamp logEntry, CRITICAL_LEVEL
    formatWithEnhancements logEntry
    
    writeLogToFile logEntry
   
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
    xhvILogger_logTrace template, args
End Sub
Public Sub LogDebug(template As String, ParamArray args() As Variant)
    xhvILogger_logDebug template, args
End Sub
Public Sub LogInfo(template As String, ParamArray args() As Variant)
    xhvILogger_logInfo template, args
End Sub
Public Sub LogWarning(template As String, ParamArray args() As Variant)
    xhvILogger_logWarning template, args
End Sub
Public Sub LogError(template As String, errNumber As Long, errDescription As String, ParamArray args() As Variant)
    xhvILogger_logError template, errNumber, errDescription, args
End Sub
Public Sub LogCritical(template As String, errNumber As Long, errDescription As String, ParamArray args() As Variant)
    xhvILogger_logCritical template, errNumber, errDescription, args
End Sub

''
' Fill the string template with values received from the calling of a logging action
' This template must be in the following form : "My template text with {{valueToFill}}"

' NOTE: The name of the valueToFill doesn't have to match the name of parameters that fills it. Values are filled in the temaplte order.
' e.g. interpolation 1 is filled with args 1
'
' @param template Template to fill
' @param args ParamArray that should contain the value with which we replace the interpolation in the template string.
'
Private Function fillTemplate(template As String, ParamArray args() As Variant) As String

On Error GoTo ERR_

    On Error GoTo ERR_

    Dim param As Variant
    Dim i As Integer
    Dim interpolations As Collection
    
    Set interpolations = getInterpolations(template)
    
    fillTemplate = template
    
    i = 0

    For Each param In interpolations
        On Error Resume Next
            fillTemplate = Replace(fillTemplate, "{{" & param & "}}", args(0)(0)(i))
            If Err.Number = 9 Then
                On Error GoTo ERR_
                Throw 3008, , "The number of values don't match the number of variables in the template in the xhvLoggerText.fillTemplate() private function." & vbCrLf & vbCrLf & "Please correct the following template:" & _
                              vbCrLf & vbCrLf & template
                
            ElseIf Err.Number <> 0 Then
                On Error GoTo ERR_
                Throw 3000, , "Unexpected exception catched in " & MODULE_NAME & "." & template & vbCrLf & vbCrLf & _
                              "Original Err.number: " & Err.Number & vbCrLf & _
                              "Original Err.Description: " & Err.Description
            End If
        i = i + 1
    Next
        
    On Error GoTo ERR_
        
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


''
' Add a timestamp and the loging level to the logged message
'
' @param message Original Message
' @param level Level of logging
'
Private Sub formatWithTimeStamp(ByRef message As String, level As xhvENUM_LogLevel)
        
        message = "[" & GetTimeStamp() & " " & GetLoggingLevelLabel(level) & "] " & message
        
End Sub

''
' Add error number and error description to the logged message
' This Sub is called when logging error and critical logging level
'
' @param message Original Message
' @param errNumber Error number
' @param errDescription Error description
'
Private Sub formatWithErrNo(ByRef message As String, errNumber As Long, errDescription As String)
        
        message = message & vbCrLf & "Error Number:" & errNumber & " : Error Description:" & errDescription
        
End Sub


''
' Add the registered enhancements values to the message in the form of a Json string
'
' @param message Original Message
'
Private Sub formatWithEnhancements(ByRef message As String)

On Error GoTo ERR_

    Dim vKey As Variant
    Dim vKey2 As Variant
    Dim enhancementsItems As Scripting.dictionary
    Dim properties As Scripting.dictionary
    Dim enhancement As xhvILoggerEnhancement
    
    Set properties = New Scripting.dictionary
    'Add the logger enhancement properties
    For Each vKey In this.enhancements.Keys()
        Set enhancement = this.enhancements(vKey)
        Set enhancementsItems = enhancement.returnEnhancementItems()
        For Each vKey2 In enhancementsItems
            properties.Add vKey2, enhancementsItems(vKey2)
        Next vKey2
    Next
    
    If properties.Count <> 0 Then
        message = message & vbCrLf & ConvertToJson(properties)
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
' Extract the placeholders from the message
' These placeholders will eventually be replaced by their corresponding values in the fillTemplate Sub
'
' @param message Original Message
' @return A collection of placeholders in the message.
'
Private Function getInterpolations(tempStringlate As String) As Collection

On Error GoTo ERR_

    Dim Stream As xhvStream
    Dim myChar As String
    Dim tempString As String
    
    Set getInterpolations = New Collection
    Set Stream = New xhvStream
    
    Stream.Stream = tempStringlate
    
    myChar = Stream.MoveNext
    
    Do While myChar <> vbNullString

        If myChar = "{" Then
        
            If Stream.peak(1) = "{" Then
                Stream.MoveNext
                tempString = vbNullString
                myChar = Stream.MoveNext
                Do While myChar <> vbNullString
                    If myChar = "}" And Stream.peak() = "}" Then Exit Do
                    tempString = tempString & myChar
                    myChar = Stream.MoveNext()
                Loop

                getInterpolations.Add tempString
                
            End If
        
        End If
            
        myChar = Stream.MoveNext
        
    Loop
    
    Set Stream = Nothing

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


Private Sub writeLogToFile(logEntry As String)

On Error GoTo ERR_

    Dim noFile As Integer
    Dim errNum As Long
    
    noFile = FreeFile
    
    If this.firstWrite Then
        'Check if we configure the logger to overwrite the file on each startup of the app. If yes, we clean it
        If this.overwriteOnAppOpen Then
            On Error Resume Next
               Open this.FileName For Output As #noFile
                If Err.Number <> 0 Then
                    errNum = Err.Number
                    On Error GoTo ERR_
                    If errNum = 55 Then
                        Throw 3002, , "The file'" & this.FileName & "' is already opened. The " & MODULE_NAME & " cannot write to log file."
                    ElseIf errNum <> 0 Then
                        Throw 3003, , "Unexpected error when opening the file '" & this.FileName & "' for logging. The " & MODULE_NAME & " cannot write to log file."
                    End If
                End If
            
            On Error GoTo ERR_
            
            Print #noFile, vbNullString
            Close #noFile ' close the file
        End If
    End If
    
    On Error Resume Next
        Open this.FileName For Append As #noFile
        If Err.Number <> 0 Then
            errNum = Err.Number
            On Error GoTo ERR_
                If errNum = 55 Then
                    Throw 3002, , "The file'" & this.FileName & "' is already opened. The " & MODULE_NAME & " cannot write to log file."
                ElseIf errNum <> 0 Then
                    Throw 3003, , "Unexpected error when opening the file '" & this.FileName & "' for logging. The " & MODULE_NAME & " cannot write to log file."
                End If
        End If
    On Error GoTo ERR_
    
    Print #noFile, logEntry
    Close #noFile ' close the file

    If this.firstWrite Then this.firstWrite = False
    
Exit Sub

ERR_:
    'Try to close the file
    Close #noFile
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        ReThrow
    End If
End Sub




