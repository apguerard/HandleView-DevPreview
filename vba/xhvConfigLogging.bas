Attribute VB_Name = "xhvConfigLogging"
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
' This module is used to configure logging services in a centralized place.
' Eventually, these configuration could be placed in a config file.
' Think of this module as if it was a config file.
'
Option Explicit
Private Const XHV_FILE_VERSION As String = "0.0.1"

Private Const MODULE_NAME As String = "xhvConfigLogging"

''
' This function is the place to configure the logging service.
'
' Eventually, these configuration could be placed in a config file.
'
' NOTE: Logger added after the logger enhancements are added won't inherit these enhancements.
'
Public Sub ConfigureLoggingServices()
On Error GoTo ERR_

    '**************************************************
    '
    ' --->  IMPORTANT  <---
    '
    ' DON'T FORGET TO ADD USED LOGGER IN THE xhvLoggerFactory
    '
    ' The Logging system will rethrow any error to this level.
    ' Decide here how you want to handle error in Logging System.
    '
    ' Ex: Do you want to stop the application if the Logging System cannot be configured correctly ? Do you want to continue the application and silent fail ?
    '
    '**************************************************
    
    'Get configuration for the log system
    '=
    xhvLog.EnabledLogging = Configuration("App.EnabledLogging")
    xhvLog.MinimumLogLevel = Configuration("App.MinimumLogLevel")
    
    ' Register your logger below this line -->
    '
    'This section only registers logging to console if we are in the DEV environment - You can remove if you want...
    If Configuration("Environment.RunningEnvironment") = "DEV" Then
        xhvLog.UseLogger "xhvLoggerConsole", , TRACE_LEVEL '<-- override global minimumLogLevel for this logger
    End If
    
    'This part enables the TextFile Logger - Comment it out if you don't want to use it.

    'xhvLoggerTextFile needs some properties to work.
    Dim props As New Scripting.dictionary
    props.Add "fileName", CurrentProject.Path & "\xhv.log"
    props.Add "overwriteOnAppOpen", False
    xhvLog.UseLogger "xhvLoggerTextFile", props

    
    'This part enables the Json Logger - Comment it out if you don't want to use it.

    'xhvLoggerJson needs some properties to work.
    Dim propsJson As New Scripting.dictionary
    propsJson.Add "fileName", CurrentProject.Path & "\xhv.json"
    propsJson.Add "overwriteOnAppOpen", False
    xhvLog.UseLogger "xhvLoggerJson", propsJson

    
    
    '
    ' Register your own loggers above this line -->

    
    'Add enhancements here
    'NOTE: Logger added after the logger enhancements are added won't inherit these enhancements.
    '
    
    xhvLog.UseLoggerEnhancement "xhvLogEnhancementWithUserContext"
    
    
    

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
