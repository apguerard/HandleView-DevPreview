VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvExceptionManager"
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
' Used for custom error (Exception) management
'
Option Explicit

Private Const MODULE_NAME As String = "xhvExceptionManager"

''
' Display an exception that happened in the framework
'
' @param errNumber Framework  number exception
' @param errMessage Framework  message exception
'
Public Sub HandleFrameworkException(errNumber As Long, errMessage As String)
On Error GoTo HandleFrameworkException_ERR
    
    If xhvConfigurator.Configurations Is Nothing Then
        MsgBox buildConfigurationIsNothingMessage(errNumber, errMessage), vbCritical, "Handleview Framework"
        If xhvConst.DEBUG_MODE Then
            Stop
        End If
        DoCmd.Close acForm, xhvConst.APP_FORM_NAME
        End
    End If
    
    'For future use.
    'For now we only display the error in a msgbox
    'Eventually, we will display the error in an HTML page like ASP.NET App
'  If Configuration.Exists("Environment.ShowDetailFrameworkError") Then
        If Configuration("Environment.ShowDetailFrameworkError") Then
            'TODO: SHOW DETAIL ERROR IN HTML
            MsgBox "Framework error" & vbCrLf & vbCrLf & _
            "Error Number:" & errNumber & vbCrLf & _
            errMessage & vbCrLf & vbCrLf, _
            vbCritical
        Else
            MsgBox "Framework error" & vbCrLf & vbCrLf & "Error Number:" & errNumber & vbCrLf & errMessage, vbCritical, "Handleview Framework"
        End If
        
        If Not (xhvConst.DEBUG_MODE) Then
            xhvLog.LogCritical "Unmanaged Exception happened in the framework. The  application had to close.", errNumber, errMessage
            MsgBox "The application will have to close. You can contact your administrator for more help." & vbCrLf & "Sorry about the inconvenience.", vbInformation
            DoCmd.Close acForm, xhvConst.APP_FORM_NAME
            End
        End If
        
'  Else
'
'      'If this Configurations is missing, we cannot go further and just show the error message in msgbox and end the app.
'      MsgBox buildConfigurationMissingMessage(errNumber, errMessage), vbCritical, "Handleview Framework"
'      If Not (xhvConst.DEBUG_MODE) Then
'          DoCmd.Close acForm, xhvConst.APP_FORM_NAME
'          End
'      End If
'  End If




Exit Sub
HandleFrameworkException_ERR:

    'This Error cannot be shown in the web document, bacause it'll cause an infinite loop and Out Of Stack Space error
    'So for any unhandled error in this Function, this will be the error message shown.
    MsgBox "Unexpected error in xhvExceptionManager.HandleFrameworkException()" & vbCrLf & vbCrLf & _
    "Error Number:" & Err.Number & vbCrLf & _
    Err.Description & vbCrLf & vbCrLf, _
    vbCritical
    If Not xhvConst.DEBUG_MODE Then
        DoCmd.Close acForm, xhvConst.APP_FORM_NAME
        End 'End Application HERE and NOW!
    Else
        Stop
        Resume
    End If
End Sub



Private Function buildConfigurationMissingMessage(errNumber As Long, errMessage As String) As String

On Error GoTo ERR_:

    buildConfigurationMissingMessage = "Error 1001" & vbCrLf & vbCrLf & _
                                        "ShowDetailFrameworkError configuration is missing in xhvExceptionManager.HandleFrameworkException() method." & vbCrLf & vbCrLf & _
                                        "The HandleView framework cannot continue to execute the application." & vbCrLf & vbCrLf & _
                                        "The original error message calling this method was:" & vbCrLf & vbCrLf & _
                                        "Framework error" & vbCrLf & _
                                        "Error Number:" & errNumber & vbCrLf & _
                                        errMessage
Exit Function
ERR_:
    MsgBox "Fatal Error in " & MODULE_NAME & ".buildConfigurationMissingMessage()" & vbCrLf & "Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    DoCmd.Close acForm, xhvConst.APP_FORM_NAME
    End 'End Application HERE and NOW!
End Function

Private Function buildConfigurationIsNothingMessage(errNumber As Long, errMessage As String) As String

On Error GoTo ERR_:

    buildConfigurationIsNothingMessage = "Error 2017" & vbCrLf & vbCrLf & _
                                        "An error occurs before Configuration global object has been initialized." & vbCrLf & vbCrLf & _
                                        "The HandleView framework cannot continue to execute the application." & vbCrLf & vbCrLf & _
                                        "The original error message calling this method was:" & vbCrLf & vbCrLf & _
                                        "Framework error" & vbCrLf & _
                                        "Error Number: " & errNumber & vbCrLf & _
                                        "Error Message: " & errMessage
Exit Function
ERR_:
    MsgBox "Fatal Error in " & MODULE_NAME & ".buildConfigurationIsNothingMessage()" & vbCrLf & "Number:" & Err.Number & vbCrLf & Err.Description, vbCritical
    DoCmd.Close acForm, xhvConst.APP_FORM_NAME
    End 'End Application HERE and NOW!

End Function
