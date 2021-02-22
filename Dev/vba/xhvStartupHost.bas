Attribute VB_Name = "xhvStartupHost"
'@Folder lib.HandleView.Config

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Contains Host Startup functionnalities for HandleView Framework
'
Option Explicit

Private Const MODULE_NAME As String = "xhvStartupHost"


''
' This Sub configures the HandleView framework
' Here we configure:
'    - Syntax used by the framework for HTML Templating
'    - Framework (Host) configuration (getting them)
'
' @param appDocument The Document object of the Browser
'
Public Sub StartupHost(appDocument As MSHTML.HTMLDocument)
On Error GoTo ERR_

    Set Document = appDocument
    
    SetFrameworkSemantics ' <--- Set the Syntax for the framework <-- DO NOT  REMOVE
    Set HV = xhv ' <--- This line calls the Class_Initialize() of the xhv Class which is static <--  DO NOT REMOVE
    
    'Get configuration (from our custum source if wanted e.g. Json file, other DB, etc...)
    xhvConfigurator.AddLocalDB xhvConst.APP_CONFIG_TABLE_NAME, xhvConst.APP_CONFIG_ID_FIELD, xhvConst.APP_CONFIG_VALUE_FIELD  ' <--- Can be changed if we use other source for configuration
    

Exit Sub

ERR_:
    'If error hapens in this sub, configuration would probably not have been set. So we show a msgbox instead of using xhvExceptionManager
    MsgBox "Unexpected error in " & MODULE_NAME & ".StartupHost()" & vbCrLf & vbCrLf & "Please contact your administrator." & vbCrLf & vbCrLf & "The application will now close.", vbCritical
    If xhvConst.DEBUG_MODE Then
        Stop
        Resume
    Else
        'Error have been handled but app cannot continue.
        DoCmd.Close acForm, xhvConst.APP_FORM_NAME
        End
    End If
End Sub
