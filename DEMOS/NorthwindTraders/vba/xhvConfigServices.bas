Attribute VB_Name = "xhvConfigServices"
'@Folder lib.HandleView.Services

' Copyright (C) 2021 Blueajacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Configurations of the services used by the app.
' Add services we develop for this application in the DI container (xhvDI class)
'
Option Explicit
Private Const MODULE_NAME As String = "xhvConfigServices"

''
' Add services we develop for this application in the DI container (xhvDI class)
'
' @see xhvServiceFactory
' @return N/A

Public Sub ConfigureAppServices()
On Error GoTo ERR_


    '**************************************************
    '
    ' --->  IMPORTANT  <---
    '
    ' DON'T FORGET TO ADD YOUR SERVICE IN THE xhvServiceFactory
    '
    '**************************************************

    ' Register your services below this line -->
    '
    
    xhvDI.AddSingleton "IDemoUserService", "DemoClientService"
    
    '
    ' End register services
    
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
