Attribute VB_Name = "xhvStartupApp"
'@Folder lib.HandleView.Config

' Copyright (C) 2021 Blueajacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Gu�rard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Contains Application Startup functions for HandleView Framework.
Option Explicit

Private Const MODULE_NAME As String = "xhvStartupApp"


''
' This function build the Application for the HandleView framework
' Here we configure:
'  - Services
'  - App configuration (getting them)
'  - Logging
'  - Routes
'  - Router - By building the root RouterPort
'
' @param WB The webbrowser control that should be in the App Form
'
Public Sub StartupApp(WB As WebBrowserControl)
On Error GoTo ERR_

    ConfigureAppServices
    
    ConfigureLoggingServices
    
    ConfigureRoutes

    'Request, load and wait for the App Host HTML file to be ready (Base htmlfile)
    'The required format is : ="Path" (with the ")
    WB.ControlSource = "=""" & CurrentProject.Path & (Configuration("App.StartupHostPath") & Configuration("App.StartupHostFile")) & """"
    Do While WB.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    
    BuildRootRouterPort
    
    'Startup the application UI
    'This route or these routes must be configured in the xhvConfigRouter.
    xhvRouter.Navigate Configuration("App.StartupRoute")


    'Keep or not open BackEnd DB connection
    If Configuration("App.UseBackEndPattern") And Configuration("App.KeepBackEndOpen") Then
        OpenAllDatabases True
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
' Create and initialize the Root RouterPort in the Router
' The name of this port must be equal to App.RootRouterPort in the App configuration
'
Public Sub BuildRootRouterPort()
On Error GoTo ERR_

    Set xhvRouter.rootRouterPort = New xhvRouterPort
    
    With xhvRouter.rootRouterPort
        .Name = Configuration("App.RootRouterPort")
        If IsNull(Document.documentElement.querySelector(Syntax.Element.routerportElement & "[name='" & Configuration("App.RootRouterPort") & "']")) Then
            Throw 2001, Err.source, "The <" & Syntax.Element.routerportElement & "> name=""" & Configuration("App.RootRouterPort") & " </" & Syntax.Element.routerportElement & "> code cannot be found in the HostFile."
        Else
            Set .DOMNodeRef = Document.documentElement.querySelector(Syntax.Element.routerportElement & "[name='" & Configuration("App.RootRouterPort") & "']")
        End If
    End With

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
