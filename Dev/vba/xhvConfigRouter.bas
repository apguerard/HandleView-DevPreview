Attribute VB_Name = "xhvConfigRouter"
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
' This module is used to configure all Routes for the application in a centralized place.
'

Option Explicit
Private Const MODULE_NAME As String = "xhvConfigRouter"

''
' Configure all Routes for Application.
'
' @remarks ---> This Function is called in the xhvAppStartup.StartupApp() <---
' @return A collection of Routes

Public Sub ConfigureRoutes()

    Dim routes As Collection
    Dim route As xhvRoute

    Set routes = New Collection

    'Typical Basic App route
    Set route = New xhvRoute
    route.path = "app"
    route.routerPortName = "app-root"
    route.componentName = "AppComponent"
    route.exitGate = vbNullString
    route.securityGate = vbNullString
    routes.Add route


    'Add your own routes below -->
    '-----------------------------

  Set route = New xhvRoute
  route.path = "home"
  route.routerPortName = "content"
  route.componentName = "HomeComponent"
  route.exitGate = vbNullString
  route.securityGate = vbNullString
  routes.Add route

  Set route = New xhvRoute
  route.path = "list"
  route.routerPortName = "content"
  route.componentName = "DemoClientListComponent"
  route.exitGate = vbNullString
  route.securityGate = vbNullString
  routes.Add route

  Set route = New xhvRoute
  route.path = "clientDetail"
  route.routerPortName = "clientDetail"
  route.componentName = "DemoClientDetailComponent"
  route.exitGate = vbNullString
  route.securityGate = vbNullString
  routes.Add route

  Set route = New xhvRoute
  route.path = "buttonExample"
  route.routerPortName = "content"
  route.componentName = "ButtonExampleComponent"
  route.exitGate = vbNullString
  route.securityGate = vbNullString
  routes.Add route

  Set route = New xhvRoute
  route.path = "manageFormExample"
  route.routerPortName = "content"
  route.componentName = "ManageFormComponent"
  route.exitGate = vbNullString
  route.securityGate = vbNullString
  routes.Add route

    Set route = Nothing
    Set xhvRouter.routes = routes

End Sub

