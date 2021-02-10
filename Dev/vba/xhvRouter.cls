VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' This class is used to Navigate through the application.
'
' <attribute>
' VB_PredeclaredId = True
' </attribute>
'
Option Explicit
Private Const MODULE_NAME As String = "xhvRouter"

Private Type TxhvRouter
    rootRouterPort As xhvRouterPort
    rootComponent As xhvIController
    routes As Collection
    history As xhvNavigationHistory
End Type
Private this As TxhvRouter

Public Property Get rootComponent() As xhvIController
    Set rootComponent = this.rootComponent
End Property
Public Property Set rootComponent(value As xhvIController)
    Set this.rootComponent = value
End Property

Public Property Get rootRouterPort() As xhvRouterPort
    Set rootRouterPort = this.rootRouterPort
End Property
Public Property Set rootRouterPort(value As xhvRouterPort)
    Set this.rootRouterPort = value
End Property

Public Property Get routes() As Collection
    Set routes = this.routes
End Property
Public Property Set routes(value As Collection)
    Set this.routes = value
End Property

Private Sub Class_Initialize()
    Set this.history = New xhvNavigationHistory
End Sub

''
' Navigate the application to the provided route.
'
' @param path The string containing the requested path
' @param backingAction Boolean that tells if it is a backing action.
'             NOTE: YOU SHOULD ALWAYS LEAVE THIS PARAMETER EMPTY. IT USED IN AUTOMATIC BACKING FUNCTION
' @return N/A. RouterPort and Components are created an displayed in the UI
Public Sub Navigate(path As String, Optional backingAction As Boolean = False)

On Error GoTo ERR_
    
    Dim route As xhvRoute
    Dim configuredRoutes As Collection
    Dim routerPort As xhvRouterPort
    Dim ParentComponent As xhvIController
    Dim component As xhvIController
    
    'Turn on the spinner
    ExecuteJS "", "ej.popups.showSpinner(document.getElementById('container'));"
    
    'Wait 5 seconds
    Dim SngSec As Long
    SngSec = Timer + 5

    Do While Timer < SngSec
        DoEvents
    Loop
    
    
    'Returns route object with filled params from the requested path
    Set configuredRoutes = getConfiguredRoutesfromString(path)

    'Treat all chained configuredRoutes one after another as separated route
    For Each route In configuredRoutes
        
        'Select the RouterPort in the routerPort Tree
        'Note:The routerPorts Tree is manage in the renderer
        Set routerPort = getRouterPort(this.rootRouterPort, route.routerPortName, False)
        If routerPort Is Nothing Then
            Throw 2022, , "The RouterPort named: '" & route.routerPortName & "' does  not exist in the Router's RouterPort Tree."
        End If
               
        If Not (routerPort.component Is Nothing) Then
            'Delete the actual Port's component in ComponentTree
            removeComponentFromTree routerPort.component.Guid
            'Delete the actual Port's component reference, because it will be replace
            Set routerPort.component = Nothing
        End If
        
        'Delete children RouterPorts of the actual RouterPort
        removeChildRouterPorts routerPort
        
        'Find the parentComponent
        If Not (routerPort.parentPort Is Nothing) Then
            Set ParentComponent = routerPort.parentPort.component
        Else
            Set ParentComponent = this.rootComponent
        End If
        
        'Render the component
        Dim tmpAttributes As New Scripting.dictionary '<-- Create a temporary attributes dic for component  directly in the port that will be filled in the componentFactory
        
        Set component = xhvRenderer.RenderComponent(route.componentName, route.params, ParentComponent, tmpAttributes, routerPort, True)
        
        If component Is Nothing Then
            Throw 2023, , "The component '" & route.componentName & "' could not be rendered by the framework."
        End If
        
        If this.rootComponent Is Nothing Then
            Set this.rootComponent = component
        End If
        
        
        'MANIPULATE the DOM
        'Add the component as a child of the routerPort
        Dim node As MSHTML.IHTMLDOMNode
        Dim nextSibling As MSHTML.IHTMLDOMNode

        Set node = routerPort.DOMNodeRef.firstChild

        Do While Not (node Is Nothing)
            If node.nodeType = 1 Then
                On Error Resume Next
                If UCase(node.attributes("component").value) <> vbNullString Then
                    Set nextSibling = node
                End If
                On Error GoTo 0
                Exit Do
            Else
                Set node = node.nextSibling
            End If
        Loop

        If nextSibling Is Nothing Then
            routerPort.DOMNodeRef.InsertBefore component.view, Null
        Else
            nextSibling.ParentNode.removeChild nextSibling
            routerPort.DOMNodeRef.InsertBefore component.view, Null
        End If
        'END - Add the component as a child of the routerPort
        
        component.AfterInsertIntoDOM
        raiseAfterParentsRendered component
        
    Next route
    
    'Register all events dispatchers in  the components
    subscribeAllEventDispatchers this.rootComponent
    
Exit Sub

ERR_:

    'As it is an Main Entry Point Sub, we do not rethrow the exception.
    xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
    
    If xhvConst.DEBUG_MODE Then
        Stop
        Resume
    End If
End Sub


''
' Call the afterParentInsertedIntoDOM "event" in all the components.
'
' @param rootComponent Root component to begin the recursive method - Not necesserely the main root component
'
Private Sub raiseAfterParentsRendered(rootComponent As xhvIController)

    Dim component As xhvIController

    If Not (rootComponent Is Nothing) Then

        If Not (rootComponent.ChildComponents Is Nothing) Then
            For Each component In rootComponent.ChildComponents
                component.AfterParentInsertedIntoDOM
                'Recursive call
                raiseAfterParentsRendered component
            Next
        End If
        
    End If

End Sub

''
' Call the subscribeEventDispatcher of all components from rootComponent.
'
' @param rootComponent Root component to begin the recursive method - Nor necesserely  the main root component
'
Private Function subscribeAllEventDispatchers(rootComponent As xhvIController) As Boolean

    Dim component As xhvIController

    If Not (rootComponent Is Nothing) Then

        rootComponent.SubscribeEventDispatcher

        If Not (rootComponent.ChildComponents Is Nothing) Then
            For Each component In rootComponent.ChildComponents
                subscribeAllEventDispatchers component
            Next
        End If
        
    End If

End Function

''
' Get the route(s) from the configured routes in the xhvConfigRouter module.
' Use the matchPattern of the configured routes VS the match pattern asked path string passed in the navigate function of the router. <- Do what? :RESEARCH
'
' @param path The string containing the requested path
' @return A collection of matching configured routes or nothing if no match found.
Private Function getConfiguredRoutesfromString(path As String) As Collection

On Error GoTo ERR_

    Dim siblingPaths() As String
    Dim route As xhvRoute
    Dim i As Integer
 
    Set getConfiguredRoutesfromString = New Collection

    'Separate the siblings route
    siblingPaths = Split(path, "||")

    For i = 0 To UBound(siblingPaths)
        
        Set route = matchRoute(Trim(siblingPaths(i))) 'Return a clone of the route that we can manipulate and change
        
        If Not (route Is Nothing) Then
            'Extract the param from the json string and put it the CloneRouted Route
            Set route.params = extractParamsFromRoute(Trim(siblingPaths(i)))
            'Add the matching route
            getConfiguredRoutesfromString.Add route
        Else
            'A matching route was not found
            Dim msg As String
            msg = "The following path: '" & siblingPaths(i) & "'" & vbCrLf & _
                   "did not match any route configured in the xhvConfigRouter module." & vbCrLf & _
                   "Please check your path or your xhvConfigRouter module."
            Throw 2024, , msg
        End If
        
    Next i

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
' Extract the parameters of a Route path and put them in a dictionary
' NOTE: These parameters are in a Json format.
'
' @param path Complete path requested in the navigate function
' @return Dictionary of parameters
Private Function extractParamsFromRoute(path) As Scripting.dictionary

On Error GoTo ERR_

    Dim jsonString As String
    Dim stringArray() As String
    
    stringArray = Split(path, "/")
    If UBound(stringArray) > 0 Then
        jsonString = stringArray(1)
        Set extractParamsFromRoute = ParseJson(jsonString)
    End If

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
' Remove a specific component in the Tree component and all its child components
' Also call the dispose method of the component
'
' @param componentGuid GUID of the component to remove
'
Private Sub removeComponentFromTree(componentGuid As String)

On Error GoTo ERR_
   
    Dim component As xhvIController

    Set component = getComponentFromTree(this.rootComponent, componentGuid)

    If Not (component Is Nothing) Then
                
        'Delete all child components in  the tree
        removeChildComponents component
        
        If Not (component.ParentComponent Is Nothing) Then
            
            component.ParentComponent.ChildComponents.Remove component.Guid
            
        End If
        
        component.Dispose
             
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
' Remove all child components of a specific component in the Tree
' Also call the dispose method of the components
'
' @param parentComponent Parent component from which we want to delete the children
'
Private Sub removeChildComponents(ParentComponent As xhvIController)

On Error GoTo ERR_

    Dim component As xhvIController

    If Not (ParentComponent.ChildComponents Is Nothing) Then
        For Each component In ParentComponent.ChildComponents
            'Recursive call
            removeChildComponents component
            
            component.Dispose
            
        Next
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
' Get a specific component from the Tree
'
' @param rootComponent The root port for the recursive search. (Not necessarily the rootport of the router) - Well only on the first call
' @param componentGuid The GUID of the  component we want to find
' @return The found component or Nothing if not found
Private Function getComponentFromTree(rootComponent As xhvIController, componentGuid As String) As xhvIController

On Error GoTo ERR_

    Dim component As xhvIController
    
    If Not (rootComponent Is Nothing) Then
        If rootComponent.Guid = componentGuid Then
            Set getComponentFromTree = rootComponent
            Exit Function
        Else
            If Not (rootComponent.ChildComponents Is Nothing) Then
                For Each component In rootComponent.ChildComponents
                    Set getComponentFromTree = getComponentFromTree(component, componentGuid)
                Next
            End If
        End If
    End If


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
' Get the Port with PortName from the Component Tree of the Router
'
' @param rootComponent The root port for the recursive search.(Not necessarily the rootport of the router) - Well only on the first call
' @param portName The name of the port we want to find
' @param found A boolean tag for the recursive search
' @return The found port or Nothing if not found
Private Function getRouterPort(rootComponent As xhvRouterPort, portName As String, found As Boolean) As xhvRouterPort

On Error GoTo ERR_

    Dim routerPort As xhvRouterPort
    
    If Not (rootComponent Is Nothing) Then
        If rootComponent.name = portName Then
            Set getRouterPort = rootComponent
            found = True
            Exit Function
        Else
            For Each routerPort In rootComponent.childPorts
                Set getRouterPort = getRouterPort(routerPort, portName, False)
                If found Then Exit Function
            Next
        End If
    End If


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


Private Sub removeChildRouterPorts(ByRef routerPort As xhvRouterPort)

    Dim rp As xhvRouterPort
    
    For Each rp In routerPort.childPorts
        removeChildRouterPorts rp
        
        Set rp.component = Nothing
        Set rp.DOMNodeRef = Nothing
        Set rp.parentPort = Nothing
        Set rp.childPorts = Nothing
        Set rp = Nothing
        
    Next
    
    Set routerPort.childPorts = New Collection


End Sub


''
' Return the first matching route in the application
' Ex:
'     RequestedPath = indicatorDetail/{"id"=123}||stepDetail/{"id"=456} --> matchig path = indicatorDetail||stepDetail
'
' @param requestedPath The requested path. This path must be initialized in the xhvConfigRouter module
' @return The FIRST matching route in configured routes or Nothing
Private Function matchRoute(requestedPath As String) As xhvRoute

On Error GoTo ERR_


    Dim route As xhvRoute

    For Each route In this.routes
        If matchPattern(requestedPath) = route.matchPattern Then
            Set matchRoute = route.CloneRoute()
            Set route = Nothing
            Exit Function
        End If
    Next
    
    Set route = Nothing
    

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
' Determine the MatchPattern for a requested route.
' This MatchPattern is use to check against the MatchPattern of a configured route.
' A MatchPattern removes parameter values from the requested route
'
' @param route The requested route.
' @return The MatchPattern (string) of a requestedRoute.
Private Function matchPattern(route As String) As String

On Error GoTo ERR_


    Dim tempString As String
    Dim sChar As String
    Dim i As Integer

    For i = 1 To Len(route)
        sChar = Mid(route, i, 1)
        Select Case sChar
            Case "/"
                If Peak(route, i + 1) <> "/" Then
                    i = i + 1
                    Do While (Mid(route, i, 1) <> "/" And Mid(route, i, 1) <> "&") And i <= Len(route)
                        i = i + 1
                    Loop
                    i = i - 1
                Else
                    tempString = tempString & sChar
                    If Peak(route, i + 1) = "/" Then
                        i = i + 1
                        tempString = tempString & Mid(route, i, 1)
                    End If
                End If
            Case "&"
                tempString = tempString & sChar
                If Peak(route, i + 1) = "&" Then
                    i = i + 1
                    tempString = tempString & Mid(route, i, 1)
                End If

            Case Else
                tempString = tempString & sChar
        End Select
    Next

    matchPattern = tempString

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
' Use to get next char(s) in a string
'
' @param inString The string we want to peak.
' @param position The position we want to peak.
' @param iPeakBy The number of char we want to peak. If not provided, 1 is used by default.
' @return A single char at the position position or multiple char starting at position position with a lenth of iPeakBy
Private Function Peak(inString As String, position As Integer, Optional iPeakBy = 1) As String

    On Error GoTo ERR_

    Peak = Mid(inString, position, iPeakBy)

Finally:

Exit Function
ERR_:
    Peak = vbNullString
    GoTo Finally
End Function


