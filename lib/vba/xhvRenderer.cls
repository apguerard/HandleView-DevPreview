VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvRenderer"
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
' This class is used to "pre-render" the Component - Ex: Loading the template.
' In fact, each Component is responsible for its own rendering in the DOM.
'
Option Explicit

Private Type TxhvScript
    Id As String
    mode As String
    source As String
End Type

''
' Initialize the component and calls its render function.
' This function is called from the Framework Router.
' YOU SHOULD NOT CALL THIS FUNCTION DIRECTLY IN YOUR CODE.
'
' @param componentName String containing the name of the component to render. IMPORTANT: This string must match the name  you used in the xhvComponentFactory Select Case
' @param params An array of params that is passed to the CreateComponent function in the xhvComponentFactory
' @param parentComponent The parent component object of the component.
' @param attributes The list of attributes set in the HTML template.
' @param actualRouterPort The actual RouterPort in which the component is rendered.
' @param isUsedAsEntryPoint Indicate if the component is a direct child of a RouterPort.
' @return The Component created as an Interface:xhvIController
'
Public Function RenderComponent(ComponentName As String, _
                                Params As Scripting.dictionary, _
                                ParentComponent As xhvIController, _
                                attributes As Scripting.dictionary, _
                                actualRouterPort As xhvRouterPort, _
                                IsUsedAsEntryPoint As Boolean _
                                                                    ) As xhvIController
                                                                    

On Error GoTo ERR_

    Dim Component As xhvIController
    Dim ChildComponents As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
    Dim j As Integer
    Dim attributeProperties As New Scripting.dictionary
    Dim workingAttributes As MSHTML.IHTMLAttributeCollection

    'Create and Initialize the component in the component factory
    Set Component = xhvComponentFactory.CreateComponent(ComponentName, attributes, Params, IsUsedAsEntryPoint)

    'Call the render of the concrete component
    Component.Render
    
    If actualRouterPort.Component Is Nothing Then
        'We are in the main component rendering pass
        Set actualRouterPort.Component = Component
    End If

    Dim RouterPort As xhvRouterPort
    'Collect the child routerPort(s) present in the component, create and put them in the routerPort Tree
    Set ChildComponents = Component.View.querySelectorAll(Syntax.Element.routerportElement)
    
    For i = 0 To ChildComponents.length - 1
        Set RouterPort = New xhvRouterPort
        RouterPort.Name = ChildComponents(i).getAttribute("name")
        Set RouterPort.DOMNodeRef = ChildComponents(i)
        Set RouterPort.ParentPort = actualRouterPort
        actualRouterPort.ChildPorts.Add RouterPort, RouterPort.Name
    Next i

    'Add component in the component Tree
    If ParentComponent Is Nothing Then
        'If the parent is nohting, it's should be the appComponent - the first component created
        Set ParentComponent = Component
    Else
        addChildComponent ParentComponent, Component
    End If
    
    'Set Ancestors list
    If Component.ParentComponent Is Nothing Then
        Component.AncestorsList = Component.NameType
    Else
        Component.AncestorsList = Component.ParentComponent.AncestorsList & ";" & Component.NameType
    End If
    

    'Render the children component in the actual component
    Set ChildComponents = Component.View.querySelectorAll(Syntax.Element.componentElement)

    For i = 0 To ChildComponents.length - 1

        'See <summary> of areDirectChildren function for explanation.
        If areDirectChildren(i, ChildComponents) Then

            'Grab the attributes of the component to pass them to the xhvComponentFactory.CreateComponent as props
            Set workingAttributes = ChildComponents(i).attributes
            Set attributeProperties = New Scripting.dictionary
            For j = 0 To workingAttributes.length - 1
                'Add only those that are not empty
                If workingAttributes.Item(j).Value <> vbNullString Then
                    attributeProperties.Add workingAttributes.Item(j).nodeName, workingAttributes.Item(j).Value
                End If
            Next j

            'Recursive call to the RenderComponent function - Then replace the html node with the component view into the parent component
            ChildComponents(i).ParentNode.replaceChild RenderComponent(ChildComponents(i).getAttribute(Syntax.Attr.componentTypeAttr), Params, Component, attributeProperties, actualRouterPort, False).View, ChildComponents(i)

        End If

    Next
    
    processCssDirectives Component.View
    processScriptDirectives Component.View

    Set RenderComponent = Component
    
Exit Function

ERR_:
    Set RenderComponent = Nothing
    If xhvConst.DEBUG_MODE Then
        xhvExceptionManager.HandleFrameworkException Err.Number, Err.Description
        Stop
        Resume
    Else
        ReThrow
    End If
End Function


''
' Check if a component is directly under another component in the HTML template or is inside another component.
' As the querySelectorAll method returns all component nodes in the query, we need to distinguish those that are in the next level from those farther.
' Those farther level components will be rendered in a next pass (recursive call)
'
' @param index Integer indicating which child in the following collection
' @param children MSHTML.IHTMLDOMChildrenCollection
' @return True if direct
'
Private Function areDirectChildren(index As Integer, children As MSHTML.IHTMLDOMChildrenCollection) As Boolean

    Dim i As Integer

    areDirectChildren = True

    For i = 0 To children.length - 1
        If i <> index Then
            If children(i).contains(children(index)) Then
                areDirectChildren = False
                Exit For
            End If
        End If
    Next

End Function


''
' Add port to the Port tree
'
' @param childComponent Collection of Children Ports
' @param parentComponent Pointer to the parent port
'
Private Function addChildComponent(ParentComponent As xhvIController, childComponent As xhvIController) As Boolean

    Dim parentTemp As xhvIController

    If ParentComponent.ChildComponents Is Nothing Then
        Set ParentComponent.ChildComponents = New Collection
    End If

    ParentComponent.ChildComponents.Add childComponent, childComponent.Guid
    Set childComponent.ParentComponent = ParentComponent

    Set parentTemp = Nothing

End Function


''
' Process the script directive found in the HTML template of the component, if there is one.
' This adds javascript file(s) and execute the script automatically.
' However, if the mode attribute of the script element is "load", then the script is only run on the first creation of the component.
' This function is use inside this class.
' You should not call this function directly.
'
' @param componentView A HTML Generic element that is the View of the component
' @return
' True if everything went well.
'
Private Sub processScriptDirectives(componentView As MSHTML.HTMLGenericElement)

On Error GoTo ERR_

    Dim directives As MSHTML.IHTMLDOMChildrenCollection
    Dim node As MSHTML.HTMLGenericElement
    Dim s As MSHTML.HTMLScriptElement
    Dim a As MSHTML.HTMLScriptElement
    Dim i As Integer
    Dim script As TxhvScript

    Set directives = componentView.querySelectorAll(Syntax.Element.scriptElement)

    For i = 0 To directives.length - 1
        Set node = directives.Item(i)

        With node
            script.Id = Replace(.getAttribute("url"), "\", vbNullString)
            script.mode = .getAttribute("mode")
            script.source = .getAttribute("url")
        End With

        'If mode=load: we want to add the script only once in the lifetime of the application
        'If mode=run: we want to add the script everytime the component is instanciated
        If script.mode = "load" Then
            If Document.querySelectorAll("[id='" & script.Id & "']").length = 0 Then
                'Add the script (run automatically when loaded)
                Set s = Document.createElement("script")
                s.Id = script.Id
                s.src = script.source
                s.Type = "text/javascript"
                Document.getElementsByTagName("head")(0).appendChild s
            End If
        Else
            'Remove the script from memory if it is already loaded
            If Document.querySelectorAll("[id='" & script.Id & "']").length <> 0 Then
                Set a = Document.querySelector("[id='" & script.Id & "']")
                a.ParentNode.removeChild a
            End If
            'Add the script (run automatically when loaded)
            Set s = Document.createElement("script")
            s.Id = script.Id
            s.src = script.source
            s.Type = "text/javascript"
            Document.getElementsByTagName("head")(0).appendChild s
        End If

    Next

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
' Process the CSS directive found in the HTML template of the component, if there is one.
' If the CSS file is already loaded, then it will be replaced and the browser will automacally apply the new CSS.
' This function is used inside this class.
' You should not call this function directly.
'
' @param poDiv A HTML Generic element that is the View of the component
' @return
' True if everything went well.
'

Private Sub processCssDirectives(poDiv As MSHTML.HTMLGenericElement)

On Error GoTo ERR_


    Dim directives As MSHTML.IHTMLDOMChildrenCollection
    Dim node As MSHTML.HTMLGenericElement
    Dim s As MSHTML.HTMLLinkElement
    Dim a As MSHTML.HTMLLinkElement
    Dim i As Integer
    Dim source As String
    Dim Id As String
    Dim Uid As String

    Set directives = poDiv.querySelectorAll(Syntax.Element.cssElement)

    Uid = NewGUID()

    For i = 0 To directives.length - 1
        Set node = directives.Item(i)

        source = node.getAttribute("url")
        Id = Replace(node.getAttribute("url"), "\", vbNullString)

        'Check if the css is already loaded
        If Document.querySelectorAll("[id='" & Id & "']").length <> 0 Then
            'If yes, replace it
            Set a = Document.querySelector("[id='" & Id & "']")

            Set s = Document.createElement("link")
            s.Id = Id
            s.href = node.getAttribute("url") & "?version" & CStr(Uid)
            s.rel = "stylesheet"
            s.Type = "text/css"

            a.replaceNode s
        Else
            'Add the css (apply automatically when loaded)
            Set s = Document.createElement("link")
            s.Id = Id
            s.href = node.getAttribute("url") & "?version" & CStr(Uid)
            s.rel = "stylesheet"
            s.Type = "text/css"
            Document.getElementsByTagName("head")(0).appendChild s
        End If

    Next

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


