VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvBaseController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Controllers

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Contains the base methods that each component needs.
' Used in each concrete component class.
'
Option Explicit

Private Const MODULE_NAME As String = "xhvBaseComponent"

''
' Loads the html template file in the component view property
'
' @param concreteComponent The concrete implementation of a component
'
Public Sub LoadTemplate(ByRef concreteComponent As TxhvIController)

On Error GoTo ERR_

    Dim tempElement As MSHTML.HTMLUnknownElement
    Dim wrapperElement As MSHTML.HTMLUnknownElement

    'Loading Template in Rendering node

    Set tempElement = Document.createElement("div")
    If Trim(concreteComponent.TemplateString) = vbNullString Then
        'This won't happen if the UseComponentsCacheManagement configuration is True, even if the TemplateString is not = to vbNullString because:
        'If UseComponentsCacheManagement configuration is True, then this part is directly done in the xhvComponentFactory before and the TemplateString is already filled with the file's content
        tempElement.innerHTML = Trim(readTemplateFromFile(concreteComponent.TemplateUrl))
    Else
        tempElement.innerHTML = concreteComponent.TemplateString
    End If

    'The tempElement firstchild now contains the HTML in the file. Assign it to our wrapper Element and get rid of the temporary div node.
    Set wrapperElement = tempElement.firstChild
    Set tempElement = Nothing
    
    'Give the component a unique ID  and add a component attribute
    wrapperElement.setAttribute "id", Syntax.Element.componentWrapperElement & concreteComponent.Guid
    wrapperElement.setAttribute Syntax.Element.componentWrapperElement, concreteComponent.NameType

    'Create and insert the eventDispatcher node, give it a unique id and insert it in the component
    Dim eventDispatcherNode As MSHTML.IHTMLElement
    Set eventDispatcherNode = Document.createElement("eventdispatcher")
    eventDispatcherNode.setAttribute "id", "eventdispatcher" & concreteComponent.Guid
    eventDispatcherNode.Style.display = "none"
    wrapperElement.insertAdjacentElement "afterbegin", eventDispatcherNode
       
    'Add and initialize all the event listeners in the component. NOTE: Event listeners have a 'xhv-eventlistener' attribute
    Dim ChildComponents As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
      
    Set ChildComponents = wrapperElement.querySelectorAll("[" & Syntax.Attr.eventListenerAttr & "]")
    
    For i = 0 To ChildComponents.length - 1
        ChildComponents(i).setAttribute Syntax.Attr.eventListenerAttr, concreteComponent.Guid
    Next

    'Return the  View
    Set concreteComponent.View = wrapperElement
   

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
' Dispatch the Event that happened in the HTML document to the right event handler in VBA.
' The event is raised in the concrete implementation of the component via its eventDispatcher_onclick() function and then pass to this Sub
' This Sub is responsible for calling the the right method in the concrete component
' and automaticaly go back in the component's parent tree to call the method if the right method is not found in the actual component.
'
' @param eventDispatcher he HTML element that raised the event. (Should be the elements with a xhv-eventlistener attribute in the template file.)
' @param componentObject The concrete component that  called this sub.
'
Public Sub DispatchEvent(eventDispatcher As MSHTML.HTMLUnknownElement, ComponentObject As Object)

On Error GoTo ERR_

    Dim tempString() As String
    Dim method As String
    Dim Params As String
    Dim paramsArray()
    Dim i As Integer

    'Get the event handler to call (Remove any unwanted space in the method's name)
    method = Replace(eventDispatcher.getAttribute(Syntax.Attr.eventHandlerAttr), " ", vbNullString)
    Params = eventDispatcher.getAttribute(Syntax.Attr.paramsAttr)
    
    'Get and separate the parameters values
    tempString = Split(Params, ";")
    If (UBound(tempString)) <> -1 Then
        ReDim paramsArray(UBound(tempString))
        For i = 0 To UBound(tempString)
            paramsArray(i) = tempString(i)
        Next
    Else
        ReDim paramsArray(0)
    End If
       
     
    'Call the method in the component
    'If parent method is called, check only    in parent
    'If child method is called, check all child
    'If direct method is calles, check in the component and then in all parents
    
    If Left(method, "7") = "parent." Then
    
        Dim parentParams(1) As Variant
        Dim Interface As xhvIController
        Dim interface2 As xhvIController
        Dim methodName As String
        
        methodName = Right(method, Len(method) - 7)
        
        parentParams(0) = methodName
        parentParams(1) = paramsArray
        
        If Not (ComponentObject.ParentComponent Is Nothing) Then
            Set Interface = ComponentObject.ParentComponent.ComponentObject

            On Error Resume Next
                CallByNameXHV Interface, "interComponentCall", parentParams
                If Err.Number <> 0 Then
                    
                    Set interface2 = ComponentObject
                    On Error GoTo ERR_
                    Throw 2005, Err.source, "The method '" & methodName & "' cannot be found in the parent component of '" & interface2.NameType & "' which is '" & Interface.NameType & "'."
                End If
            On Error GoTo ERR_
        Else
            Set interface2 = ComponentObject
            On Error GoTo ERR_
            Throw 2006, Err.source, "You're trying to call the method '" & methodName & "' in the parent component of '" & interface2.NameType & "' , but this component is a root component, so it does not have a parent."
        End If
        
    ElseIf Left(method, "6") = "child." Then
        'TODO: Call a method in a child component. Which one? All? All descendants?
    Else
        On Error Resume Next
            CallByNameXHV ComponentObject, method, paramsArray
            If Err.Number <> 0 Then
                'Retry to call the method on the parent component until it reaches the root parent
                If Not (ComponentObject.ParentComponent Is Nothing) Then
                    DispatchEvent eventDispatcher, ComponentObject.ParentComponent.ComponentObject
                Else
                'Throw an error if no method have been found in all the parent
                On Error GoTo ERR_
                Throw 2007, Err.source, "The method: '" & method & "' have not been found in the component or any parent of the component or it raise an error."
                End If
            End If
        On Error GoTo ERR_
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
' Calls a method in another component. This Sub in used in the DispatchEvent Sub when calling parent or child method.
'
' @param component The concrete component for which we want to call a method.
' @param paramArray An array of parameters for the specific methods.
'
Public Sub InterComponentCall(Component As Object, ParamArray args() As Variant)
   
    Dim method As String
    Dim paramsArray()
    method = args(0)(0)
    paramsArray = args(0)(1)

    CallByNameXHV Component, method, paramsArray
    
End Sub

''
' Replace the interpolation in a concrete component with propery value or function return value. > interpolation  = {{propertyname}}
'
' @param component The concrete component in which we want to interpolate e.g. replace values
' @param props A dictionary containing all properties of the components..
'
Public Sub interpolate(Component As xhvIController, props As Scripting.dictionary)

On Error GoTo ERR_

    Dim vKey As Variant
    Dim interpolations As Collection
    Dim interp As xhvInterpolation
    Dim valueString() As String
    
    Set interpolations = getInterpolations(Component.View.innerHTML)
    
    For Each interp In interpolations
    
        Select Case interp.TypeInterpolation
            Case "props"
                'Check if we call a scalar property or a property of an object property
                valueString = Split(interp.Value, ".")
                If UBound(valueString) = 0 Then
                    'Scalar property
                    If props.Exists(interp.Value) Then
                        Component.View.innerHTML = Replace(Component.View.innerHTML, "{{props:" & interp.Value & "}}", props.Item(interp.Value))
                    Else
                        Throw 2008, Err.source, "The interpolation: '" & interp.Value & "' cannot be resolved by the framework."
                    End If
                ElseIf UBound(valueString) = 1 Then
                    'Object property
                    If props.Exists(valueString(0)) Then
                        Component.View.innerHTML = Replace(Component.View.innerHTML, "{{props:" & interp.Value & "}}", CallByName(props.Item(valueString(0)), valueString(1), VbGet))
                    Else
                        Throw 2008, Err.source, "The interpolation: '" & interp.Value & "' cannot be resolved by the framework."
                    End If
                Else
                    'More than one level of reference
                    'The framework do not support calling an object property of an object property
                    Throw 2009, Err.source, "The interpolation: '" & interp.Value & "' cannot be resolved by the framework." & vbCrLf & vbCrLf & "Handleview do not support more than one(1) level of object property for now."
                End If
            Case "func"
                valueString = Split(interp.Value, ".")
                If UBound(valueString) = 0 Then
                    On Error Resume Next
                        Component.View.innerHTML = Replace(Component.View.innerHTML, "{{func:" & interp.Value & "}}", CallByName(Component.ComponentObject, valueString(0), VbMethod))
                        If Err.Number <> 0 Then
                            On Error GoTo ERR_
                            Throw 2010, , "The component : " & Component.NameType & " cannot resolve the 'func' interpolation: " & valueString(0) & vbCrLf & vbCrLf & "This component probably dont't have this function or this function generate an unhandled error."
                        End If
                    On Error GoTo ERR_
                Else
                    'More than one level of reference
                    'The framework do not support calling an object method of an object property
                    Throw 2011, Err.source, "The interpolation: " & interp.Value & " cannot be resolved by the framework." & vbCrLf & vbCrLf & "Handleview do not support call to method of an object property for now."
                    'TODO: Next version add this feature
                End If
            Case Else
                Throw 2012, Err.source, "Interpolation type does not exist: " & interp.TypeInterpolation
        End Select
    
    Next
    
    Set interpolations = Nothing
    Set interp = Nothing
    
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
' Get a collection of all interpolations that exist in the component HTML Template > interpolation  = {{propertyname}}
' This function is used in interpolate() Sub
'
' @param view The string representing the HTML template of the concrete component
' @return A collection returning all interpolations.
Private Function getInterpolations(View As String) As Collection

On Error GoTo ERR_


    Dim Stream As xhvStream
    Dim myChar As String
    Dim tempString As String
    Dim tempString2() As String
    Dim interp As xhvInterpolation
    
    Set getInterpolations = New Collection
    Set Stream = New xhvStream
    
    'Use a xhvStream object and run through all characters in the string
    Stream.Stream = View
    
    myChar = Stream.MoveNext
    
    Do While myChar <> vbNullString

        If myChar = "{" Then
            'Check if it is indeed an interpolation in looking for the next char
            If Stream.peak(1) = "{" Then
                Set interp = New xhvInterpolation
                Stream.MoveNext
                tempString = vbNullString
                myChar = Stream.MoveNext
                Do While myChar <> vbNullString
                    If myChar = "}" And Stream.peak() = "}" Then Exit Do
                    tempString = tempString & myChar
                    myChar = Stream.MoveNext()
                Loop
                
                If InStr(1, tempString, ":") = 0 Then
                    Throw 2003, Err.source, "The following interpolation does not have the right format: " & tempString & vbCrLf & vbCrLf & "Expecting 'type:name' format."
                End If
                
                tempString2 = Split(tempString, ":")
                
                interp.TypeInterpolation = tempString2(0) 'Is it a props or a func call
                interp.Value = tempString2(1)
                
                getInterpolations.Add interp
                
            End If
        
        End If
            
        myChar = Stream.MoveNext
        
    Loop
    
    Set Stream = Nothing
    Set interp = Nothing
    
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
' Reads the html template from file
'
' @param templateFileNam HTML Template file name (ex: mycomponent.component.html). This file name is set in the concerete component class.
' @return String: content of the file
'
Private Function readTemplateFromFile(templatefileName As String) As String

On Error GoTo ERR_

        Dim adoStream As ADODB.Stream
        Dim pathFileName As String
               
        pathFileName = CurrentProject.Path & Configuration("App.BaseViewsFolder") & templatefileName
        
        If Not (FileExists(pathFileName)) Then
            Throw 2004, Err.source, "The component HTML file template could not be found." & vbCrLf & "File : " & pathFileName
        End If
        
        Set adoStream = New ADODB.Stream
        adoStream.Charset = "UTF-8"

        adoStream.Open
        adoStream.LoadFromFile pathFileName
        readTemplateFromFile = adoStream.ReadText()
        
        adoStream.Close
        Set adoStream = Nothing

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

