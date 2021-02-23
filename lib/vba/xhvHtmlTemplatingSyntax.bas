Attribute VB_Name = "xhvHtmlTemplatingSyntax"
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
' This module contains the type def for the HandleView Syntax of HTML templates
'

Private Const MODULE_NAME As String = "xhvHtmlTemplateSyntax"

Global Syntax As TxhvSyntax

''
' Custom HTML elements
'
Public Type TxhvHTMLElement
    xhvElement As String                'Any element starting with the prefix defined below will be parsable by the Framework
    routerportElement As String         'An element used as a place to mount child components that receive navigation
    componentElement As String          'An element used as a place to mount child components as defined below
    eventhandlerElement As String       'The element that defines which VBA method to call as defined below
    componentWrapperElement As String   'The parent element as defined below
    scriptElement As String             'Javascript script tag element as defined below
    cssElement As String                'CSS tag element as defined below
End Type

''
' Custom HTML attributes
'
Public Type TxhvHTMLAttribute
    eventAttr As String                 'Event attributes define which event to respond to as defined below
    paramsAttr As String                'This will hold the definition of the handler parameters attribute as defined below
    eventHandlerAttr As String          'Attribute that will hold the name of the tag that will define the event handler method to call as defined below
    eventListenerAttr As String         'This will hold the definition of the event listener attribute which denotes the beginning of an event handler as defined below
    componentTypeAttr As String         'Defines the tag name to represent the component type data
End Type

''
' Syntax
'
Public Type TxhvSyntax
    Element As TxhvHTMLElement
    Attr As TxhvHTMLAttribute
End Type


''
' This function initializes the HandleView Framework Base HTML Sementics used in the Framework code
'
Public Sub SetFrameworkSemantics()
 
    Syntax.Element.routerportElement = "xhv-routerport"         'Defines 'xhv-routerport' as the routerportElement (Child components mount here and receive navigation)
    Syntax.Element.xhvElement = "xhv"                           'Defines 'xhv' as the prefix to identify HandleView elements
    Syntax.Element.componentElement = "xhv-component"           'Defines 'xhv-component' as the tag to identify a HandleView component
    Syntax.Element.eventhandlerElement = "xhv-eventhandler"     'Defines 'xhv-eventhandler' as the tag used to identify the event handler to call :RESEARCH
    Syntax.Element.componentWrapperElement = "component"        'Defines 'component' tag as a wrapper tag for a given child component
    Syntax.Element.scriptElement = "xhv-script"                 'Defines 'xhv-script' as a named function that can be called
    Syntax.Element.cssElement = "xhv-css"                       'Defines 'xhv-css' as a css tag that HandleView can manage
    
        
    Syntax.Attr.eventListenerAttr = "xhv-eventlistener"         'Defines 'xhv-eventlistner' as the start of an event handler in html
    Syntax.Attr.eventAttr = "xhv-event"                         'Defines 'xhv-event' as the tag to identify which event to subscribe to (currently only click)
    Syntax.Attr.eventHandlerAttr = "xhv-eventhandler"           'Defines 'xhv-eventhandler' as the tag to identify which method to call when the event is raised
    Syntax.Attr.paramsAttr = "xhv-params"                       'Defines 'xhv-params' as the tag to hold Parameters passed to the handling script
    Syntax.Attr.componentTypeAttr = "xhvtype"                   'Holds component Type data:  @See xhvType.bas

End Sub
