VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ComponentTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'*****************************************************************************
' <summary>
'   This class is the Application component
' </summary>
' <attribute>
'   VB_PredeclaredId = False
' </attribute>
'*****************************************************************************
Option Explicit

Const COMPONENT_NAME_TYPE As String = "componentName"
Const TEMPLATE_URL As String = "componentName\componentName.component.html"

Implements xhvIComponent

Private this As TxhvIComponent
Private base As xhvBaseComponent
Private props As Scripting.Dictionary
Private WithEvents eventDispatcher As MSHTML.HTMLUnknownElement
Attribute eventDispatcher.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set base = New xhvBaseComponent
    Set props = New Scripting.Dictionary
    
    this.nameType = COMPONENT_NAME_TYPE
    this.templateUrl = TEMPLATE_URL
    this.templateString = ""
    this.guid = NewGUID()
    base.loadTemplate this
    
End Sub
'-------------------------------------------------------------------------------

Public Property Get parentComponent() As xhvIComponent
    Set parentComponent = this.parentComponent
End Property
'-------------------------------------------------------------------------------

Private Property Set xhvIComponent_childrenComponents(value As Collection)
    Set this.childrenComponents = value
End Property
Private Property Get xhvIComponent_childrenComponents() As Collection
    Set xhvIComponent_childrenComponents = this.childrenComponents
End Property

Private Property Let xhvIComponent_AncestorsList(value As String)
    this.ancestorsList = value
End Property
Private Property Get xhvIComponent_AncestorsList() As String
    xhvIComponent_AncestorsList = this.ancestorsList
End Property

'ReadOnly
Private Property Get xhvIComponent_componentObject() As Object
    Set xhvIComponent_componentObject = Me
End Property


'ReadOnly
Private Property Get xhvIComponent_Guid() As String
    xhvIComponent_Guid = this.guid
End Property

Private Property Get xhvIComponent_isUsedAsEntryPoint() As Boolean
    xhvIComponent_isUsedAsEntryPoint = this.isUsedAsEntryPoint
End Property
Private Property Let xhvIComponent_isUsedAsEntryPoint(value As Boolean)
    this.isUsedAsEntryPoint = value
End Property

Private Property Get xhvIComponent_nameType() As String
    xhvIComponent_nameType = this.nameType
End Property
Private Property Let xhvIComponent_nameType(value As String)
    this.nameType = value
End Property

Public Property Get xhvIComponent_parentComponent() As xhvIComponent
    Set xhvIComponent_parentComponent = this.parentComponent
End Property
Public Property Set xhvIComponent_parentComponent(value As xhvIComponent)
    Set this.parentComponent = value
End Property

Private Property Set xhvIComponent_routerPort(value As xhvRouterPort)
    Set this.routerPort = value
End Property
Private Property Get xhvIComponent_routerPort() As xhvRouterPort
    Set xhvIComponent_routerPort = this.routerPort
End Property

Private Property Get xhvIComponent_TemplateString() As String
    xhvIComponent_TemplateString = this.templateString
End Property
Private Property Let xhvIComponent_TemplateString(value As String)
    this.templateString = value
End Property

'ReadOnly
Private Property Get xhvIComponent_TemplateUrl() As String
    xhvIComponent_TemplateUrl = this.templateUrl
End Property

Private Property Set xhvIComponent_view(value As MSHTML.DispHTMLGenericElement)
    Set this.view = value
End Property
Private Property Get xhvIComponent_view() As MSHTML.DispHTMLGenericElement
    Set xhvIComponent_view = this.view
End Property

Private Property Let xhvIComponent_wrapperElementType(value As String)
    this.wrapperElementType = value
End Property
Private Property Get xhvIComponent_wrapperElementType() As String
    xhvIComponent_wrapperElementType = this.wrapperElementType
End Property


' <summary>
'   Receive EVERY event called in the HTML Document that have a xhv-eventlistener attribute
' </summary>
' <returns>N/A</returns>
Private Function eventDispatcher_onclick() As Boolean
    base.dispatchEvent eventDispatcher, Me
End Function

'Interface Function Override
Private Function xhvIComponent_interComponentCall(ParamArray agrs() As Variant) As Boolean
    base.interComponentCall Me, agrs
End Function

Private Sub xhvIComponent_subscribeEventDispatcher()
    Set eventDispatcher = this.view.querySelector("#" & "eventdispatcher" & this.guid)
    ExecuteJS "", "attachEventDispatcher('" & this.guid & "')"
End Sub

'================================================================================================================================================
' Edit your component below this line =>
'================================================================================================================================================

Private Sub xhvIComponent_onAfterInit(oAttributes As Scripting.Dictionary)

End Sub


'Interface Function Override
Private Sub xhvIComponent_Init(oParams As Scripting.Dictionary)
    'No params

    'Add your custom init code below =>
    '-------------------------------------

    
    
End Sub
'Interface Function Override
Private Sub xhvIComponent_dispose()
    Set base = Nothing
    Set props = Nothing
    'Add your custom dispose code below =>
    '--------------------------------------

End Sub

'Interface Function Override
Private Function xhvIComponent_render() As MSHTML.HTMLGenericElement

    Set xhvIComponent_render = this.view
    base.interpolate Me, props
    'Add your custom render code below =>
    '-----------------------------------
    

End Function

'Interface Function Override
Private Sub xhvIComponent_onChildrenPortRendered()

End Sub
'Interface Function Override
Private Sub xhvIComponent_refresh()

End Sub
'Interface Function Override
Private Sub xhvIComponent_afterParentInsertedIntoDOM()

End Sub
'Interface Function Override()
Private Sub xhvIComponent_afterInsertIntoDOM()

End Sub




Public Function test(ParamArray args() As Variant) As Boolean

    MsgBox "Allo From " & COMPONENT_NAME_TYPE

End Function
