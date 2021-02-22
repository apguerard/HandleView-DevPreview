VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Controllers

' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class is a button component
'
Option Explicit

Const COMPONENT_NAME_TYPE As String = "ButtonComponent"
Const TEMPLATE_URL As String = "Shared\ButtonView\button.view.html"
Const WRAPPER_ELEMENT_TYPE As String = "button"

Implements xhvIController

Private this As TxhvIController
Private base As xhvBaseController
Private props As Scripting.dictionary
Private WithEvents eventDispatcher As MSHTML.HTMLUnknownElement
Attribute eventDispatcher.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set base = New xhvBaseController
    Set props = New Scripting.dictionary
    
    this.NameType = COMPONENT_NAME_TYPE
    this.TemplateUrl = TEMPLATE_URL
    this.WrapperElementType = WRAPPER_ELEMENT_TYPE
    this.TemplateString = vbNullString
    this.Guid = NewGUID()
    base.LoadTemplate this
End Sub

Public Property Get ParentComponent() As xhvIController
    Set ParentComponent = this.ParentComponent
End Property

Private Property Set xhvIController_childComponents(Value As Collection)
    Set this.ChildComponents = Value
End Property

Private Property Get xhvIController_childComponents() As Collection
    Set xhvIController_childComponents = this.ChildComponents
End Property

Private Property Let xhvIController_ancestorsList(Value As String)
    this.AncestorsList = Value
End Property

Private Property Get xhvIController_ancestorsList() As String
    xhvIController_ancestorsList = this.AncestorsList
End Property

Private Property Get xhvIController_componentObject() As Object
    Set xhvIController_componentObject = Me
End Property

Private Property Get xhvIController_guid() As String
    xhvIController_guid = this.Guid
End Property

Private Property Get xhvIController_isUsedAsEntryPoint() As Boolean
    xhvIController_isUsedAsEntryPoint = this.IsUsedAsEntryPoint
End Property

Private Property Let xhvIController_isUsedAsEntryPoint(Value As Boolean)
    this.IsUsedAsEntryPoint = Value
End Property

Private Property Get xhvIController_nameType() As String
    xhvIController_nameType = this.NameType
End Property

Private Property Let xhvIController_nameType(Value As String)
    this.NameType = Value
End Property

Public Property Get xhvIController_ParentComponent() As xhvIController
    Set xhvIController_ParentComponent = this.ParentComponent
End Property

Public Property Set xhvIController_ParentComponent(Value As xhvIController)
    Set this.ParentComponent = Value
End Property

Private Property Set xhvIController_routerPort(Value As xhvRouterPort)
    Set this.RouterPort = Value
End Property

Private Property Get xhvIController_routerPort() As xhvRouterPort
    Set xhvIController_routerPort = this.RouterPort
End Property

Private Property Get xhvIController_templateString() As String
    xhvIController_templateString = this.TemplateString
End Property

Private Property Let xhvIController_templateString(Value As String)
    this.TemplateString = Value
End Property

Private Property Get xhvIController_templateUrl() As String
    xhvIController_templateUrl = this.TemplateUrl
End Property

Private Property Set xhvIController_view(Value As MSHTML.DispHTMLGenericElement)
    Set this.View = Value
End Property

Private Property Get xhvIController_view() As MSHTML.DispHTMLGenericElement
    Set xhvIController_view = this.View
End Property

Private Property Let xhvIController_wrapperElementType(Value As String)
    this.WrapperElementType = Value
End Property

Private Property Get xhvIController_wrapperElementType() As String
    xhvIController_wrapperElementType = this.WrapperElementType
End Property

''
' Receive every event called in the HTML Document that has a xhv-command directive
'
Private Function eventDispatcher_onclick() As Boolean
    base.DispatchEvent eventDispatcher, Me
End Function

'Interface Function Override
Private Function xhvIController_interComponentCall(ParamArray agrs() As Variant) As Boolean
    base.InterComponentCall Me, agrs
End Function

Private Sub xhvIController_subscribeEventDispatcher()
    Set eventDispatcher = this.View.querySelector("#" & "eventdispatcher" & this.Guid)
    ExecuteJS vbNullString, "attachEventDispatcher('" & this.Guid & "')"
End Sub


' Edit your component below this line =>


Private Sub xhvIController_onAfterInit(attributes As Scripting.dictionary)

    If Not (this.IsUsedAsEntryPoint) Then
        props.Add "class", attributes("class")
        props.Add "name", attributes("name")
        props.Add "action", attributes("action")
        props.Add "params", attributes("params")
        props.Add "buttonId", attributes("buttonid")
    End If
    
End Sub



'Interface Function Override
Private Sub xhvIController_init(Params As Scripting.dictionary)
    'No params
    
    'Add your custom init code below =>
    '-------------------------------------
    If this.IsUsedAsEntryPoint Then
        props.Add "class", "btn btn-danger"
        props.Add "name", "DANGER"
        props.Add "action", "test"
        props.Add "params", vbNullString
        props.Add "buttonid", "no-id"
    End If
    
    
End Sub
'Interface Function Override
Private Sub xhvIController_dispose()
    Set base = Nothing
    Set props = Nothing
    'Add your custom dispose code below =>
    '--------------------------------------

End Sub

'Interface Function Override
Private Function xhvIController_render() As MSHTML.HTMLGenericElement

    Set xhvIController_render = this.View
    base.interpolate Me, props
    'Add your custom render code below =>
    '-----------------------------------
    

End Function

'Interface Function Override
Private Sub xhvIController_onChildrenPortRendered()

End Sub
'Interface Function Override
Private Sub xhvIController_refresh()

End Sub
'Interface Function Override
Private Sub xhvIController_afterParentInsertedIntoDOM()

End Sub
'Interface Function Override
Private Sub xhvIController_afterInsertIntoDOM()

End Sub

Public Function HelloFromComponent(ParamArray args() As Variant) As Boolean

    MsgBox "Hello from VBA component" & vbCrLf & "Button HTML id:" & props("buttonId"), vbInformation

End Function

