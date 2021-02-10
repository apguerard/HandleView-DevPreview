VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClientListController"
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
' This class is the DemoClientList component
'
Option Explicit

Const COMPONENT_NAME_TYPE As String = "DemoClientListComponent"
Const TEMPLATE_URL As String = "demoClientListView\demoClientList.view.html"

Implements xhvIController

Private this As TxhvIController
Private base As xhvBaseController
Private props As Scripting.dictionary
Private WithEvents eventDispatcher As MSHTML.HTMLUnknownElement
Attribute eventDispatcher.VB_VarHelpID = -1

Private ClientService As DemoClientService



Private Sub Class_Initialize()
    Set base = New xhvBaseController
    Set props = New Scripting.dictionary
    
    'Inject dependencies here if needed->
    Set ClientService = xhvDI.inject("IDemoUserService")
    
    this.NameType = COMPONENT_NAME_TYPE
    this.TemplateUrl = TEMPLATE_URL
    this.TemplateString = vbNullString
    this.Guid = NewGUID()
    base.LoadTemplate this
    
End Sub


Public Property Get ParentComponent() As xhvIController
    Set ParentComponent = this.ParentComponent
End Property


Private Property Set xhvIController_childComponents(value As Collection)
    Set this.ChildComponents = value
End Property
Private Property Get xhvIController_childComponents() As Collection
    Set xhvIController_childComponents = this.ChildComponents
End Property

Private Property Let xhvIController_ancestorsList(value As String)
    this.AncestorsList = value
End Property
Private Property Get xhvIController_ancestorsList() As String
    xhvIController_ancestorsList = this.AncestorsList
End Property

'ReadOnly
Private Property Get xhvIController_componentObject() As Object
    Set xhvIController_componentObject = Me
End Property

'ReadOnly
Private Property Get xhvIController_guid() As String
    xhvIController_guid = this.Guid
End Property

Private Property Get xhvIController_isUsedAsEntryPoint() As Boolean
    xhvIController_isUsedAsEntryPoint = this.IsUsedAsEntryPoint
End Property
Private Property Let xhvIController_isUsedAsEntryPoint(value As Boolean)
    this.IsUsedAsEntryPoint = value
End Property

Private Property Get xhvIController_nameType() As String
    xhvIController_nameType = this.NameType
End Property
Private Property Let xhvIController_nameType(value As String)
    this.NameType = value
End Property

Public Property Get xhvIController_ParentComponent() As xhvIController
    Set xhvIController_ParentComponent = this.ParentComponent
End Property
Public Property Set xhvIController_ParentComponent(value As xhvIController)
    Set this.ParentComponent = value
End Property

Private Property Set xhvIController_routerPort(value As xhvRouterPort)
    Set this.routerPort = value
End Property
Private Property Get xhvIController_routerPort() As xhvRouterPort
    Set xhvIController_routerPort = this.routerPort
End Property

Private Property Get xhvIController_templateString() As String
    xhvIController_templateString = this.TemplateString
End Property
Private Property Let xhvIController_templateString(value As String)
    this.TemplateString = value
End Property

'ReadOnly
Private Property Get xhvIController_templateUrl() As String
    xhvIController_templateUrl = this.TemplateUrl
End Property

Private Property Set xhvIController_view(value As MSHTML.DispHTMLGenericElement)
    Set this.view = value
End Property
Private Property Get xhvIController_view() As MSHTML.DispHTMLGenericElement
    Set xhvIController_view = this.view
End Property

Private Property Let xhvIController_wrapperElementType(value As String)
    this.WrapperElementType = value
End Property
Private Property Get xhvIController_wrapperElementType() As String
    xhvIController_wrapperElementType = this.WrapperElementType
End Property

''
' Receive every event called in the HTML Document that has a xhv-command directive
'
' @return N/A
Private Function eventDispatcher_onclick() As Boolean
    base.DispatchEvent eventDispatcher, Me
End Function

'Interface Function Override
Private Function xhvIController_interComponentCall(ParamArray agrs() As Variant) As Boolean
    base.InterComponentCall Me, agrs
End Function

Private Sub xhvIController_subscribeEventDispatcher()
    Set eventDispatcher = this.view.querySelector("#" & "eventdispatcher" & this.Guid)
    ExecuteJS vbNullString, "attachEventDispatcher('" & this.Guid & "')"
End Sub



' Edit your component below this line =>


Private Sub xhvIController_onAfterInit(attributes As Scripting.dictionary)
    
End Sub

'Interface Function Override
Private Sub xhvIController_init(params As Scripting.dictionary)
    'No params

    'Add your custom init code below =>
    '-------------------------------------
    'MsgBox "Received param:" & params("from")
    
    
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

    Set xhvIController_render = this.view
    base.interpolate Me, props
    'Add your custom render code below =>
    '-----------------------------------
    Dim demoClients As Collection
    
    Set demoClients = ClientService.GetAllClients()
    
    'Insert Users
    Dim demoClient As DemoClientModel
    Dim tableBody As MSHTML.HTMLTableSection
    Dim rowTemplate As MSHTML.HTMLTableRow
    Dim newRow As MSHTML.HTMLTableRow
    
    Set rowTemplate = this.view.querySelector("[xhv-table-row]")
    Set tableBody = this.view.querySelector("[xhv-table-body]")
    
    For Each demoClient In demoClients
        Set newRow = Document.createElement("TR")
        Set newRow = rowTemplate.cloneNode(True)
         
        newRow.setAttribute "xhv-params", demoClient.Id
        newRow.cells(0).innerText = demoClient.Id
        newRow.cells(1).innerText = demoClient.name
        
        tableBody.appendChild newRow
    Next
    
    tableBody.removeChild rowTemplate
    
    'MsgBox ClientService.getMessage()
    


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
    ' Turn off the spinner
    ExecuteJS "", "ej.popups.hideSpinner(document.getElementById('container'));"
End Sub





'Public Function showDetail(ParamArray args() As Variant) As Boolean
'  xhvRouter.navigate "home/{'id':" & args(0) & ",'name':'" & args(1) & "'}"
'End Function





Public Function showDetail(ParamArray args() As Variant) As Boolean

    
    xhvRouter.Navigate "clientDetail/{'id':" & args(0) & "}"
    
    
    
End Function

