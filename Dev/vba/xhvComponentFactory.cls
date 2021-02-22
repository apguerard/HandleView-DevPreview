VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvComponentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' This factory class is responsible for creating and initializing components in the framework.
' As we cannot use CreateObject with classes within the project, we have to add our classes here manually...
'
Option Explicit

Private Type TxhvComponent
    Cache As Scripting.dictionary
End Type

Private this As TxhvComponent

Private Sub Class_Initialize()
    Set this.Cache = New Scripting.dictionary
End Sub

''
' Creates and initilizes a new xhvcomponent
' NOTE: THIS FUNCTION SHOULD BE CALLED ONLY FROM THE xhvRouter CLASS IN ITS navigate() METHOD
'
' @param ComponentName The name of the Component to create.
' @param attributes List of HTML Attributes of the component present in the template
' @param params List of parameters. Theses parameters come from the requested route in the navigate function.
' @param isUsedAsEntryPoint Indicate if the component is a direct child of a RouterPort
' @return An xhvComponent Interface component
Public Function CreateComponent(ComponentName As String, attributes As Scripting.dictionary, Params As Scripting.dictionary, IsUsedAsEntryPoint As Boolean) As xhvIController

On Error GoTo ERR_

    Select Case UCase(ComponentName)
        
        Case "APPCOMPONENT"
            Set CreateComponent = New AppController
        Case "TOPBARCOMPONENT"
            Set CreateComponent = New TopBarController
        Case "HOMECOMPONENT"
            Set CreateComponent = New HomeController
        Case "DEMOCLIENTLISTCOMPONENT"
            Set CreateComponent = New DemoClientListController
        Case "DEMOCLIENTDETAILCOMPONENT"
            Set CreateComponent = New DemoClientDetailController
        Case "BUTTONEXAMPLECOMPONENT"
            Set CreateComponent = New ButtonExampleController
        Case "BUTTONCOMPONENT"
            Set CreateComponent = New ButtonController
        Case "MANAGEFORMCOMPONENT"
            Set CreateComponent = New ManageFormController
            
        Case "XHVBOOTSTRAPSUBMITBUTTONCOMPONENT"
            Set CreateComponent = New BootstrapSubmitButtonController
            
        Case Else
            Throw 2013, , "Component '" & ComponentName & "' cannot be found." & vbCrLf & vbCrLf & "Add '" & ComponentName & "' in the Select Case of the createComponent() method of xhvComponentFactory class or correct the name of the component in the HTML template."
    End Select

    CreateComponent.IsUsedAsEntryPoint = IsUsedAsEntryPoint
    
    If CreateComponent.IsUsedAsEntryPoint Then
        'The attributes could not have not been set in the RenderComponent so fill them here before calling the onAfterInit sub.
            Dim viewAttributes As MSHTML.IHTMLAttributeCollection
            Set viewAttributes = CreateComponent.View.attributes
            Dim i As Integer
            'Reset  to New  for component UsedAsEntryPoint, because it's a static class
            Set attributes = New Scripting.dictionary
            For i = 0 To viewAttributes.length - 1
                'Add only those that are not empty
                If viewAttributes.Item(i).Value <> vbNullString Then
                    attributes.Add viewAttributes.Item(i).nodeName, viewAttributes.Item(i).Value
                End If
            Next i
    End If
    
    If Configuration("Environment.UseComponentsCacheManagement") Then
        'This loads the TemplateUrl file kept in memory in the TemplateString property so the Component does not have to load the file from the disk everytime
        preInitFromCache CreateComponent
    End If
    
    CreateComponent.Init Params
    CreateComponent.OnAfterInit attributes

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
' Initialize the HTML Template and Scripts, CSS , etc.. from memory Cache
'
' @param component The component to initialize from cache
'
Private Sub preInitFromCache(ByRef Component As xhvIController)

On Error GoTo ERR_


    Dim cachedComponent As xhvCachedComponent

    'Check if component is cached and cache it if needed
    If Not (this.Cache.Exists(Component.NameType)) Then
        Set cachedComponent = New xhvCachedComponent

        cachedComponent.ComponentName = Component.NameType

        If Trim(Component.TemplateString) = vbNullString Then
            cachedComponent.HtmlTemplate = readTemplateFromFile(Component.TemplateUrl)
        Else
            cachedComponent.HtmlTemplate = Component.TemplateString
        End If
        'TODO : In future version: Load CSS and Scripts in memory also

        this.Cache.Add cachedComponent.ComponentName, cachedComponent
    Else
        Set cachedComponent = this.Cache(Component.NameType)
    End If

    If Configuration("Environment.UseComponentCacheCompressing") Then
        'TODO : In  future version : Remove all unwanted space, line feed, etc.. from the string. Its just not a replace because some space need to be kept!
    End If

    Component.TemplateString = cachedComponent.HtmlTemplate


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
' Reads the html template from file
'
' @param templateFileName HTML Template file name (ex: mycomponent.component.html). This file name is set in the concerete component class.
' @return String: content of the file
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

