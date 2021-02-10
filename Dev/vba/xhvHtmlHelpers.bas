Attribute VB_Name = "xhvHtmlHelpers"
'@Folder lib.HandleView.Helpers

' Copyright (C) 2019 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Framework HTML Helper functions.
' Can be useful in the rendering of the component
'
Option Explicit
Private Const MODULE_NAME As String = "xhvHtmlHelpers"

''
' Execute javascript in the main document
'
' @param ScriptId Name of the function. Leave Empty ("") if the script does not define a function you need to remove later.
'                 However, it won't remove script from browser memory.
' @param Script Contains the javascript you want ot run. Can be any valid javascript.
' @return True if everything went ok.
' @remarks This function will remove a javascript tag with a given ID.  Then it will load the Script into the DOM causing it to execute.
Public Function ExecuteJS(ScriptId As String, script As String) As Boolean
On Error GoTo ERR_

    Dim s As MSHTML.HTMLScriptElement
    Dim a As MSHTML.HTMLScriptElement
        
    'Remove the script if it is already loaded. NOTE: It doesn't remove code already in browser memory.
    If Document.querySelectorAll("[id='" & ScriptId & "']").length <> 0 Then
        Set a = Document.querySelector("[id='" & ScriptId & "']")
        a.ParentNode.removeChild a
    End If
    'Add the script (run automatically when loaded)
    Set s = Document.createElement("script")
    s.Id = ScriptId
    s.innerText = script
    s.Type = "text/javascript"
    Document.getElementsByTagName("head")(0).appendChild s
        
    'Cleaning
    Set s = Nothing
    Set a = Nothing
    
    ExecuteJS = True

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
' Populate combo box
'
' @param dropDown ID of HTML Dropdown from DOM
' @param service Data service to call to get list
' @param sqlString String sql command to query data service
' @return N/A
'
'Public Sub InitializeDropdown(dropDown As MSHTML.HTMLSelectElement, _
'                            service As Object, _
'                            sqlString As String)
'
'  Dim listItem As dropDownListItemModel
'  Dim itemsCollection As Collection
'  Dim options As MSHTML.HTMLOptionElement
'
'  Set itemsCollection = CallByName(service, SQLMethod, VbMethod)
'
'  ExecuteJS vbNullString, "$('#" & dropDown.Id & "').empty();"
'
'  For Each listItem In itemsCollection
'      Set options = Document.createElement("option")
'      options.Text = listItem.DisplayValue
'      options.value = listItem.ItemID
'      dropDown.Add options
'  Next
'
'End Sub

