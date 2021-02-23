Attribute VB_Name = "xhvGlobalObjects"
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
' Various shortcuts to various Global Objects and Functions in the Framework.
' These are just use to simplify code and calls to these Objects/Functions
'
Option Explicit

Global Document As MSHTML.HTMLDocument
Global HV As xhv
Global OpenedBackendDatabases() As DAO.Database

''
' Act as a Global Object and return a configuration value.
' This facade to xhvConfigurator.Configurations is used to check globaly if a configuration exists.
'
' @param configurationId Id (Name)of the configuration
' @return Value of configuration or raise an error if the configuration do not exist
Public Function Configuration(configurationId As String) As Variant

    If xhvConfigurator.Configurations.Exists(configurationId) Then
        Configuration = xhvConfigurator.Configurations(configurationId)
    Else
        Throw 2002, , "The configuration '" & configurationId & "' was not found. " & vbCrLf & vbCrLf & "Please check that you have initialized this configuration or correct the Id in the Configuration call."
    End If


End Function
