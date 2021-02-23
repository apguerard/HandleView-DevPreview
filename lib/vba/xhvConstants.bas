Attribute VB_Name = "xhvConstants"
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
' This module contains the Global Const for HandleView

Option Explicit
Private Const MODULE_NAME As String = "xhvStartupApp"

Public Type TxhvConst
    FRAMEWORK_VERSION As String
    DEBUG_MODE As Boolean
    FAIL_SILENT_LOG_EXCEPTION As Boolean
    
    APP_CONFIG_TABLE_NAME As String
    APP_CONFIG_ID_FIELD As String
    APP_CONFIG_VALUE_FIELD As String
    APP_FORM_NAME As String
End Type

Global xhvConst As TxhvConst

Public Sub SetConst()

    xhvConst.FRAMEWORK_VERSION = "0.0.2"
    xhvConst.DEBUG_MODE = False
    xhvConst.FAIL_SILENT_LOG_EXCEPTION = False
    
    xhvConst.APP_CONFIG_TABLE_NAME = "xhvAppConfig"
    xhvConst.APP_CONFIG_ID_FIELD = "ConfigId"
    xhvConst.APP_CONFIG_VALUE_FIELD = "ConfigValue"
    xhvConst.APP_FORM_NAME = "App"
    
End Sub
