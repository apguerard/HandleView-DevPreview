Attribute VB_Name = "xhvEnum"
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
' This module contains the public Enums for HandleView
'

Public Enum xhvENUM_DependencyInjectionScope
    SINGLETON = 1
    TRANSIENT = 2
    SCOPED = 3
End Enum

Public Enum xhvENUM_LogLevel
    TRACE_LEVEL = 0
    DEBUG_LEVEL = 1
    INFO_LEVEL = 2
    WARNING_LEVEL = 3
    ERROR_LEVEL = 4
    CRITICAL_LEVEL = 5
End Enum
