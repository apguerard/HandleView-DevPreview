Attribute VB_Name = "xhvExceptionManagement"
'@Folder lib.HandleView.Logging

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Contains functions to help manage error in HandleView and Application
'
Option Explicit

Public Function ReThrow()

    Err.Raise Err.Number, Err.source, Err.Description & " - Rethrown"

End Function

Public Function Throw(errNumber As Long, Optional errSource As String, Optional errDescription As String)
    
    Err.Raise errNumber, errSource, errDescription

End Function

