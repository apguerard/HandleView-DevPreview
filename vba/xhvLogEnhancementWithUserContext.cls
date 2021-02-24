VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvLogEnhancementWithUserContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Logging

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Gu�rard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
' *********************************************************************

''
' This class an enhancement for the Logging Service that return the username of the
' actual user logged on to the machine
'
Option Explicit

Private Const ENHANCEMENT_NAME As String = "UserName"

Implements xhvILoggerEnhancement


Private Property Get xhvILoggerEnhancement_enhancementName() As String
    xhvILoggerEnhancement_enhancementName = ENHANCEMENT_NAME
End Property

''
' Return a list of additional user context variables for Loggers
'
' @return Dictionnary with User Context Vars
'
Private Function xhvILoggerEnhancement_returnEnhancementItems() As Scripting.dictionary

On Error GoTo ERR_

    Dim dictTemp As Scripting.dictionary
    Set dictTemp = New Scripting.dictionary

    dictTemp.Add "UserName", VBA.Interaction.Environ$("UserName")
    dictTemp.Add "ComputerName", VBA.Interaction.Environ$("COMPUTERNAME")
    dictTemp.Add "OS", VBA.Interaction.Environ$("OS")

    Set xhvILoggerEnhancement_returnEnhancementItems = dictTemp
    
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
