VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvConfigurator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' Provides methods to add configuration in the application.
' These Configurations will be available in the Global Configuration Dictionary
'
Option Explicit

Private Type TxhvConfigurator
    Configurations As Scripting.dictionary
End Type
Private this As TxhvConfigurator

Private Sub Class_Initialize()
    Set this.Configurations = New Scripting.dictionary
End Sub


'ReadOnly
Public Property Get Configurations() As Scripting.dictionary
    Set Configurations = this.Configurations
End Property

''
' Adds configuration in the Configuration Global Dictionary via a table in the actual Database
'
' @param tableName Name of the table that contains the Configurations
' @param idFieldLabel Name of the field that contains the configuration id
' @param ivalueStringLabel Name of the field that contains the configuration value
'
Public Sub AddLocalDB(tableName As String, idFieldLabel, ivalueStringLabel, Optional filter As String)

On Error GoTo ERR_

    Dim sSQL As String
    Dim rs As DAO.Recordset
    
    sSQL = "SELECT " & idFieldLabel & " as ConfigIdField, " & ivalueStringLabel & " as ConfigValueField FROM " & tableName
    
    If Trim(filter) <> vbNullString Then
        sSQL = sSQL & " " & filter
    End If
    
    Set rs = CurrentDb.OpenRecordset(sSQL)
    
        If Not (rs.EOF And rs.BOF) Then
            Dim sID As String
            Do While Not rs.EOF
                sID = IIf(IsNull(rs![ConfigIdField].Value), vbNullString, rs![ConfigIdField].Value)
                If this.Configurations.Exists(sID) Then
                    xhvLog.LogWarning "The configuration '{{config}}' appears in double but won't be overwritten.", sID
                ElseIf IsNull(rs![ConfigValueField].Value) Then
                    xhvLog.LogWarning "The configuration '{{config}}' contains the value: null.", sID
                Else
                    this.Configurations.Add sID, tryCastType(rs![ConfigValueField].Value)
                End If
                rs.MoveNext
            Loop
        End If
    
    rs.Close
    
    Set rs = Nothing
    
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


Private Function tryCastType(valueString As String) As Variant
    
On Error Resume Next
        tryCastType = CBool(valueString)
        If Err.Number = 0 Then Exit Function

        tryCastType = CLng(valueString)
        If Err.Number = 0 Then Exit Function

    
        tryCastType = valueString
    On Error GoTo 0

End Function


