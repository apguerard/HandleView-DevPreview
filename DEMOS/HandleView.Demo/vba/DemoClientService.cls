VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClientService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Services

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

Option Explicit

Implements IDemoClientService

Private counter As Integer

Public Function GetAllClients() As Collection

    Dim qdef As QueryDef
    Dim rs As DAO.Recordset
    Dim clientModel As DemoClientModel
    
    Set GetAllClients = New Collection
    
    Set qdef = CurrentDb.QueryDefs("qry_DemoClientService_GetAllClients")
    
    Set rs = qdef.OpenRecordset(, dbReadOnly)
    
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            Set clientModel = New DemoClientModel
            
            clientModel.Id = rs!Id
            clientModel.Name = rs!ClientName
            clientModel.Email = rs!ClientEmail
            
            GetAllClients.Add clientModel
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Set qdef = Nothing
    Set clientModel = Nothing

End Function


Private Function IDemoClientService_GetAllClients() As Collection
    Set IDemoClientService_GetAllClients = Me.GetAllClients()
End Function

Private Function IDemoClientService_getClientById(Id As Long) As DemoClientModel

    Dim qdef As QueryDef
    Dim rs As DAO.Recordset
    
    Set qdef = CurrentDb.QueryDefs("qry_DemoClientService_getClientById")
    qdef.Parameters!PID = Id
    
    Set rs = qdef.OpenRecordset(, dbReadOnly)
    
    If Not (rs.EOF And rs.BOF) Then
        Set IDemoClientService_getClientById = New DemoClientModel
        IDemoClientService_getClientById.Id = rs!Id
        IDemoClientService_getClientById.Name = rs!ClientName
        IDemoClientService_getClientById.Email = rs!ClientEmail
    Else
        Set IDemoClientService_getClientById = Nothing
        Exit Function
    End If

    rs.Close
    Set rs = Nothing
    Set qdef = Nothing

End Function
