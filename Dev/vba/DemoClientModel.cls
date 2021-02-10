VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClientModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Models

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class is the Application component model
'

Option Compare Database

Private Type TDemoClientModel
    Id As Long
    name As String
    email As String
End Type
Private this As TDemoClientModel

Public Property Get Id() As Long
    Id = this.Id
End Property
Public Property Let Id(value As Long)
    this.Id = value
End Property

Public Property Get name() As String
    name = this.name
End Property
Public Property Let name(value As String)
    this.name = value
End Property

Public Property Get email() As String
    email = this.email
End Property
Public Property Let email(value As String)
    this.email = value
End Property
