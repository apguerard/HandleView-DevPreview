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
    Name As String
    Email As String
End Type
Private this As TDemoClientModel

Public Property Get Id() As Long
    Id = this.Id
End Property
Public Property Let Id(Value As Long)
    this.Id = Value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property
Public Property Let Name(Value As String)
    this.Name = Value
End Property

Public Property Get Email() As String
    Email = this.Email
End Property
Public Property Let Email(Value As String)
    this.Email = Value
End Property
