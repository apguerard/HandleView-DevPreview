VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvNavigationHistoryNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Routing

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Gu�rard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents a history node (which is a xhvRoute).
'
Option Explicit

Private Type TxhvNavigationHistoryNode
    NextNode As xhvNavigationHistoryNode
    NodeValue As xhvRoute
    NodeValueTarget As String
    PrevNode As xhvNavigationHistoryNode
End Type

Private this As TxhvNavigationHistoryNode

Public Property Get NextNode() As xhvNavigationHistoryNode
    Set NextNode = this.NextNode
End Property
Public Property Set NextNode(Value As xhvNavigationHistoryNode)
    Set this.NextNode = Value
End Property

Public Property Get PrevNode() As xhvNavigationHistoryNode
    Set PrevNode = this.PrevNode
End Property
Public Property Set PrevNode(Value As xhvNavigationHistoryNode)
    Set this.PrevNode = Value
End Property

Public Property Get NodeValue() As xhvRoute
    Set NodeValue = this.NodeValue
End Property
Public Property Set NodeValue(Value As xhvRoute)
    Set this.NodeValue = Value
End Property

Public Property Get NodeValueTarget() As String
    NodeValueTarget = this.NodeValueTarget
End Property
Public Property Let NodeValueTarget(Value As String)
    this.NodeValueTarget = Value
End Property


