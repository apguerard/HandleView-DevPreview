VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvNavigationHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder lib.HandleView.Routing

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' This class represents the History of navigation made by the Router. It serves for the Back and Forward navigation implementation.
'
Option Explicit

Private Type TxhvNavigationHistory
    FirstNode As xhvNavigationHistoryNode
    LastNode As xhvNavigationHistoryNode
    ActualNode As xhvNavigationHistoryNode
    length As Integer
End Type

Private this As TxhvNavigationHistory

'ReadOnly
Public Property Get ActualNode() As xhvNavigationHistoryNode
    Set ActualNode = this.ActualNode
End Property

''
' Add a history node in the history
'
' @param node The history  node to add
'
Public Sub Push(node As xhvNavigationHistoryNode)
    
    Dim newFirstNode As xhvNavigationHistoryNode
    
    Set newFirstNode = node
    
    If this.length = 0 Then
        Set this.FirstNode = newFirstNode
        Set this.LastNode = newFirstNode
    Else
        Set this.FirstNode.PrevNode = newFirstNode
        Set newFirstNode.NextNode = this.FirstNode
        Set this.FirstNode = newFirstNode
    End If
    
    Set this.ActualNode = this.FirstNode
    this.length = this.length + 1
    
    Set newFirstNode = Nothing
    
End Sub

''
' Remove the first node in history
'
' @return The history node popped
Public Function Pop() As xhvNavigationHistoryNode

    If this.length = 0 Then
        Set Pop = Nothing
    ElseIf this.length = 1 Then
        Set Pop = this.FirstNode
        Set this.FirstNode = Nothing
        Set this.LastNode = Nothing
        Set this.ActualNode = Nothing
        this.length = this.length - 1
    Else
        Set Pop = this.FirstNode
        Set this.FirstNode = this.FirstNode.NextNode
        Set this.FirstNode.PrevNode = Nothing
'      Set this.ActualNode = this.FirstNode
        this.length = this.length - 1
    End If
    
End Function

''
' Move the ActualNode Pointer backward in the history stack
'
' @return The new ActualNode after moving backward in the history
'
Public Function MoveBack() As xhvNavigationHistoryNode

    If Not (this.ActualNode.PrevNode Is Nothing) Then
        Set this.ActualNode = this.ActualNode.PrevNode
    End If

    Set MoveBack = this.ActualNode

End Function

''
' Move the ActualNode Pointer forward in the history stack
'
' @return The new ActualNode after moving forward in the history
'
Public Function Forward() As xhvNavigationHistoryNode
    'TODO: Add in a future version
End Function

''
' Delete all the history point from the FirstNode to the ActualNode
'
Public Sub Refresh()

    If Not (this.ActualNode Is Nothing) Then
        If Not (this.ActualNode.PrevNode Is Nothing) Then
            Pop
            Refresh
        End If
    End If

End Sub
