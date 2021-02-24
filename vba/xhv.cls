VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhv"
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
' This class represent the HandleView framework
'
Option Explicit

Private Type TxhvStore
    version As String
End Type
Private this As TxhvStore

Private Sub Class_Initialize()
    
    this.version = xhvConst.FRAMEWORK_VERSION
    
End Sub


'TODO: Add context system
'''
'' Create a new Context and Add it to the store
''
'' @param ContextType The type of the Context to Create
'' @return The new COntext created
'Public Function CreateContext(ContextName As String, ContextState As Variant) As xhvContext
'
'  Set CreateContext = New xhvContext
'
'  CreateContext.ContextName = ContextName
'  Set CreateContext.state = ContextState
'
'  this.store.Contexts.Add ContextName, CreateContext
'
'End Function
'
'''
'' Remove a context  from the store
''
'' @param ContextType The type of the Context to Remove
'' @return N/A  But the context is destroyed and removed from the store
'Public Sub RealeaseContext(ContextName As String)
'
'  Set this.store.Contexts.Item(ContextName) = Nothing
'  this.store.Contexts.Remove ContextName
'
'End Sub
'
'''
'' Return a context stored in the store
''
'' @param ContextName Name of the context to return
'' @return The context named ContextName
'Public Function UseContext(ContextName) As xhvContext
'
'  MsgBox "TO DO: Check if context exits.", vbCritical
'
'  Set UseContext = this.store.Contexts.Item(ContextName)
'
'End Function


