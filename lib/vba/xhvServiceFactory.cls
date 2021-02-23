VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvServiceFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

''
' This class is responsible for creating and initializing services in the  DI framework
'
' <attribute>
' VB_PredeclaredId = True
' </attribute>
'
Option Explicit


''
' Return a service by  calling the function put in the useFactory property of the service when registered.
'
' IMPORTANT
'
' THIS FUNCTION MUST  BE DEFINED AS PUBLIC IN THIS MODULE
'
'
' @param factoryFunction Name of the function to call
' @return The object returned by the calling of the factory function.
'

Public Function GetServiceFromFactory(factoryFunction As String) As Object
    Set GetServiceFromFactory = CallByName(Me, factoryFunction, VbMethod)
End Function


''
' Return a service by creating the concrete class set in  the useClass property of the service when registered.
'
' IMPORTANT
'
' THE className CORRESPONDS TO THE USECLASS PROPERTY VALUE WHEN YOU REGISTER THE SERVICE
'
'
' @param className Name of the class of the service
' @return Object of a concrete class.
Public Function GetServiceFromclassName(className As String) As Object

    Select Case className
        Case "DemoClientService"
            Set GetServiceFromclassName = New DemoClientService
    End Select

End Function





