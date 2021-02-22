VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xhvDI"
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
' HandleView Dependency Injection System
'
Option Explicit

Private Type TxhvDI
    services As Scripting.dictionary
    singletons As Scripting.dictionary
End Type
Private this As TxhvDI

Private Sub Class_Initialize()
    Set this.services = New Scripting.dictionary
    Set this.singletons = New Scripting.dictionary
End Sub


''
' Adds a service as a singleton in the Di container.
' NOTE: We must provide at least one of the 2 optional parameters. If the 2 are provided, the useClass will have precedency.
'     If both are empty, it will throw an exception.
'
' @param interface Id/Name of the interface under which the service will be added to the container. This id/name will be used when injecting services
' @param useClass Optional - Name of the concrete class to use when creating the requested interface
' @param useFactory Optional - Name of the factory function to call when creating the requested interface.
'
Public Sub AddSingleton(Interface As String, Optional UseClass As String, Optional UseFactory As String)

On Error GoTo ERR_

    Dim service As xhvService
    Set service = New xhvService
    
    If Trim(UseClass) = vbNullString And Trim(UseFactory) = vbNullString Then
        Throw 2014, , "You try to add a service in the DI container, but both useClass and useFactory properties are empty." & vbCrLf & vbCrLf & "You need to provide one of the two or the DI system won't be able to create the service."
    End If
    
    If this.services.Exists(Interface) Then
        Throw 2015, , "The service with the interface '" & Interface & "' already exits in the DI container." & vbCrLf & vbCrLf & "You cannot add two different services with the same interface id. Check your Service Configuration Module."
    End If
    
    service.Interface = Interface
    service.UseClass = UseClass
    service.UseFactory = UseFactory
    service.Scope = SINGLETON

    this.services.Add Interface, service
    
    Set service = Nothing
    
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


''
' Adds a service as transient in the Di container.
' This means that a new service will be created each time it is injected.
' NOTE:   We must provide at least one of the 2 optional parameters. If the 2 are provided, the useClass will have precedency.
'       If both are empty, it will throw an exception.
'
' @param interface Id/Name of the interface under which the service will be added to the container. This id/name will be used when injecting services
' @param useClass Optional - Name of the concrete class to use when creating the requested interface
' @param useFactory Optional - Name of the factory function to call when creating the requested interface.
'
Public Sub AddTransient(Interface As String, Optional UseClass As String, Optional UseFactory As String)

On Error GoTo ERR_

    Dim service As xhvService
    Set service = New xhvService
    
    If Trim(UseClass) = vbNullString And Trim(UseFactory) = vbNullString Then
        Throw 2014, , "You try to add a service in the DI container, but both useClass and useFactory properties are empty." & vbCrLf & vbCrLf & "You need to provide one of the two or the DI system won't be able to create the service."
    End If
    
    If this.services.Exists(Interface) Then
        Throw 2015, , "The service with the interface '" & Interface & "' already exits in the DI container." & vbCrLf & vbCrLf & "You cannot add two different services with the same interface id. Check your Service Configuration Module."
    End If
    
    service.Interface = Interface
    service.UseClass = UseClass
    service.UseFactory = UseFactory
    service.Scope = TRANSIENT

    this.services.Add Interface, service
    
    Set service = Nothing
    
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


''
' Return the concrete class requested by the interface name
' If the service was added as singleton, it return any previously created object.
' If the service eas added as transient, it always return a new object
'
' @param interface Id/Name of the interface for which we want to create the object.
' @return Concrete implementation correponding to the interface id/name
'
Public Function Inject(interfaceName As String) As Object

On Error GoTo ERR_

    Dim service As xhvService
    
    If this.services.Exists(interfaceName) Then
        Set service = this.services.Item(interfaceName)
        
        If Trim(service.UseClass) <> vbNullString Then
        
            If service.Scope = SINGLETON Then
                If Not (this.singletons.Exists(interfaceName)) Then
                    'Create the original service in not already in the singletons Dictionary
                    this.singletons.Add interfaceName, xhvServiceFactory.GetServiceFromclassName(service.UseClass)
                End If
                Set Inject = this.singletons.Item(interfaceName)
            ElseIf service.Scope = TRANSIENT Then
                Set Inject = xhvServiceFactory.GetServiceFromclassName(service.UseClass)
            End If
            
        ElseIf Trim(service.UseFactory) <> vbNullString Then
        
            If service.Scope = SINGLETON Then
                If Not (this.singletons.Exists(interfaceName)) Then
                    'Create the original service in not already in the singletons Dictionary
                    this.singletons.Add interfaceName, xhvServiceFactory.GetServiceFromFactory(service.UseFactory)
                End If
                Set Inject = this.singletons.Item(interfaceName)
            ElseIf service.Scope = TRANSIENT Then
                Set Inject = xhvServiceFactory.GetServiceFromFactory(service.UseFactory)
            End If
  
        Else
            'Check just in case...
            Throw 2014, , "You try to inject a service in the DI container, but both useClass and useFactory properties are empty." & vbCrLf & vbCrLf & "You need to provide one of the two or the DI system won't be able to create the service."
        End If
    Else
        Throw 2016, , "The service:'" & interfaceName & "' is not configured in the dependency injection system."
    End If
        
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

