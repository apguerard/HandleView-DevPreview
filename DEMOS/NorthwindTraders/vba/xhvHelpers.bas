Attribute VB_Name = "xhvHelpers"
'@Folder lib.HandleView.Helpers

' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

''
' Bunch of little functions for trivial tasks.
'

Option Explicit
Private Const MODULE_NAME As String = "xhvHelpers"

''
' Remove space and empty lines from input
'
' @param temp String to manipulate
' @return temp less spaces and empty lines
'
Public Function RemoveSpaceAndLine(temp As String) As String
    RemoveSpaceAndLine = Replace(Replace(Replace(Replace(temp, " ", vbNullString), vbCrLf, vbNullString), vbCr, vbNullString), vbLf, vbNullString)
End Function

''
' Rplace single quote with double quote in input string
'
' @param inString input string to modify
' @return inString with single quotes replaced with doublequotes
'
Public Function DoubleQuote(inString As String) As String
    DoubleQuote = Replace(inString, "'", "''")
End Function

''
' Given a fileNameand path, return just the filename
'
' @param inPath String representing fully qualified file and path
' @return file name as String
Public Function FileName(ByVal inPath As String) As String
    FileName = Mid(inPath, InStrRev(inPath, "\") + 1)
End Function

''
' Gets a timestamp in the format of
' YYYY-MM-DDTHH:nn:ss.2 digit miliseconds
'
' @return Timestamp as String
'
Public Function GetTimeStamp() As String
    GetTimeStamp = Format(Now, "YYYY-MM-DDTHH:nn:ss") & "." & Right(Strings.Format(Timer, "#0.00"), 2)
End Function

''
' Helper function to get logging level label
'
' @param level as xhvENUM_LogLevel
' @return label as String
'
Public Function GetLoggingLevelLabel(level As xhvENUM_LogLevel) As String

    Select Case level
        Case TRACE_LEVEL
            GetLoggingLevelLabel = "TRACE   "
        Case DEBUG_LEVEL
            GetLoggingLevelLabel = "DEBUG   "
        Case INFO_LEVEL
            GetLoggingLevelLabel = "INFO    "
        Case WARNING_LEVEL
            GetLoggingLevelLabel = "WARNING "
        Case ERROR_LEVEL
            GetLoggingLevelLabel = "ERROR   "
        Case CRITICAL_LEVEL
            GetLoggingLevelLabel = "CRITICAL"
    End Select
    

End Function

''
' Helper function to see if file exists on disk
'
' @param fileName Path/To/File
' @return True/False
'
Public Function FileExists(FileName As String) As Boolean
    
    If Len(Dir(FileName)) = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
    
End Function

''
' Cut the input string at the required length and add a trailing mark at the end if needed
'
' @param inString The string we want the Excerpt from
' @param length The total length of the string after it has been cut. Total means that it will cut the string at the length minus the length of the trailing mark.
' @param tailingMark Optional. The trailing mark to add at the end. If no trailing mark is passed, it will use "[...]"
' @return@return String. The Excerpt from the string and trailing mark
' @remarks Not currently used anywhere
'
Public Function Excerpt(inString As String, length As Integer, Optional trailingMark As String = "[...]") As String

Dim temp As String
Dim pos As Integer

If Len(inString) <= length Then
    Excerpt = inString
Else
    temp = Trim(Left(inString, (length) - Len(trailingMark) + 1))
    pos = InStrRev(temp, " ")
    Excerpt = Left(temp, pos - 1) & " " & trailingMark
End If

End Function

''
' Get list from database helper function
'
' @param table The table name to get the list from
' @param column The comman separated columns (field) to get the data from
' @param sortColumn The column used to sort the data
' @param delimField Charcater to delimit fields
' @param delimRow Character to delimit rows
' @param filter Expression that returns true or false as predicate
' @return String Data returned from database
' @remarks Not currently used anywhere
Public Function GetList(table As String, column As String, sortColumn As String, delimField As String, delimRow As String, Optional filter As String = "True") As String

    Const NOCURRENTRECORD = 3021
    Dim rst As ADODB.Recordset
    Dim SQL As String
    Dim list As String
    
    SQL = "SELECT " & column & " FROM " & table & " WHERE " & filter & " ORDER BY " & sortColumn
   
    Set rst = New ADODB.Recordset
    
    With rst
        Set .ActiveConnection = CurrentProject.Connection
        .Open _
            source:=SQL, _
            CursorType:=adOpenForwardOnly, _
            options:=adCmdText
        
        On Error Resume Next
        list = .GetString(adClipString, , delimField, delimRow)
        .Close
        Select Case Err.Number
            Case 0
            ' no error so remove trailing delimiter
            ' and return string
            GetList = Left(list, Len(list) - Len(delimRow))
            Case NOCURRENTRECORD
            ' no rows in table so return
            ' zero length string
            Case Else
            ' unknown error
            GetList = "Error"
        End Select
        On Error GoTo 0
    End With
        
End Function

''
' Open a handle to all databases and keep it open during the entire time the application runs.
' Source  : Total Visual SourceBook
'
' @param init TRUE to initialize (call when application starts), FALSE to close (call when application ends)
' @return N/A
'
Sub OpenAllDatabases(Init As Boolean)
      
  Dim x As Integer
  Dim dbName As String
  Dim message As String

  ' Maximum number of back end databases to link
  Const MAX_DATABASES As Integer = 1

  ' List of databases kept in a static array so we can close them later
  Static openDatabases() As DAO.Database

  If Init Then
    ReDim openDatabases(1 To MAX_DATABASES)
    For x = 1 To MAX_DATABASES
      ' Specify your back end databases
      Select Case x
        Case 1:
          dbName = Configuration("App.BackEndPath")
      End Select
      message = vbNullString

      On Error Resume Next
      Set openDatabases(x) = openDatabases(dbName, False, False) 'Use next line if your bd as a password. It should :)  This one is only for demo purpose.
      'Set openDatabases(X) = openDatabases(dbName, False, False, ";pwd=password")
      If Err.Number > 0 Then
        message = "Trouble opening database: " & dbName & vbCrLf & _
                 "Make sure the drive is available." & vbCrLf & _
                 "Error: " & Err.Description & " (" & Err.Number & ")"
      End If
      On Error GoTo -1
      
      If message <> vbNullString Then
        MsgBox message
        Exit For
      End If
    Next x
  Else
    On Error Resume Next
    For x = 1 To MAX_DATABASES
      openDatabases(x).Close
    Next x
    On Error GoTo -1
  End If
  

  
End Sub


'TODO: Make a better version of this function....
'''
'' Open a handle to all databases and keep it open during the entire time the application runs.
'' Initial Source  : Total Visual SourceBook
''
'' @return N/A
'Sub hvxOpenAllBackendDatabases()
'
'
'
' Dim x As Integer
' Dim dbName As String
' Dim message As String
'
' ' Maximum number of back end databases to link
' Const MAX_DATABASES As Integer = 1
'
''  OpenedBackendDatabases
'
'
'  ReDim openDatabases(1 To MAX_DATABASES)
'  For x = 1 To MAX_DATABASES
'    ' Specify your back end databases
'    Select Case x
'      Case 1:
'        dbName = Configuration("App.BackEndPath")
'    End Select
'    message = vbNullString
'
'    On Error Resume Next
'    Set openDatabases(x) = openDatabases(dbName, False, False) 'Use next line if your bd as a password. It should :)  This one is only for demo purpose.
'    'Set openDatabases(X) = openDatabases(dbName, False, False, ";pwd=password")
'    If Err.Number > 0 Then
'      message = "Trouble opening database: " & dbName & vbCrLf & _
'               "Make sure the drive is available." & vbCrLf & _
'               "Error: " & Err.Description & " (" & Err.Number & ")"
'    End If
'    On Error GoTo -1
'
'    If message <> vbNullString Then
'      MsgBox message
'      Exit For
'    End If
'  Next x
'
'
'End Sub





