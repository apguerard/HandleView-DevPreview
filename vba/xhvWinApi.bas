Attribute VB_Name = "xhvWinApi"
'@Folder lib.HandleView.Config

' Copyright (C) 2021 Bluejacket Software - All Rights Reserved
' Copyright (C) 2019 Alain Guérard - All Rights Reserved
' You may use, distribute and modify this code under the
' terms of the MIT license.
'
' You should have received a copy of the MIT license with
' this file. If not, please visit : https://opensource.org/licenses/MIT
'

Option Explicit
#If VBA7 And Win64 Then

  Private Declare PtrSafe Function rtcCallByName Lib "VBE7.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As LongPtr, _
    ByVal CallType As VbCallType, _
    ByRef args() As Any, _
    Optional ByVal lcid As Long) As Variant
    
    Declare PtrSafe Sub CoCreateGuid Lib "ole32" (ByVal Guid As LongPtr)
    
#Else

  Private Declare Function rtcCallByName Lib "VBE6.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef args() As Any, _
    Optional ByVal lcid As Long) As Variant
    
    Declare Sub CoCreateGuid Lib "ole32" (ByVal guid As Long)
    
#End If


''
' As the basic CallByName function in VBA, except we can pass a regular Array as args instead of mandatory ParamArray in original CallByName VBA function
' NOTE: Only works for VbMethod calls
'
' @param Object The object upon which the function will be executed.
' @param ProcName A string expression containing the name of a property or method of the object we wish to call
' @param args A regular array containing the parameters we want to pass to the method
'
Public Function CallByNameXHV(Object As Object, ProcName As String, args() As Variant)
   assignResult CallByNameXHV, rtcCallByName(Object, StrPtr(ProcName), VbMethod, args)
End Function

''
' Returns a new GUID
'
' @return A string expression representing a GUID WITHOUT hypens
Public Function NewGUID() As String
    Dim b(15) As Byte
    CoCreateGuid VarPtr(b(0))
    NewGUID = Replace(Mid(Application.StringFromGUID(b), 8, 36), "-", vbNullString)
End Function

Private Sub assignResult(target, result)
  If VBA.IsObject(result) Then Set target = result Else target = result
End Sub



