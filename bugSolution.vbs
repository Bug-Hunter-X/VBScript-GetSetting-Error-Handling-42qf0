Function GetValue(key)
  Dim result
  On Error GoTo ErrorHandler
  result = GetSetting("MySection", key)
  On Error GoTo 0
  GetValue = result
  Exit Function

ErrorHandler:
  If Err.Number = 5 Then 'Setting not found
    result = ""
  Else
    MsgBox "Error: " & Err.Number & " - " & Err.Description
    ' Add more sophisticated error logging if needed
  End If
  Err.Clear
  GetValue = result
End Function

' Example Usage
MsgBox GetValue("MyKey")
MsgBox GetValue("NonExistentKey") 