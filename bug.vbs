Function GetValue(key)
  On Error Resume Next
  SetValue key, "some value"
  result = GetSetting("MySection", key)
  If Err.Number <> 0 Then
    Err.Clear
    result = ""
  End If
  On Error GoTo 0
  GetValue = result
End Function

'Example usage
MsgBox GetValue("MyKey")