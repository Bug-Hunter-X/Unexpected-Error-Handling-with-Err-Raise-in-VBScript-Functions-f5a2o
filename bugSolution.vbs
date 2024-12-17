Function MyFunction(param1, param2)
  On Error GoTo MyErrorHandler
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Parameters cannot be empty!" 
  End If
  ' ... rest of the function
  Exit Function
MyErrorHandler:
  MsgBox "Error in MyFunction: " & Err.Description, vbCritical
  Err.Clear
End Function

'Example of calling the function and handling potential errors
On Error Resume Next
Dim result: result = MyFunction(1, "")
If Err.Number <> 0 Then
  MsgBox "Error calling MyFunction: " & Err.Description, vbCritical
  Err.Clear
End If
On Error GoTo 0 'Reset error handling