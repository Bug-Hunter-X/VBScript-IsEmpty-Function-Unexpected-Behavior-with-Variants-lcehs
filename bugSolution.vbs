Function MyFunction(param1, param2)
  If TypeName(param1) = "Empty" Then
    param1 = ""
  End If
  If TypeName(param2) = "Empty" Then
    param2 = 0
  ElseIf Not IsNumeric(param2) Then
    'Handle non-numeric input for param2 appropriately, e.g., raise an error
    Err.Raise vbError, , "param2 must be numeric"
  End If
  ' ... rest of the function ...
End Function