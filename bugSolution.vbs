Function MyImprovedFunction(param1)
  If VarType(param1) = vbEmpty Or param1 = "" Then
    ' Handle empty or null parameter
  Else
    ' Process non-empty parameter
  End If
End Function

'More robust check using VarType to explicitly check for empty variants or empty strings.