Attribute VB_Name = "modArray"
Public Function GetUpper(varArray As Variant) As Long
Dim Upper As Integer
On Error Resume Next
Upper = UBound(varArray)
If Err.Number Then
     If Err.Number = 9 Then
          Upper = 0
     Else
          With Err
               MsgBox "Error:" & .Number & "-" & .Description
          End With
          Exit Function
     End If
Else
     Upper = UBound(varArray) + 1
End If
On Error GoTo 0
GetUpper = Upper
End Function
