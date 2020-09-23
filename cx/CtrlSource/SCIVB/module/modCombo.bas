Attribute VB_Name = "modCombo"
Option Explicit

' This module serves one purpose to support saving/loading and updating
' the combo boxes in the find, and find/replace dialogs.

Public Function AddCombo(Cmb As ComboBox, strNew As String)
  Dim i As Long, y As Long
  Dim lPos As Long
  Dim strTemp As String
  If strNew = "" Then Exit Function
  ' First let's verify that what were entering into the combo box
  ' isn't already there.
  For i = 0 To Cmb.ListCount - 1
    If LCase(Cmb.List(i)) = LCase(strNew) Then  'Case won't matter here.
      ' We already have this item in the combobox so let's go ahead
      ' and move it to the 0 spot and move everything else
      If i = 0 Then
        ' The item in question already is in the 0 spot
        ' so just exit the function all together
        RemExtra Cmb
        ComboSaveHistory Cmb
        Cmb.Text = strNew
        Exit Function
      End If
      ' Apparently it's not in the 0 spot so we can just move everything
      strTemp = Cmb.List(i)
      ' Store the string as were gonna destroy it
      Cmb.RemoveItem (i)
      
      ' Add the last item in before moving everything
      Cmb.AddItem Cmb.ListCount - 1
      'Move everything to the next spot up
      For y = Cmb.ListCount - 1 To 1 Step -1
        Cmb.List(y) = Cmb.List(y - 1)
      Next y
      Cmb.List(0) = strTemp ' Go ahead and set the 0 spot to the temp str
      RemExtra Cmb
      ComboSaveHistory Cmb
      Cmb.Text = strNew
      
      Exit Function  ' Get out of this function
    End If
  Next i
  Cmb.AddItem Cmb.ListCount - 1
  For i = Cmb.ListCount - 1 To 0 Step -1
    Cmb.List(i) = Cmb.List(i - 1)
  Next i
  RemExtra Cmb
  Cmb.List(0) = strNew
  ComboSaveHistory Cmb
  Cmb.Text = strNew
End Function

Public Sub ComboSaveHistory(ByRef comboObj As ComboBox)
  Dim nCount As Integer
  For nCount = 0 To comboObj.ListCount - 1
    Call SaveSetting(App.Title, "History", comboObj.Name & Format(nCount), comboObj.List(nCount))
  Next nCount
  ' Mark End
  On Local Error Resume Next
  DeleteSetting App.Title, "History", comboObj.Name & Format(nCount)
End Sub
Public Sub ComboLoadHistory(ByRef comboObj As ComboBox)
Dim Temp As String
Dim nCount As Integer
  comboObj.Clear
  Do
    On Error GoTo e_Trap
    Temp = GetSetting(App.Title, "History", comboObj.Name & Format(nCount), Default:=Chr$(255))
    If Not Temp = Chr$(255) Then
     ' Add item to ComboBox list
      comboObj.AddItem Temp
    Else
      Exit Do
    End If
      nCount = nCount + 1
  Loop
  Exit Sub
e_Trap:
    Exit Sub
End Sub

Public Sub RemExtra(Cmb As ComboBox)
  Dim i As Long
  If Cmb.ListCount > 9 Then
    For i = Cmb.ListCount - 1 To 10 Step -1
      Cmb.RemoveItem i
    Next i
  End If
End Sub
