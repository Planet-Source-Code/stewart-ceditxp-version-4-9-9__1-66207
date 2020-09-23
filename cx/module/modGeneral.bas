Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub FlatBorder(ByVal hwnd As Long)
  Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Function InputStr(Optional Question As String, Optional WinTitle As String, Optional default As String, Optional Start As Integer, Optional IconFile As String) As String
  If Question <> "" Then
    frmInput.lblInfo.Caption = Question
  End If
  If WinTitle <> "" Then
    frmInput.Caption = WinTitle
  Else
    frmInput.Caption = App.Title
  End If
  If default <> "" Then
    frmInput.txtInput.Text = default
  End If
  If IconFile <> "" Then
    frmInput.picIcon.Picture = LoadPicture(IconFile)
  End If
  If Start <> 0 Then
    frmInput.txtInput.SelStart = Start
  End If
  frmInput.Show vbModal
  InputStr = Result
End Function


Public Sub GetAccounts(cbo As ComboBox)
  Dim s As String
  cbo.Clear
  s = Dir(App.path & "\accounts\")
  Do While s <> ""
    If Right(s, 3) = "ftp" Then
      cbo.AddItem Left(s, Len(s) - 4)
    End If
    s = Dir
  Loop
End Sub

Public Function StrWrap(str As String) As String
  StrWrap = """" & str & """"
End Function

Public Function SplitStr(str As String, ReturnStr As String) As String
  ' This will split a set of words in quotes
  ' Used in the ftp portion
  Dim FindQ As Long, FindQ2 As Long
  FindQ = InStr(1, str, """")
  If FindQ = 0 Then
    SplitStr = ""
    Exit Function
  End If
  FindQ2 = InStr(FindQ + 1, str, """")
  If FindQ2 = 0 Then
    SplitStr = ""
    Exit Function
  End If
  SplitStr = Mid(str, FindQ + 1, FindQ2 - 2)
  ReturnStr = Mid(str, FindQ2 + 1, Len(str) - FindQ2)
End Function


Public Sub Flatten(ByVal frm As Form)
  Dim CTL As Control
  For Each CTL In frm.Controls
    Select Case TypeName(CTL)
      Case "CommandButton", "TextBox", "ListBox", "FileTree", "TreeView", "ProgressBar", "PictureBox"
        FlatBorder CTL.hwnd
    End Select
  Next
End Sub

'+--------------------------------------------------------------------+
'| CheckPath is a simple function that will insert the needed \ on the|
'| end of a path if it's not there. Thats all :)                      |
'+--------------------------------------------------------------------+
Public Function CheckPath(ByVal path As String) As String
If Right$(path, 1) <> "\" Then
  CheckPath = path & "\"
Else
  CheckPath = path
End If
End Function

Public Sub InsertString(rt As SCIVB, str As String)
      rt.SelText = str
      rt.SetFocus
End Sub


Public Function GetExtension(sFileName As String) As String
    Dim lPos As Long
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then
        GetExtension = " "
    Else
        GetExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
    End If
End Function


Public Function StripPath(t As String) As String
Dim x As Integer
Dim ct As Integer
    StripPath = t
    x = InStr(t, "\")
    Do While x
        ct = x
        x = InStr(ct + 1, t, "\")
    Loop
    If ct > 0 Then StripPath = Mid(t, ct + 1)
End Function

Public Sub SaveCoolbar(clBar As CoolBar, iniFile As String)
  'On Error Resume Next
  Dim x As Long
  For x = 1 To clBar.Bands.Count
    writeini "COLBAR" & x, "Position", clBar.Bands(x).Position, iniFile
    writeini "COLBAR" & x, "Width", clBar.Bands(x).Width, iniFile
    writeini "COLBAR" & x, "Visible", clBar.Bands(x).Visible, iniFile
    writeini "COLBAR" & x, "Control", clBar.Bands(x).Child.Name, iniFile
  Next x
End Sub

Public Sub ReadCoolbar(clBar As CoolBar, iniFile As String)
  'On Error Resume Next
  Dim x As Long
  Dim lPos As Long
  Debug.Print clBar.Bands.Count
  For x = 1 To clBar.Bands.Count
    Debug.Print x
    'clBar.Bands(x).Position = ReadINI("COLBAR" & x, "Position", iniFile)
    lPos = ReadINI("COLBAR" & x, "Position", iniFile)
    clBar.Bands(x).Width = ReadINI("COLBAR" & x, "Width", iniFile)
    clBar.Bands(x).Visible = ReadINI("COLBAR" & x, "Visible", iniFile)
    Set clBar.Bands(lPos).Child = FindToolbar(ReadINI("COLBAR" & x, "Control", iniFile))
  Next x
End Sub

Public Function FindToolbar(strName As String) As ToolBar
  Dim ctrl As Control
  For Each ctrl In frmMain.Controls
    Select Case TypeName(ctrl)
      Case "Toolbar"
        If LCase(ctrl.Name) = LCase(strName) Then
          Debug.Print strName
          Set FindToolbar = ctrl
          Exit Function
        End If
    End Select
  Next
End Function
