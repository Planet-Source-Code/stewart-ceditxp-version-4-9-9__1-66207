VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.1#0"; "SCIVBX.ocx"
Begin VB.Form frmDoc 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   4425
   ClientTop       =   2910
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   6960
   Begin SCIVBX.SCIVB sciMain 
      Left            =   1680
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isFile As Boolean
Public FileName
Public FTP As Boolean
Public Changed As Boolean
Public FTPAccount As String
Public FTPDir As String
Public lLangIndex As Integer
Private LineIndex As Long
Private LastLine As New Collection

Private Sub Form_Activate()
  On Error Resume Next
  SetCheck lLangIndex
  sciMain.SetFocus
  dnum = Tag
End Sub

Private Sub Form_GotFocus()
'  RedrawWin frmMain.hwnd
End Sub

Private Sub Form_Load()
  sciMain.InitScintilla Me.hwnd
  LastLine.Add 0
  LineIndex = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim mResult As VbMsgBoxResult
  If sciMain.Modified = True Then
    mResult = MsgBox("File: " & Caption & vbCrLf & "This file has been modified since it was last saved!" & vbCrLf & "Would you like to save this file now?", vbYesNoCancel, "Save")
    If mResult = vbYes Then
      frmMain.Save
    ElseIf mResult = vbCancel Then
      Cancel = True
      StopClose = True
    End If
  End If
  
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  sciMain.MoveSCI 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuCopy_Click()
  sciMain.Copy
End Sub

Private Sub mnuCut_Click()
  sciMain.Cut
End Sub

Private Sub mnuFileNew_Click()
  mdiFrmMain.mnuNew_Click
End Sub

Private Sub mnuFind_Click()
  sciMain.DoFind
End Sub

Private Sub mnuFindNext_Click()
  sciMain.FindNext
End Sub

Private Sub mnuFindPrev_Click()
  sciMain.FindPrev
End Sub

Private Sub mnuHighlighters_Click(Index As Integer)

End Sub

Private Sub mnuGoto_Click()
  sciMain.DoGoto
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set LastLine = Nothing
  FState(Me.Tag).Deleted = True
End Sub

Private Sub sciMain_KeyPress(Char As Long)
  'Debug.Print "Character Entered: " & Chr(char)
End Sub

Private Sub sciMain_MouseDown(Button As Integer, Shift As Integer, x As Long, Y As Long)
  'Debug.Print IIf(Button = vbLeftButton, "Left ", "Right ") & " Button | X = " & x & " | Y = " & Y
  If Button = vbRightButton Then
    PopupMenu frmMain.mnuEdit
  End If
End Sub

Private Sub sciMain_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
  'Debug.Print IIf(Button = vbLeftButton, "Left ", "Right ") & " Button | X = " & x & " | Y = " & Y
End Sub

Private Sub sciMain_UpdateUI()
  frmMain.stbMain.Panels(4).Text = "CurrentLine: " & sciMain.GetCurrentLine + 1 & " Column: " & sciMain.GetColumn & " Lines: " & sciMain.DirectSCI.GetLineCount
  If LastLine.Count <> 0 Then
    If LastLine(LineIndex) <> sciMain.GetCurrentLine Then
      LastLine.Add sciMain.GetCurrentLine
      LineIndex = LastLine.Count
    End If
  Else
    LastLine.Add sciMain.GetCurrentLine
    LineIndex = LastLine.Count
  End If
End Sub

Public Sub SetCheck(iIndex As Integer)
  On Error Resume Next
  Dim I As Long
  For I = 0 To iLngCount
    frmMain.mnuHighlighter(I).Checked = False
  Next I
  frmMain.mnuHighlighter(iIndex).Checked = True
  lLangIndex = iIndex
End Sub

Public Sub NextLine()
  On Error Resume Next
  If LineIndex + 1 > LastLine.Count Then Exit Sub 'this won't work we are at the end of lines
  LineIndex = LineIndex + 1
  'rt.ExecuteCmd cmCmdGotoLine, LastLine(LineIndex)
  sciMain.DirectSCI.GotoLine LastLine(LineIndex)
End Sub

Public Sub PrevLine()
  On Error Resume Next
  If LineIndex - 1 < 1 Then Exit Sub 'this won't work we are at the end of lines
  LineIndex = LineIndex - 1
  sciMain.GotoLine LastLine(LineIndex)
End Sub

