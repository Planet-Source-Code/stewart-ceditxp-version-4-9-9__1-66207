VERSION 5.00
Object = "*\A..\ScintillaWrapper.vbp"
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
   Begin ScintillaWrapper.usrScintilla sciMain 
      Left            =   1800
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      ShowCallTips    =   -1  'True
      AutoIndent      =   -1  'True
      IgnoreAutoCCase =   -1  'True
      AutoShowAutoComplete=   -1  'True
      AutoCompleteString=   $"frmMain.frx":0000
      LineEOL         =   0
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iLngCount As Integer
Public strFileName

Private Sub Form_Activate()
  sciMain.setfocusex
End Sub

Private Sub Form_Load()
  iLngCount = 0
  sciMain.InitScintilla
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim mResult As VbMsgBoxResult
  If sciMain.Modified = True Then
    mResult = MsgBox("File: " & Caption & vbCrLf & "This file has been modified since it was last saved!" & vbCrLf & "Would you like to save this file now?", vbYesNoCancel, "Save")
    If mResult = vbYes Then
      frmMain.Save
    ElseIf mResult = vbCancel Then
      Cancel = True
    End If
  End If
  
End Sub

Private Sub Form_Resize()
  sciMain.MoveIt 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuCopy_Click()
  sciMain.Copy
End Sub

Private Sub mnuCut_Click()
  sciMain.Cut
End Sub

Private Sub mnuExport_Click()
  sciMain.exporttohtml "c:\test1234567.html"
End Sub

Private Sub mnuFileNew_Click()
  mdiFrmMain.mnuNew_Click
End Sub

Private Sub mnuFind_Click()
  sciMain.doFind
End Sub

Private Sub mnuFindNext_Click()
  sciMain.FindNext
End Sub

Private Sub mnuFindPrev_Click()
  sciMain.findprev
End Sub

Private Sub mnuHighlighters_Click(Index As Integer)

End Sub

Private Sub mnuGoto_Click()
  sciMain.dogoto
End Sub


Private Sub sciMain_FindFailed(strFind As Variant)
  MsgBox "Filed to find [" & strFind & "]"
End Sub

Private Sub sciMain_KeyPress(char As Long)
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
  frmMain.stbMain.Panels(4).Text = "CurrentLine: " & sciMain.sciMain.GetCurLine + 1 & " Column: " & sciMain.sciMain.GetColumn & " Lines: " & sciMain.sciMain.GetLineCount
End Sub
