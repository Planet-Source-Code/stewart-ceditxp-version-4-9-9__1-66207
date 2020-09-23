VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.0#0"; "SCIVBX.ocx"
Begin VB.Form frmDoc 
   BackColor       =   &H8000000A&
   Caption         =   "SCIVBNote"
   ClientHeight    =   5160
   ClientLeft      =   3645
   ClientTop       =   2925
   ClientWidth     =   7770
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7770
   Begin SCIVBX.SCIVB sciMain 
      Left            =   2040
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      EdgeColumn      =   80
      EdgeMode        =   1
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strFile As String
Private iLngCount As Integer

Private Sub Form_Load()
  Dim str As String
  'On Error Resume Next
  ' Load the Highlighters
  sciMain.InitScintilla Me.hWnd
  Form_Resize
  frmMain.mnuBookmarks.Checked = sciMain.ShowFlags
  frmMain.mnuLineNumbers.Checked = sciMain.LineNumbers
  If sciMain.WordWrap = noWrap Then
    frmMain.mnuWordWrap.Checked = False
  Else
    frmMain.mnuWordWrap.Checked = True
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim msgRes As VbMsgBoxResult
  If sciMain.Modified = True Then
    msgRes = MsgBox("This document has been modified.  Do you wish to save?", vbYesNoCancel + vbQuestion, "Modified")
    If msgRes = vbCancel Then
      Cancel = True
      sciMain.SetFocus
      Exit Sub
    ElseIf msgRes = vbNo Then
      Cancel = False
    Else
      ' Save it
    End If
  End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  sciMain.MoveSCI 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub sciMain_FindFailed(FindText As String)
  Beep
End Sub

Private Sub sciMain_OnError(Number As String, Description As String)
'  MsgBox "Error Number: " & Number & vbCrLf & vbCrLf & Description, vbOKOnly, "Error: " & Number
End Sub

  
Private Sub sciMain_UpdateUI()
  frmMain.stbMain.PanelText(1) = "Line: " & sciMain.GetCurrentLine + 1 & " Col: " & sciMain.GetColumn
End Sub
