VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#11.0#0"; "SCIVBX.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "SCIVBNote"
   ClientHeight    =   5160
   ClientLeft      =   3645
   ClientTop       =   3225
   ClientWidth     =   7770
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7770
   Begin SCIVBX.SCIHighlighter hlMain 
      Left            =   4560
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SCIVBX.SCIVB sciMain 
      Left            =   2520
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin Sample.ucStatusbar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Top             =   4905
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   450
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDateTime 
         Caption         =   "&Date/Time"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindPrev 
         Caption         =   "Find &Previous"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggle 
         Caption         =   "&Toggle Bookmark"
      End
      Begin VB.Menu mnuNextBookmark 
         Caption         =   "&Next Bookmark"
      End
      Begin VB.Menu mnuPrevBookmark 
         Caption         =   "&Previous Bookmark"
      End
      Begin VB.Menu mnuClearBookmark 
         Caption         =   "&Clear all Bookmarks"
      End
   End
   Begin VB.Menu mnuHigh 
      Caption         =   "Highlighters"
      Begin VB.Menu mnuHighlighter 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSyntax 
         Caption         =   "&Syntax Settings"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuWordWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnuLineNumbers 
         Caption         =   "&Line Numbers"
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "&Bookmarks"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iLngCount As Integer

Private Sub Form_Load()
  Dim str As String
  'On Error Resume Next
  ' Load the Highlighters
  sciMain.InitScintilla Me.hWnd
  SetupStatus
  Form_Resize
  InitCmnDlg Me.hWnd
  cmndlg.flags = 5
  hlMain.LoadHighlighters App.Path & "\highlighters"
  hlMain.ReadSettings "Software\SCIVBNote\SyntaxOptions\"
  hlMain.SetStylesAndOptions sciMain, "CPP"
  mnuBookmarks.Checked = sciMain.ShowFlags
  mnuLineNumbers.Checked = sciMain.LineNumbers
  If sciMain.WordWrap = noWrap Then
    mnuWordWrap.Checked = False
  Else
    mnuWordWrap.Checked = True
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
  sciMain.MoveSCI 0, 0, Me.ScaleWidth, Me.ScaleHeight - IIf(stbMain.Visible, stbMain.Height, 0)
End Sub

Private Sub hlMain_AddHighlighter(HighlighterName As String, Filter As String)
  AddMenu HighlighterName, HighlighterName, iLngCount
  iLngCount = iLngCount + 1
End Sub

Private Sub hlMain_ClearHighlighters()
  Dim i As Integer
  ' Lets clear our menu
  For i = mnuHighlighter.Count - 1 To 1 Step -1
    Unload mnuHighlighter(i)
  Next i
  iLngCount = 0 ' Set our Lang count to 0 since we've just
                ' erased all languages from the menu
End Sub

Private Sub mnuBookmarks_Click()
  mnuBookmarks.Checked = Not mnuBookmarks.Checked
  sciMain.ShowFlags = mnuBookmarks.Checked
End Sub

Private Sub mnuClearBookmark_Click()
  sciMain.ClearBookmarks
End Sub

Private Sub mnuCopy_Click()
  sciMain.Copy
End Sub

Private Sub mnuCut_Click()
  sciMain.Cut
End Sub

Private Sub mnuDateTime_Click()
  sciMain.SelText = Date & " / " & Time
End Sub

Private Sub mnuExit_Click()
  Unload Me
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

Private Sub mnuGoto_Click()
  sciMain.DoGoto
End Sub

Private Sub mnuLineNumbers_Click()
  mnuLineNumbers.Checked = Not mnuLineNumbers.Checked
  sciMain.LineNumbers = mnuLineNumbers.Checked
End Sub

Private Sub mnuNew_Click()
  sciMain.Text = ""
  sciMain.SetSavePoint
End Sub

Private Sub mnuNextBookmark_Click()
  sciMain.NextBookmark
End Sub

Private Sub mnuOpen_Click()
  Dim strFile As String
  With cmndlg
    OpenFile
    If Len(.filename) = 0 Then Exit Sub
    sciMain.LoadFile .filename
    hlMain.SetHighlighterExt sciMain, .filename
  End With
End Sub

Private Sub mnuPaste_Click()
  sciMain.Paste
End Sub

Private Sub mnuPrevBookmark_Click()
  sciMain.PrevBookmark
End Sub

Private Sub mnuPrint_Click()
  sciMain.PrintDoc
End Sub

Private Sub mnuRedo_Click()
  sciMain.Redo
End Sub

Private Sub mnuReplace_Click()
  sciMain.DoReplace
End Sub

Private Sub mnuSelectAll_Click()
  sciMain.SelectAll
End Sub

Private Sub mnuStatusBar_Click()
  mnuStatusBar.Checked = Not mnuStatusBar.Checked
  stbMain.Visible = mnuStatusBar.Checked
  Form_Resize
  sciMain.SetFocus
End Sub

Private Sub mnuSyntax_Click()
  If hlMain.DoOptions(App.Path & "\highlighters") Then
    hlMain.SetStylesAndOptions sciMain, sciMain.CurHigh
    hlMain.WriteSettings "Software\SCIVBNote\SyntaxOptions\"
  End If
  sciMain.SetFocus
End Sub

Private Sub mnuToggle_Click()
  sciMain.ToggleMarker
End Sub

Private Sub mnuUndo_Click()
  sciMain.Undo
End Sub

Private Sub mnuWordWrap_Click()
  mnuWordWrap.Checked = Not mnuWordWrap.Checked
  sciMain.WordWrap = IIf(mnuWordWrap.Checked, wrap, noWrap)
End Sub

Private Sub sciMain_FindFailed(FindText As String)
  Beep
End Sub

Private Sub sciMain_OnError(Number As String, Description As String)
'  MsgBox "Error Number: " & Number & vbCrLf & vbCrLf & Description, vbOKOnly, "Error: " & Number
End Sub

  
Private Sub sciMain_UpdateUI()
  stbMain.PanelText(1) = "Line: " & sciMain.GetCurrentLine + 1 & " Col: " & sciMain.GetColumn
End Sub

Public Function AddMenu(sCaption As String, sTag As String, iIndex As Integer) As Integer
  Dim i As Long
  On Error Resume Next
  For i = 0 To iIndex - 1
    If mnuHighlighter(i).Caption = sCaption Then Exit Function
  Next i
  If iIndex > 0 Then Load mnuHighlighter(iIndex)
  mnuHighlighter(iIndex).Caption = sCaption ' sCaption we got from the "Identify" function on the plugin
  mnuHighlighter(iIndex).Visible = True
  mnuHighlighter(iIndex).Enabled = True
  mnuHighlighter(iIndex).Tag = sTag ' We store the interface to the plugin in here, to later use it on the event of a menu click

End Function

