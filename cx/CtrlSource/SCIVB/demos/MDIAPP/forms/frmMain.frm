VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.1#0"; "SCIVBX.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Source Edit"
   ClientHeight    =   5025
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6795
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin SCIVBX.SCIHighlighter hlMain 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3360
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MDITest.ucStatusbar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   4650
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   661
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
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export to HTML"
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
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFoldAll 
         Caption         =   "&Fold All"
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

Private lDocCnt As Long
Private iLngCount As Integer
Dim strFilter As String

Private Function ActiveDoc() As frmDoc
  On Error Resume Next
  Set ActiveDoc = ActiveForm
End Function

Private Sub hlMain_AddHighlighter(HighlighterName As String, Filter As String)
  AddMenu HighlighterName, HighlighterName, iLngCount
  strFilter = strFilter & Filter
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

Private Sub MDIForm_Activate()
  On Error Resume Next
  ActiveForm.sciMain.SetFocus
  SetupStatus
End Sub

Private Sub MDIForm_Load()
  lDocCnt = 0
  hlMain.ReadSettings "\Software\MDITest\Syntax\Settings\"
  hlMain.LoadHighlighters App.Path & "\highlighters"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  hlMain.WriteSettings "\Software\MDITest\Syntax\Settings\"
End Sub

Private Sub mnuBookmarks_Click()
  On Error Resume Next
  mnuBookmarks.Checked = Not mnuBookmarks.Checked
  ActiveDoc.sciMain.ShowFlags = mnuBookmarks.Checked
End Sub

Private Sub mnuClearBookmark_Click()
  On Error Resume Next
  ActiveDoc.sciMain.ClearBookmarks
End Sub

Private Sub mnuCopy_Click()
  On Error Resume Next
  ActiveDoc.sciMain.Copy
End Sub

Private Sub mnuCut_Click()
  On Error Resume Next
  ActiveDoc.sciMain.Cut
End Sub

Private Sub mnuDateTime_Click()
  On Error Resume Next
  ActiveDoc.sciMain.SelText = Date & "\" & Time
End Sub

Private Sub mnuExport_Click()
  On Error GoTo errHandle
  With CD
    .CancelError = True
    .Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
    .ShowSave
    hlMain.ExportToHTML .FileName, ActiveDoc.sciMain
  End With
errHandle:
  Exit Sub
End Sub

Private Sub mnuFind_Click()
  On Error Resume Next
  ActiveDoc.sciMain.DoFind
End Sub

Private Sub mnuFindNext_Click()
  On Error Resume Next
  ActiveDoc.sciMain.FindNext
End Sub

Private Sub mnuFindPrev_Click()
  On Error Resume Next
  ActiveDoc.sciMain.FindPrev
End Sub

Private Sub mnuFoldAll_Click()
  ActiveDoc.sciMain.FoldAll
End Sub

Private Sub mnuGoto_Click()
  On Error Resume Next
  ActiveDoc.sciMain.DoGoto
End Sub

Private Sub mnuHighlighter_Click(Index As Integer)
  On Error Resume Next
  Call hlMain.SetStylesAndOptions(ActiveForm.sciMain, mnuHighlighter(Index).Caption)
End Sub

Private Sub mnuLineNumbers_Click()
  On Error Resume Next
  mnuLineNumbers.Checked = Not mnuLineNumbers.Checked
  ActiveDoc.sciMain.LineNumbers = mnuLineNumbers.Checked
End Sub

Private Sub mnuNew_Click()
  Dim frm As New frmDoc
  frm.Caption = "New Doc " & lDocCnt
  frm.Show
  'hlMain.SetStylesAndOptions frm.sciMain, "CPP"
  hlMain.SetHighlighter frm.sciMain, "VB"
  lDocCnt = lDocCnt + 1
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

Private Sub mnuNextBookmark_Click()
  On Error Resume Next
  ActiveDoc.sciMain.NextBookmark
End Sub

Private Sub mnuOpen_Click()
  Dim strFile As String
  On Error GoTo errHandler
  With CD
    .CancelError = True
    .Filter = strFilter
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    Dim frm As New frmDoc
    Load frm
    hlMain.SetHighlighterExt frm.sciMain, .FileName
    frm.sciMain.LoadFile .FileName
    frm.strFile = .FileName
    frm.Caption = .FileName
    frm.Show
  End With
errHandler:
  Exit Sub
End Sub

Private Sub mnuPaste_Click()
  On Error Resume Next
  ActiveDoc.sciMain.Paste
End Sub

Private Sub mnuPrevBookmark_Click()
  On Error Resume Next
  ActiveDoc.sciMain.PrevBookmark
End Sub

Private Sub mnuRedo_Click()
  On Error Resume Next
  ActiveDoc.sciMain.Redo
End Sub

Private Sub mnuReplace_Click()
  On Error Resume Next
  ActiveDoc.sciMain.DoReplace
End Sub

Private Sub mnuSave_Click()
  On Error Resume Next
  If ActiveDoc.strFile <> "" Then
    ' We have a filename so save to it
    Call ActiveDoc.sciMain.SaveToFile(ActiveDoc.strFile)
  Else
    ' No filename so do save as
    mnuSaveAs_Click
  End If
End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo errHandler
  Dim msgRes As VbMsgBoxResult
  With CD
    .Filter = strFilter
showIt:     ' I know this is considered bad practice but this is
            ' a generic sample.
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    If Dir(.FileName) <> "" Then
      msgRes = MsgBox("Filename: " & .FileName & " already exists.  Do you wish to overwrite?", vbYesNoCancel + vbQuestion, "OverWrite?")
      If msgRes = vbYes Then
        Call ActiveDoc.sciMain.SaveToFile(.FileName)
        ActiveDoc.strFile = .FileName
        ActiveDoc.Caption = .FileName
        Call hlMain.SetHighlighterExt(ActiveDoc.sciMain, .FileName)
      ElseIf msgRes = vbNo Then
        GoTo showIt
      Else
        Exit Sub
      End If
    End If
    ActiveDoc.sciMain.SaveToFile .FileName
    ActiveDoc.strFile = .FileName
    ActiveDoc.Caption = .FileName
    Call hlMain.SetHighlighterExt(ActiveDoc.sciMain, .FileName)
  End With
errHandler:
  Exit Sub
End Sub

Private Sub mnuSelectAll_Click()
  On Error Resume Next
  ActiveDoc.sciMain.SelectAll
End Sub

Private Sub mnuStatusBar_Click()
  mnuStatusBar.Checked = Not mnuStatusBar.Checked
  stbMain.Visible = mnuStatusBar.Checked
End Sub

Private Sub mnuSyntax_Click()
  On Error Resume Next
  Dim frm As Form
  strFilter = ""
  If hlMain.DoOptions(App.Path & "\highlighters") Then
    For Each frm In Forms
      If frm.Name = "frmDoc" Then
        Call hlMain.SetStylesAndOptions(frm.sciMain, frm.sciMain.CurHigh)
      End If
    Next
    hlMain.WriteSettings "Software\SCIVBNote\SyntaxOptions\"
  End If
  
End Sub

Private Sub mnuToggle_Click()
  On Error Resume Next
  ActiveDoc.sciMain.ToggleMarker
End Sub

Private Sub mnuUndo_Click()
  On Error Resume Next
  ActiveDoc.sciMain.Undo
End Sub

Private Sub mnuWordWrap_Click()
  On Error Resume Next
  mnuWordWrap.Checked = Not mnuWordWrap.Checked
  If mnuWordWrap.Checked Then
    ActiveForm.sciMain.WordWrap = wrap
  Else
    ActiveDoc.sciMain.WordWrap = noWrap
  End If
End Sub
