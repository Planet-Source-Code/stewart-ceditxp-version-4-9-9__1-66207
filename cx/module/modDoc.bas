Attribute VB_Name = "modDoc"
Option Explicit

Public Type FormState
    Deleted As Boolean
    dirty As Boolean
    Type As Integer
    Color As Long
End Type
Public FState() As FormState
Public fIndex As Integer
Public Document() As New frmDoc
Public dnum As Integer
Public Type Recent
  Recent1 As String
  Recent2 As String
  Recent3 As String
  Recent4 As String
  Recent5 As String
  Recent6 As String
End Type
Public Recnt As Recent

Public Sub InitDocumentInterface()
  ReDim Document(0 To 1)
End Sub

Public Sub doNew(str As String)
  On Error Resume Next
  Dim x As Integer
  fIndex = FindFreeIndex()
  If fIndex = 0 Then
    fIndex = 1
    ReDim Document(1 To 1)
    ReDim FState(1 To 1)
  End If
  Document(fIndex).Caption = "Untitled " & Document(fIndex).Tag
  Document(fIndex).Changed = False
  Document(fIndex).Tag = fIndex
  dnum = fIndex
  'Document(fIndex).Move 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
  'Document(fIndex).WindowState = vbMaximized
  frmMain.Highlighters.SetStylesAndOptions Document(fIndex).sciMain, "HTML"
  For x = 0 To frmMain.mnuHighlighter.Count - 1
    If frmMain.mnuHighlighter(x).Caption = "HTML" Then
      Document(fIndex).SetCheck x
    End If
  Next x
  Document(fIndex).Visible = True
  Document(fIndex).sciMain.SetFocus
End Sub


Public Sub InitDocInterface()
End Sub

Function FindFreeIndex() As Integer
    'On Error GoTo errhandler
    Dim I As Integer
    Dim ArrayCount As Integer
    ArrayCount = UBound(Document)
    For I = 1 To ArrayCount
        If FState(I).Deleted Then
            FindFreeIndex = I
            FState(I).Deleted = False
            Exit Function
        End If
    Next
    ReDim Preserve Document(1 To ArrayCount + 1)
    ReDim Preserve Document(1 To ArrayCount + 1)
    ReDim Preserve FState(1 To ArrayCount + 1)
    FindFreeIndex = UBound(Document)
    Exit Function
errhandler:
    FindFreeIndex = 0
End Function


Public Function DoOpen(path As String) As Boolean
  On Error Resume Next
  Dim x As Long
  
  If Dir(path) = "" Then
    If MsgBox("The file: " & path & Chr(10) & "does not exist. Do you wish to create it?", vbYesNo + vbQuestion, "Create File") = vbNo Then Exit Function
  End If
  
  If IsProject(path) Then
    LoadProject path
'    LoadVBProject path
    Exit Function
  End If
  
  'If there's 0 open docs no need to do the loop to verify the file's not open.
  'but if there are files we wanna make sure that none are the one were about
  'to open (whats the point of opening the same file twice) and if this is the
  'case then we will just setfocus :)
  If fIndex > 0 Then
    'First lets check and find out of this file is open or not
    For x = 1 To UBound(Document)
      If FState(x).Deleted = False Then
        If Document(x).isFile = True And Document(x).FileName = path Then
          Document(x).SetFocus
          DoOpen = False
          Exit Function
        End If
      End If
    Next
  End If
  fIndex = FindFreeIndex()
  If fIndex = 0 Then
    fIndex = 1
    ReDim Document(1 To 1)
    ReDim FState(1 To 1)
  End If
  dnum = fIndex
  Document(fIndex).Caption = StripPath(path)
  Document(fIndex).Tag = fIndex
  Document(fIndex).isFile = True
  Document(fIndex).FileName = path
  Document(fIndex).sciMain.LoadFile path
  Document(fIndex).Changed = False
  frmMain.Highlighters.SetHighlighterExt Document(fIndex).sciMain, path
  Dim I As Integer
  For I = 0 To iLngCount - 1
    If frmMain.mnuHighlighter(I).Caption = Document(fIndex).sciMain.CurHigh Then
      Document(fIndex).SetCheck I
    End If
  Next I
  Document(fIndex).Show
  Document(fIndex).sciMain.SetFocus
  DoOpen = True
End Function

Public Sub OpenFTP(str As String, path As String, FTPDir As String, FTPAccount As String)
  On Error Resume Next
  fIndex = FindFreeIndex()
  Document(fIndex).Changed = False
  Document(fIndex).Tag = fIndex
  Document(fIndex).Caption = path
  Document(fIndex).FileName = path
  Document(fIndex).sciMain.Text = str
  Document(fIndex).sciMain.ClearUndoBuffer
  Document(fIndex).sciMain.SetSavePoint
  frmMain.Highlighters.SetHighlighterExt Document(fIndex).sciMain, path
  Dim I As Integer
  For I = 0 To iLngCount - 1
    If frmMain.mnuHighlighter(I).Caption = Document(fIndex).sciMain.CurHigh Then
      Document(fIndex).SetCheck I
    End If
  Next I
  Document(fIndex).FTP = True
  Document(fIndex).FTPAccount = FTPAccount
  Document(fIndex).FTPDir = FTPDir
  Document(fIndex).Show
End Sub


Public Function GetMDIChildCount() As Long
  Dim frm As Form
  For Each frm In Forms
    If frm.Name = "frmDoc" Then
      GetMDIChildCount = GetMDIChildCount + 1
    End If
  Next
End Function

Public Function IsProject(sFile As String) As Boolean
Dim sExtension As String
    sExtension = GetExtension(sFile)
    If InStr(1, PROJECT_EXTENSIONS & ";", "." & sExtension & ";") Then IsProject = True
End Function

Private Sub LoadProject(path As String)
  'determine what type of file we are dealing with.
  Dim ext As String
  ext = GetExtension(path)
  Select Case ext
    Case "vbp"
      LoadVBProject path, frmMain.tvMain
    Case "vbg"
      LoadVBGroup path, frmMain.tvMain
  End Select
End Sub

Public Sub doSaveAs()
  On Error GoTo errhandler
  Dim msg As VbMsgBoxResult
  frmMain.cd.CancelError = True
  frmMain.cd.FileName = ""
  frmMain.cd.DialogTitle = "Save document... " & Document(dnum).Caption
  frmMain.cd.Filter = strExt ' & FilterB  '"All Files|*.*|Text Files|*.txt|Html Files|*.html;*.htm|Style Sheets|*.css|Java Scripting|*.js|C Files|*.c|C++ Files|*.cpp|C/C++ Header Files|*.h|Perl Files|*.pl|CGI/Perl Files|*.cgi|XML Files|*.xml|Pascal Files|*.pas|Basic Module Files|*.bas|Basic Form Files|*.frm|Basic Project Files|*.vbp|Basic Class Modules|*.cls"
  frmMain.cd.ShowSave
  'If frmMain.cd.filename = "" Then Exit Function
  If FileExists(frmMain.cd.FileName) = True Then
     msg = MsgBox(frmMain.cd.FileName & " Already exists." & Chr(10) & "Do you want to replace it?", vbYesNo + vbQuestion, "Overwrite")
    If msg = vbYes Then
      'continue
    ElseIf msg = vbNo Then
      doSaveAs
    End If
  End If
  Document(dnum).isFile = True
  Document(dnum).sciMain.SaveToFile frmMain.cd.FileName
  Document(dnum).FileName = frmMain.cd.FileName
  Document(dnum).Caption = StripPath(frmMain.cd.FileName)
  frmMain.Highlighters.SetHighlighterExt Document(dnum).sciMain, frmMain.cd.FileName
  AddRecent frmMain.cd.FileName
errhandler:
  If Err.Number = 32755 Or Err.Number = 0 Then
    Exit Sub
  Else
    MsgBox "Error: " & Err.Number & Chr(10) & Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
  End If
  Exit Sub
End Sub


Public Function FileExists(FullFileName As String) As Boolean
    On Error Resume Next
    If Dir(FullFileName) = "" Then
      FileExists = False
    Else
      FileExists = True
    End If
End Function

Public Function ShowSite(URL As String)
  If frmBrowse.Visible = False Then frmBrowse.Visible = True
  frmBrowse.SetFocus
  frmBrowse.www.Navigate URL
End Function

Public Sub AddRecent(str As String)
  Dim FreeFileNum As Integer
  With Recnt
    .Recent6 = .Recent5
    .Recent5 = .Recent4
    .Recent4 = .Recent3
    .Recent3 = .Recent2
    .Recent2 = .Recent1
    .Recent1 = str
  End With
  FreeFileNum = FreeFile()
  Open App.path & "\temp\recent.rct" For Binary Access Write As #FreeFileNum
    Put #FreeFileNum, , Recnt
  Close #FreeFileNum
  With frmMain
    If Recnt.Recent1 <> "" Then
      .mnuRec(0).Caption = Recnt.Recent1
      .mnuRec(0).Visible = True
    End If
    If Recnt.Recent2 <> "" Then
      .mnuRec(1).Caption = Recnt.Recent2
      .mnuRec(1).Visible = True
    End If
    If Recnt.Recent3 <> "" Then
      .mnuRec(2).Caption = Recnt.Recent3
      .mnuRec(2).Visible = True
    End If
    If Recnt.Recent4 <> "" Then
      .mnuRec(3).Caption = Recnt.Recent4
      .mnuRec(3).Visible = True
    End If
    If Recnt.Recent5 <> "" Then
      .mnuRec(4).Caption = Recnt.Recent5
      .mnuRec(4).Visible = True
    End If
    If Recnt.Recent6 <> "" Then
      .mnuRec(5).Caption = Recnt.Recent6
      .mnuRec(5).Visible = True
    End If
  End With

End Sub

