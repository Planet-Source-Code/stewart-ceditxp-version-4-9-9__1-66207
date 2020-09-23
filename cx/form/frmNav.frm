VERSION 5.00
Begin VB.Form frmNav 
   Caption         =   "Quick Nav"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Icon Sizes in pixels

Private Sub Form_Resize()
  On Error Resume Next
  'Tbs.Move 30 + frmMain.fDock.DockedFormCaptionOffsetLeft("frmNav"), 30 + frmMain.fDock.DockedFormCaptionOffsetTop("frmNav"), Me.ScaleWidth - 60 - frmMain.fDock.DockedFormCaptionOffsetLeft("frmNav") - 60, Me.ScaleHeight - 60 - frmMain.fDock.DockedFormCaptionOffsetTop("frmNav") - 60
  Picture4.Move Tbs.ClientLeft, Tbs.ClientTop, Tbs.ClientWidth, Tbs.ClientHeight
  Picture5.Move Tbs.ClientLeft, Tbs.ClientTop, Tbs.ClientWidth, Tbs.ClientHeight
  picSnippet.Move Tbs.ClientLeft, Tbs.ClientTop, Tbs.ClientWidth, Tbs.ClientHeight
  TagsD.Move 0, 30, Picture5.ScaleWidth, Picture5.ScaleHeight - 30
End Sub

Private Sub Dir1_Change()

End Sub






Private Sub imgSize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picSize.Visible = True
End Sub

Private Sub imgSize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim nxtY As Long
  If Button = 1 Then
    nxtY = (imgSize.Top + Y)
    If nxtY < 800 Then nxtY = 800
    If nxtY > (Picture4.ScaleHeight - 800) Then nxtY = Picture4.Height - 800
    picSize.Top = nxtY
    imgSize.Move picSize.Left, picSize.Top, picSize.Width, picSize.Height
  End If
End Sub

Private Sub imgSize_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picSize.Visible = False
  Resize
End Sub

Private Sub Resize()
  On Error Resume Next
  imgSize.Left = 0
  imgSize.Width = Picture4.ScaleWidth
  picSize.Move 0, imgSize.Top, imgSize.Width, imgSize.Height
  Drive1.Move 0, 30, Picture4.ScaleWidth
  Dir1.Move 0, Drive1.Top + Drive1.Height + 30, Picture4.ScaleWidth, imgSize.Top - Dir1.Top
  If Dir1.Height > (Picture4.ScaleHeight - 1500) Then Dir1.Height = Picture4.ScaleHeight - 1500
  imgSize.Move 0, Dir1.Top + Dir1.Height, Picture4.ScaleWidth
  File1.Move 0, imgSize.Top + imgSize.Height, Picture4.ScaleWidth, Picture4.Height - (imgSize.Top + imgSize.Height)
End Sub
 
Private Sub Picture1_Click()

End Sub

Private Sub Picture1_Resize()

End Sub

Private Sub lstSnippet_DblClick()
End Sub

Private Sub lstSnippet_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
    Dim OLEFilename As String, ext As String, file2 As String
    Dim I As Integer
    For I = 1 To Data.Files.Count
        If Data.GetFormat(vbCFFiles) Then
            OLEFilename = Data.Files(I)
        End If
        On Error GoTo errexit
       ext = GetExtension(OLEFilename)
       
       ext = Left(OLEFilename, Len(OLEFilename) - (Len(ext) + 1))
       file2 = StripPath(ext)
       CopyFile OLEFilename, App.path & "\snippets\" & file2 & ".snippet", False
    Next I
    AddSnippets
errexit:
    Exit Sub
End Sub

Private Sub lstSnippet_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
  On Error Resume Next
    If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub

Private Sub picSnippet_Resize()
  lstSnippet.Move 0, 0, picSnippet.ScaleWidth, picSnippet.ScaleHeight
End Sub

Private Sub Picture4_Resize()
  Resize
End Sub

Private Sub tagsd_DblClick()
  Dim timedate As String
  On Error Resume Next
'  Dim r As CodeSenseCtl.range
'  Set r = New CodeSenseCtl.range
'  timedate = TagsD.SelectedItem.Text
'  Document(dnum).sciMain.SelText = timedate
'  Set r = Document(dnum).sciMain.GetSel(False)
'  Document(dnum).sciMain.SetCaretPos r.StartLineNo + 1, r.StartColNo + Len(timedate)
'  Document(dnum).sciMain.SetFocus
End Sub

Private Sub tbs_Click()
  Picture4.Visible = False
  Picture5.Visible = False
  picSnippet.Visible = False
  If Tbs.SelectedItem.Index = 1 Then
    Picture4.Visible = True
  ElseIf Tbs.SelectedItem.Index = 2 Then
    Picture5.Visible = True
  ElseIf Tbs.SelectedItem.Index = 3 Then
    picSnippet.Visible = True
  End If
End Sub

