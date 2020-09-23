Attribute VB_Name = "modTags"
Option Explicit

Public Sub addTags()
  Dim fFile As Integer, tmpStr As String, tmpKey As String
  fFile = FreeFile
  Open App.path & "\temp\tags.dat" For Input As #fFile
    Do Until EOF(fFile)
      Input #fFile, tmpStr
      If Left(tmpStr, 1) <> "+" Then
        frmMain.TagsD.Nodes.Add , , tmpStr, tmpStr, 1
        tmpKey = tmpStr
      Else
        frmMain.TagsD.Nodes.Add tmpKey, tvwChild, Mid(tmpStr, 2, Len(tmpStr) - 1) & frmMain.TagsD.Nodes.Count, Mid(tmpStr, 2, Len(tmpStr) - 1), 2
      End If
    Loop
  Close #fFile
End Sub

