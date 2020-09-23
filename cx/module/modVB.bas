Attribute VB_Name = "modVB"
Option Explicit
'********************************************************************************
'* This has all the code required to load a vb project/project group up. The    *
'* actual opening of the file will be handled by on the treeview on the main    *
'* form, and isn't neccisary here. This will just open the selected project     *
'* or project group, parse it, and any neccisary files and add them to the      *
'* project treeview on the main form.                                           *
'********************************************************************************

Private ProjectGroup As Boolean

Public Const PROJECT_EXTENSIONS = "*.vbp;*.vbg;*.vbproj;*.cep"

Public Sub LoadVBProject(strFile As String, tvMain As TreeView)
  Dim HasModule As Boolean, HasForm As Boolean
  Dim HasUser As Boolean, HasRes As Boolean, RetStr As String, nNode As Node
  Dim FindItem As Long, fFile As Long, ProjData As String, ProjName As String
  Dim FileDir As String, InsKey As String, StrLeft As String, ModPath As String
  Dim HasClass As Boolean
  If Dir(strFile) = "" Then Exit Sub 'This file does not exist
  'First thing to do is load the project into a variable.
  fFile = FreeFile()
  Open strFile For Input As #fFile
    ProjData = Input(LOF(fFile), fFile)
  Close #fFile
  'Next we want to get ahold of the name of the project
  FindItem = InStr(1, ProjData, "Name=")
  If FindItem = 0 Then
    'This isn't a valid project file apparently.
    Exit Sub
  End If

  
  'Clear the project list
  If ProjectGroup = False Then tvMain.Nodes.Clear 'Clear the treeview
  'if this is not a project group. Otherwise we don't have to worry about it
  'because the project group code will handle that
  RetStr = StripGarbage(Mid$(ProjData, FindItem, InStr(FindItem, ProjData, vbCrLf) - FindItem))
  'Now we add this in as a project
  InsKey = "Project" & RetStr
  If ProjectGroup = False Then
    Set nNode = tvMain.Nodes.Add(, , InsKey, RetStr, 2)
    'The project group is false so we set this as the first node
  Else
    Set nNode = tvMain.Nodes.Add("Group", tvwChild, InsKey, RetStr, 2)
  End If
  ProjName = RetStr
  'Lets get the filedir setup
  FindItem = InStrRev(strFile, "\")

  If FindItem = 0 Then Exit Sub 'Shouldn't ever happen but best to be safe
  FileDir = Mid(strFile, 1, FindItem)
  'Now we just loop through the lines of the project, and add what we need.
  fFile = FreeFile()
  Open strFile For Input As #fFile
    Do While Not EOF(fFile)
      Input #fFile, ProjData
      FindItem = InStr(1, ProjData, "=")
      If FindItem <> 0 Then
        StrLeft = Left(ProjData, FindItem - 1)
        Select Case StrLeft
          Case "Module"
            
            If HasModule = False Then
              InsKey = "Module" & ProjName
              tvMain.Nodes.Add nNode, tvwChild, InsKey, "Modules", 12
              HasModule = True
            End If
            
            ModPath = FileDir & ModuleData(ProjData, RetStr)
            InsKey = "Module" & ProjName
            tvMain.Nodes.Add InsKey, tvwChild, ModPath, RetStr & " (" & StripPath(ModPath) & ")", 4
          Case "Form"
            ModPath = FileDir & GetFormFile(ProjData)
            RetStr = GetVBFormName(ModPath)
            If HasForm = False Then
              InsKey = "Form" & ProjName
              tvMain.Nodes.Add nNode, tvwChild, InsKey, "Forms", 12
              HasForm = True
            End If
            InsKey = "Form" & ProjName
            tvMain.Nodes.Add InsKey, tvwChild, ModPath, RetStr & " (" & StripPath(ModPath) & ")", 3
          Case "Class"
            If HasClass = False Then
              InsKey = "Class" & ProjName
              tvMain.Nodes.Add nNode, tvwChild, InsKey, "Class Modules", 12
              HasClass = True
            End If
            
            ModPath = FileDir & ModuleData(ProjData, RetStr)
            InsKey = "Class" & ProjName
            tvMain.Nodes.Add InsKey, tvwChild, ModPath, RetStr & " (" & StripPath(ModPath) & ")", 5
          Case "UserControl"
            ModPath = FileDir & GetFormFile(ProjData)
            RetStr = GetVBFormName(ModPath)
            If HasUser = False Then
              InsKey = "USER" & ProjName
              tvMain.Nodes.Add nNode, tvwChild, InsKey, "User Controls", 12
              HasUser = True
            End If
            InsKey = "USER" & ProjName
            tvMain.Nodes.Add InsKey, tvwChild, ModPath, RetStr & " (" & StripPath(ModPath) & ")", 6
          Case "ResFile32"
            ModPath = StripGarbage(ProjData)
            ModPath = Replace(ModPath, "..\..\", "c:\")
            If HasRes = False Then
              InsKey = "RES" & ProjName
              tvMain.Nodes.Add nNode, tvwChild, InsKey, "Related Documents", 12
              HasRes = True
            End If
            InsKey = "RES" & ProjName
            tvMain.Nodes.Add InsKey, tvwChild, ModPath, "(" & StripPath(ModPath) & ")", 7
        End Select
      End If
    Loop
  Close #fFile
  nNode.Expanded = True
End Sub

Private Function StripGarbage(str As String) As String
  Dim p As Long, s As Long
  p = InStr(1, str, Chr(34))
  If p = 0 Then Exit Function
  s = InStr(p + 1, str, Chr(34))
  If s = 0 Then Exit Function
  StripGarbage = Mid$(str, p + 1, s - 1 - p)
End Function

Private Function ModuleData(str As String, ModName As String) As String
  'This one will return the module name and file local
  'Works for modules and class modules
  Dim p As Long, s As Long
  p = InStr(1, str, "=")
  If p = 0 Then Exit Function
  s = InStr(1, str, ";")
  If s = 0 Then Exit Function
  ModName = Mid$(str, p + 1, s - 1 - p)
  ModuleData = Mid$(str, s + 2, Len(str) - s + 2)
End Function


Private Function GetVBFormName(frmFile As String) As String
  Dim fFile As Integer, FindItem As Long
  fFile = FreeFile()
  Open frmFile For Input As #fFile
    GetVBFormName = Input(LOF(fFile), fFile)
  Close #fFile
  FindItem = InStr(1, GetVBFormName, "Attribute VB_Name = ")
  If FindItem = 0 Then  'It's an invalid form file
    GetVBFormName = ""
    Exit Function
  End If
  GetVBFormName = StripGarbage(Mid$(GetVBFormName, FindItem, InStr(FindItem, GetVBFormName, vbCrLf) - FindItem))
End Function

Public Function GetFormFile(strdata As String) As String
  Dim p As Long
  p = InStr(1, strdata, "=")
  If p = 0 Then Exit Function
  GetFormFile = Mid$(strdata, p + 1, Len(strdata) - p)
End Function

Public Sub LoadVBGroup(strFile As String, tvMain As TreeView)
  Dim fFile As Long, ProjData As String, StrLeft As String
  Dim FindItem As Long, FileDir As String
  If Dir(strFile) = "" Then Exit Sub  'This file does not exist
  fFile = FreeFile()  'Set fFile to a free file
  ProjectGroup = True  'This is a project group so handle it accordinly
  tvMain.Nodes.Clear  'clear the contents of the project display
  tvMain.Nodes.Add , , "Group", StripPath(strFile), 1
  tvMain.Nodes("Group").Expanded = True
  FindItem = InStrRev(strFile, "\")

  If FindItem = 0 Then Exit Sub 'Shouldn't ever happen but best to be safe
  FileDir = Mid(strFile, 1, FindItem)
  Open strFile For Input As #fFile  'Open it
    Do While Not EOF(fFile)  'Loop through the file
      Input #fFile, ProjData
      FindItem = InStr(1, ProjData, "=")
      If FindItem <> 0 Then
        StrLeft = Left(ProjData, FindItem - 1)
        If StrLeft = "StartupProject" Or StrLeft = "Project" Then
          StrLeft = GetFormFile(ProjData)
          LoadVBProject FileDir & StrLeft, tvMain
        End If
      End If
    Loop
  Close #fFile  'Close it
  ProjectGroup = False
End Sub
