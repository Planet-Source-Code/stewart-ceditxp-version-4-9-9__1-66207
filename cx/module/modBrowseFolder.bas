Attribute VB_Name = "modBrowseFolder"
Option Explicit
'----------------------------------------------------------------
'
'                Browse for folders in VB5
'
'              written by D. Rijmenants 2004
'
'----------------------------------------------------------------
'
' This module enables you to get use the browse dialog to select
' a folder in vb5. Only one functionis required.  As you call
' the function, the browse dialog pops up. Easy to apply !
'
' return = BrowseFolder(Title, MyForm)
'
' Where:
'
' Title  (string) is the title you want to display on the dialog
' MyForm (form) is the form on wich you call the dialog
' return (string) the path of the selected folder after pressing OK
'
' Note: if cancel is selected, the dialog will return
'       an empty string !
'
'
' That's all folks...
'
' Comments or suggestions are most welcome at
' mail: dr.defcom@telenet.be
'
'----------------------------------------------------------------
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Public Function BrowseFolder(ByVal aTitle As String, ByVal aForm As Form) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, path$, pos%
Dim BrowsePath As String
Dim Browse As String
Dim t As Long
bInfo.hOwner = aForm.hwnd
bInfo.lpszTitle = aTitle
'the type of folder(s) to return
bInfo.ulFlags = &H1
'show the dialog box
pidl& = SHBrowseForFolder(bInfo)
'set the maximum characters
path = Space(512)
'get the selected path
t = SHGetPathFromIDList(ByVal pidl&, ByVal path)
pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
'set the extracted path to SpecIn
BrowseFolder = Left(path$, pos - 1)
'clean up the path string
If Right$(Browse, 1) = "\" Then
    BrowseFolder = BrowseFolder
    Else
    BrowseFolder = BrowseFolder + "\"
End If
If Right(BrowseFolder, 2) = "\\" Then BrowseFolder = Left(BrowseFolder, Len(BrowseFolder) - 1)
If BrowseFolder = "\" Then BrowseFolder = ""
End Function


