VERSION 5.00
Begin VB.UserControl VSFileSearch 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1695
   ScaleWidth      =   7620
   ToolboxBitmap   =   "VSFileSearch.ctx":0000
   Begin VB.ListBox SearchResults 
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "VSFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------------------------------------------------
'    EasyASP - Web Development Platform
'    Copyright 2001 Eric Banker, Inc. All Rights Reserved.
'    Confidential and proprietary.
' --------------------------------------------------------------------------

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = 201
Private Const EM_LINEINDEX = 187
Private Const EM_LINELENGTH = 193
Private Const LB_SETHORIZONTALEXTENT = &H194

Private tmpFile As New Collection
Private GlobalOccurance As Long
Private LastLineStart As Long
Private LastLineEnd As Long
Private TotalFound As Long

Public Event DblClick(SelectedFile As String, LineNumber As String)

Private Sub UserControl_Initialize()
    Debug.Print vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    SearchResults.Width = UserControl.ScaleWidth
    SearchResults.Height = UserControl.ScaleHeight
End Sub

Private Sub SetScroll()
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long

    For i = 0 To SearchResults.ListCount - 1
        'Debug.Print TextWidth(SearchResults.List(i)) & " > " & TextWidth(SearchResults.List(intGreatestLen))
        If TextWidth(SearchResults.List(i)) > TextWidth(SearchResults.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i

    lngGreatestWidth = TextWidth(SearchResults.List(intGreatestLen) + Space(1))
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    SendMessage SearchResults.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth * 3, 0
End Sub

Private Sub SearchResults_Click()
On Error Resume Next
    If SearchResults.Selected(0) = True Then
        SearchResults.Selected(0) = False
    End If
    
    If SearchResults.Selected(SearchResults.ListCount - 1) = True Then
        SearchResults.Selected(SearchResults.ListCount - 1) = False
    End If
End Sub

Private Sub SearchResults_DblClick()
On Error GoTo errhandler
    Dim tempString1 As String, tempString2 As String
    Dim SelectedFilename As String, SelectedLine As String
    Dim LineEnd As Long, LineStart As Long, LineLength As Long, FileEnd As Long
    
    tempString1 = SearchResults.Text
    tempString2 = SearchResults.Text
    
    FileEnd = InStr(1, tempString1, "(")
    SelectedFilename = Mid$(tempString1, 1, FileEnd - 2)

    LineStart = InStr(1, tempString2, "(")
    LineEnd = InStr(1, tempString2, ")")
    LineLength = LineEnd - LineStart
    SelectedLine = Mid$(tempString2, LineStart + 1, LineLength - 1)

    RaiseEvent DblClick(SelectedFilename, SelectedLine)
    
errhandler:
    Exit Sub
End Sub

Public Sub FindInFiles(UserDir As String, SearchString As String, MatchCase As Boolean)
    Dim i As Long
    
    SearchResults.Clear
    GlobalOccurance = 0
    TotalFound = 0
    
    SearchResults.AddItem "Searching for '" & SearchString & "'...", TotalFound

    If Right(UserDir, 1) = "\" Then
        Call ListFiles(UserDir)
    Else
        Call ListFiles(UserDir & "\")
    End If
    
    DoEvents
    
    For i = 1 To tmpFile.Count
        Call TraverseFile(tmpFile(i), SearchString, MatchCase)
    Next
    
    SearchResults.AddItem GlobalOccurance & " occurrence(s) have been found.", TotalFound + 1
    
    Call SetScroll
End Sub

Private Sub ListFiles(ByVal Pathname As String)
On Error Resume Next
    Dim Count, i, FileName
    
    Count = 0
    FileName = Dir(Pathname)

    Do While tmpFile.Count > 0
        tmpFile.Remove 1
    Loop
    
    Do While Not FileName = ""
        If Not FileName = "." And Not FileName = ".." Then
            If Not GetAttr(Pathname & FileName) And vbDirectory Then
                tmpFile.Add Pathname & FileName
                Count = Count + 1
            End If
        End If
        FileName = Dir
    Loop
End Sub

Private Sub TraverseFile(FileName As String, SearchString As String, MatchCase As Boolean)
    Dim intCount As Long
    Dim lngTotal As Long
    
    Dim myline As String, tmpString As String
    Dim tmp As Long
    Dim fs As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    tmp = 0
    
    Set fs = New Scripting.FileSystemObject
    Set ts = fs.OpenTextFile(FileName, ForReading)
    
    Do Until ts.AtEndOfStream
        tmp = tmp + 1
        myline = ts.ReadLine
        If MatchCase = True Then
            If InStr(myline, SearchString) > 0 Then
                TotalFound = TotalFound + 1
                tmpString = FileName & " (" & tmp & "):  " & myline
                SearchResults.AddItem tmpString, TotalFound
                intCount = intCount + 1
            End If
        Else
            If InStr(UCase(myline), UCase(SearchString)) > 0 Then
                TotalFound = TotalFound + 1
                tmpString = FileName & " (" & tmp & "):  " & myline
                SearchResults.AddItem tmpString, TotalFound
                intCount = intCount + 1
            End If
        End If
    Loop
    GlobalOccurance = GlobalOccurance + intCount
    
Exit Sub

errhandler:
    Exit Sub
End Sub
