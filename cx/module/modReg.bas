Attribute VB_Name = "modReg"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_CURRENT_USER = &H80000001


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
'MakeFileType "txt", "Text Document", "C:\windows\notepad.exe,0", "open", "C:\windows\notepad.exe %1", False, True

Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function


Private Function GetString(hKey As Long, strPath As String, strValue As String, DefaultStr As Long) As String
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    RegOpenKey hKey, strPath, keyhand
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
    If strBuf = "" Then GetString = DefaultStr
End Function


Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    RegCreateKey hKey, strPath, keyhand
    RegSetValueEx keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata)
    RegCloseKey keyhand
End Sub




'This is the section to read\write all the options :)

Public Sub WriteOptions()
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "comment", frmDoc.sciMain.GetColor(cmClrComment)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "bookmark", frmDoc.sciMain.GetColor(cmClrBookmark)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "bookmarkbk", frmDoc.sciMain.GetColor(cmClrBookmarkBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "commentbk", frmDoc.sciMain.GetColor(cmClrCommentBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "divider", frmDoc.sciMain.GetColor(cmClrHDividerLines)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "highlight", frmDoc.sciMain.GetColor(cmClrHighlightedLine)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "keyword", frmDoc.sciMain.GetColor(cmClrKeyword)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "keywordbk", frmDoc.sciMain.GetColor(cmClrKeywordBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "left", frmDoc.sciMain.GetColor(cmClrLeftMargin)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "linenum", frmDoc.sciMain.GetColor(cmClrLineNumber)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "linenumbk", frmDoc.sciMain.GetColor(cmClrLineNumberBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "number", frmDoc.sciMain.GetColor(cmClrNumber)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "numberbk", frmDoc.sciMain.GetColor(cmClrNumberBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "operator", frmDoc.sciMain.GetColor(cmClrOperator)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "operatorbk", frmDoc.sciMain.GetColor(cmClrOperatorBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "scope", frmDoc.sciMain.GetColor(cmClrScopeKeyword)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "scopebk", frmDoc.sciMain.GetColor(cmClrScopeKeywordBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "string", frmDoc.sciMain.GetColor(cmClrString)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "stringbk", frmDoc.sciMain.GetColor(cmClrStringBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagattrib", frmDoc.sciMain.GetColor(cmClrTagAttributeName)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagattribbk", frmDoc.sciMain.GetColor(cmClrTagAttributeNameBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagele", frmDoc.sciMain.GetColor(cmClrTagElementName)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagelebk", frmDoc.sciMain.GetColor(cmClrTagElementNameBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagent", frmDoc.sciMain.GetColor(cmClrTagEntity)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagentbk", frmDoc.sciMain.GetColor(cmClrTagEntityBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagtxt", frmDoc.sciMain.GetColor(cmClrTagText)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagtxtbk", frmDoc.sciMain.GetColor(cmClrTagTextBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "text", frmDoc.sciMain.GetColor(cmClrText)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "textbk", frmDoc.sciMain.GetColor(cmClrTextBk)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "vdivider", frmDoc.sciMain.GetColor(cmClrVDividerLines)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "window", frmDoc.sciMain.GetColor(cmClrWindow)
'  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "selbounds", frmDoc.sciMain.SelBounds
'  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "numbering", frmDoc.sciMain.LineNumbering
'  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "lttips", frmDoc.sciMain.LineToolTips
'  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "numberingstyle", frmDoc.sciMain.LineNumberStyle
'  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "numberingstart", frmDoc.sciMain.LineNumberStart
'  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "leftmargin", frmDoc.sciMain.DisplayLeftMargin
'  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "bold", frmDoc.sciMain.Font.Bold
'  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "italic", frmDoc.sciMain.Font.Italic
'  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "size", frmDoc.sciMain.Font.Size
'  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "strike", frmDoc.sciMain.Font.Strikethrough
'  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "under", frmDoc.sciMain.Font.Underline
'  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "name", frmDoc.sciMain.Font.Name
  'savestring HKEY_CLASSES_ROOT, "cEdit\data\", "leftmargin", frmdoc.sciMain.
End Sub


Public Sub SaveFormData(frm As Form)
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Width", frm.Width
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Height", frm.Height
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Top", frm.Top
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Left", frm.Left
End Sub

Public Sub LoadFormData(frm As Form)
  frm.Width = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Width", frm.Width)
  frm.Height = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Height", frm.Height)
  frm.Top = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Top", frm.Top)
  frm.Left = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Left", frm.Left)
End Sub

'Window Data

Public Sub WriteData()
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "windowstate", frmMain.WindowState
  frmMain.WindowState = vbNormal
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "left", frmMain.Left
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "top", frmMain.Top
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "width", frmMain.Width
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "height", frmMain.Height
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "toolbar", frmMain.tBar.Visible
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "mactoolbar", frmMain.tbMacro.Visible
  'SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "statusbar", frmMain.stBar.Visible
  'SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "quicknav", frmMain.Picture2.Visible
  'SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "quicknavwidth", frmMain.Picture2.Width
End Sub

Public Sub ReadData()
  Dim m As Boolean
  frmMain.Left = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "left", 1980)
  frmMain.Top = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "top", 1980)
  frmMain.Width = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "width", 10080)
  frmMain.Height = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "height", 5640)
  frmMain.WindowState = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "windowstate", 0)
  'frmMain.Picture2.Width = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "quicknavwidth", 3005)
'  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "quicknav", True)
'  frmMain.quicknav.Checked = m
  'frmMain.Picture2.Visible = m
  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "toolbar", True)
  frmMain.tBar.Visible = m
  'frmMain.Toolbar.Checked = m
  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "mactoolbar", True)
  frmMain.tbMacro.Visible = m
  'frmMain.mnuMacBar.Checked = m
  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "statusbar", True)
  'frmMain.stBar.Visible = m
  'frmMain.statusbar2.Checked = m
End Sub

Public Sub WriteInput()
'  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "whitespace", frmMain.whitespace.Checked
'  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "hlline", frmMain.hlline.Checked
End Sub

Public Sub ReadInput()
'  WhiteSpaced = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "whitespace", False)
'  frmMain.whitespace.Checked = WhiteSpaced
'  frmDoc.sciMain.DisplayWhitespace = WhiteSpaced
'  HighLight = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "hlline", False)
'  frmMain.hlline.Checked = HighLight
End Sub

