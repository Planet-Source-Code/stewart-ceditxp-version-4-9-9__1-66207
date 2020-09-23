Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Result As String
Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16
Public Const MAX_PATH = 260

Public Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Public Const SHGFI_LARGEICON = &H0       'Large icon
Public Const SHGFI_SMALLICON = &H1       'Small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400

Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal I&, ByVal hDCDest&, _
    ByVal x&, ByVal Y&, ByVal FLAGS&) As Long


'----------------------------------------------------------
'public variables
'----------------------------------------------------------
Public ShInfo As SHFILEINFO

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Sub RedrawWin(hWndA As Long)
  Dim rc As RECT
  GetWindowRect hWndA, rc
  RedrawWindow hWndA, rc, 0, 1
End Sub
