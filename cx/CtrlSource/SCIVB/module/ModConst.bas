Attribute VB_Name = "ModConst"
Option Explicit

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_HOME = &H24
Public Const VK_DOWN = &H28
Public Const VK_END = &H23
Public Const VK_UP = &H26


Public Type CharRange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Public Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Public Const WM_USER As Long = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4


Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_PRESSED As Integer = &H1000

Public Declare Function GetDeviceCaps Lib "gdi32" ( _
   ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const VERTSIZE = 6           '  Vertical size in millimeters
Public Const HORZSIZE = 4           '  Horizontal size in millimeters
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOCOPYBITS = &H100



Public Const MK_CONTROL = &H8
Public Const MK_SHIFT = &H4



Public Const VK_SHIFT = &H10&
Public Const VK_CONTROL = &H11&
Public Const VK_MENU = &H12& ' Alt key
