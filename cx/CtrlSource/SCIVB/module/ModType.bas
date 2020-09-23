Attribute VB_Name = "ModType"
Option Explicit

Public Const WM_NOTIFY = &H4E

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units

Public Const WM_COMMAND = &H111
Public Const WM_CLOSE = &H10
Public Const ALL_MESSAGES = -1
Public Const WM_SETFOCUS = &H7


Public Const WM_ACTIVATE = &H6

Public Const SC_CP_UTF8 = 65001

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_LBUTTONUP = &H202
Public Const WM_CHAR = &H102
Public Const WM_KEYDOWN = &H100
Public Const WM_MOUSEMOVE = &H200
Public Const WM_KEYUP = &H101
Public Const GWL_WNDPROC = (-4)
Public Const MK_RBUTTON = &H2
Public Const MK_LBUTTON = &H1
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_CLIPCHILDREN = &H2000000





Type NMHDR
    hwndFrom As Long
    idFrom As Long
    Code As Long
End Type

Public Type SCNotification
    NotifyHeader As NMHDR
    Position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    X As Long
    Y As Long
End Type

Public Const CB_FINDSTRING = &H14C

