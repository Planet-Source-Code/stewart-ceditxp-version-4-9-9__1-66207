VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GHook"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Interface to connect to and use the vbAccelerator Hook Library functions."
Option Explicit

' ======================================================================================
' Name:     vbAccelerator Windows Hook Library
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     25 August 1999
'
' Requires:
'
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Public Interface to mWindowsHook.bas from outside
' vbalWinHook.DLL. Also provides classes to interpret
' the lParam member for most hooks.
'
' ======================================================================================

' Interface to mWindowsHook for DLL:
Public Enum EHTHookTypeConstants
   [_WH_MIN] = -1
   WH_CALLWNDPROC = 4
   WH_CBT = 5
   WH_DEBUG = 9
   WH_FOREGROUNDIDLE = 11
   WH_GETMESSAGE = 3
   'WH_HARDWARE = 8 ' Not implemented in Win32
   WH_JOURNALRECORD = 0
   WH_JOURNALPLAYBACK = 1
   WH_KEYBOARD = 2
   WH_MOUSE = 7
   WH_MSGFILTER = (-1)
   WH_SHELL = 10
   WH_SYSMSGFILTER = 6
   WH_CALLWNDPROCRET = 12
   [_WH_MAX] = 14
End Enum

Public Enum EHTMsgFilterCodes
   MSGF_DDEMGR = &H8001
   MSGF_DIALOGBOX = 0
   MSGF_MAINLOOP = 8
   MSGF_MAX = 8
   MSGF_MENU = 2
   MSGF_MESSAGEBOX = 1
   MSGF_MOVE = 3
   MSGF_NEXTWINDOW = 6
   MSGF_SCROLLBAR = 5
   MSGF_SIZE = 4
   MSGF_USER = 4096
End Enum
Public Enum EHTShellCodes
   HSHELL_ACTIVATESHELLWINDOW = 3
   HSHELL_WINDOWCREATED = 1
   HSHELL_WINDOWDESTROYED = 2
   HSHELL_WINDOWACTIVATED = 4
   HSHELL_GETMINRECT = 5
   HSHELL_REDRAW = 6
   HSHELL_TASKMAN = 7
   HSHELL_LANGUAGE = 8
   HSHELL_ACCESSIBILITYSTATE = 11
End Enum
Public Enum EHTMousewParamValues
   WM_MOUSEMOVE = &H200
   WM_LBUTTONDBLCLK = &H203
   WM_LBUTTONDOWN = &H201
   WM_LBUTTONUP = &H202
   WM_MBUTTONDBLCLK = &H209
   WM_MBUTTONDOWN = &H207
   WM_MBUTTONUP = &H208
   WM_RBUTTONDBLCLK = &H206
   WM_RBUTTONUP = &H205
   WM_RBUTTONDOWN = &H204
   WM_NCLBUTTONDBLCLK = &HA3
   WM_NCLBUTTONDOWN = &HA1
   WM_NCLBUTTONUP = &HA2
   WM_NCMBUTTONDBLCLK = &HA9
   WM_NCMBUTTONDOWN = &HA7
   WM_NCMBUTTONUP = &HA8
   WM_NCMOUSEMOVE = &HA0
   WM_NCRBUTTONDBLCLK = &HA6
   WM_NCRBUTTONDOWN = &HA4
   WM_NCRBUTTONUP = &HA5
End Enum
Public Enum EHTHookErrorConstants
   eehHookBase = vbObjectError + 1048
End Enum
Public Enum EHTHookCode
   HC_ACTION = 0
   HC_GETNEXT = 1
   HC_SKIP = 2
   HC_NOREMOVE = 3
   HC_SYSMODALON = 4
   HC_SYSMODALOFF = 5
End Enum

Public Function InstallHook( _
      ByVal IHook As IWindowsHook, _
      ByVal eType As EHTHookTypeConstants _
   ) As Boolean
Attribute InstallHook.VB_Description = "Installs a Hook of the specified type."
   InstallHook = mWindowsHook.InstallHook(IHook, eType)
End Function
Public Sub RemoveHook( _
      ByVal IHook As IWindowsHook, _
      ByVal eType As EHTHookTypeConstants _
   )
Attribute RemoveHook.VB_Description = "Removes a previously installed Hook."
   mWindowsHook.RemoveHook IHook, eType
End Sub
Public Property Get MouselParam( _
      ByVal lParam As Long _
   ) As cMouselParam
Attribute MouselParam.VB_Description = "Converts the lParam member of a WH_MOUSE hook notification into the consituent members."
   If ValidlParamType = WH_MOUSE Then
      Dim c As New cMouselParam
      c.Init lParam
      Set MouselParam = c
   Else
      Debug.Print "ERROR IN MOUSELPARAM"
      pError 1
   End If
End Property
Public Property Get KeyboardlParam( _
      ByVal lParam As Long _
   ) As cKeyboardlParam
Attribute KeyboardlParam.VB_Description = "Converts the lParam member of a WH_KEYBOARD hook notification into the consituent members."
   If ValidlParamType = WH_KEYBOARD Then
      Dim c As New cKeyboardlParam
      c.Init lParam
      Set KeyboardlParam = c
   Else
      pError 1
   End If
End Property
Public Property Get GetMsglParam( _
      ByVal lParam As Long _
   ) As cGetMsglParam
Attribute GetMsglParam.VB_Description = "Converts the lParam member of a WH_GETMSG hook notification into the consituent members."
   If ValidlParamType = WH_GETMESSAGE Then
      Dim c As New cGetMsglParam
      c.Init lParam
      Set GetMsglParam = c
   Else
      pError 1
   End If
End Property
Public Property Get MsgFilterlParam( _
      ByVal lParam As Long _
   ) As cGetMsglParam
   If ValidlParamType = WH_MSGFILTER Then
      Dim c As New cGetMsglParam
      c.Init lParam
      Set MsgFilterlParam = c
   Else
      pError 1
   End If
End Property
Public Property Get CallWndProclParam( _
      ByVal lParam As Long _
   ) As cCallWndProclParam
   If ValidlParamType = WH_CALLWNDPROC Then
      Dim c As New cCallWndProclParam
      c.Init lParam
      Set CallWndProclParam = c
   Else
      pError 1
   End If
End Property
Public Property Get CallWndProcRetlParam( _
      ByVal lParam As Long _
   ) As cCallWndProcRetlParam
Attribute CallWndProcRetlParam.VB_Description = "Converts the lParam member of a WH_CALLWNDPROCRET hook notification into the consituent members."
   If ValidlParamType = WH_CALLWNDPROCRET Then
      Dim c As New cCallWndProcRetlParam
      c.Init lParam
      Set CallWndProcRetlParam = c
   Else
      pError 1
   End If
End Property
Public Property Get JournalPlaybacklParam( _
      ByVal lParam As Long _
   ) As cJournallParam
   If ValidlParamType = WH_JOURNALPLAYBACK Then
      Dim c As New cJournallParam
      c.Init lParam
      Set JournalPlaybacklParam = c
   Else
      pError 1
   End If
End Property
Public Property Get JournalRecordlParam( _
      ByVal lParam As Long _
   ) As cJournallParam
   If ValidlParamType = WH_JOURNALRECORD Then
      Dim c As New cJournallParam
      c.Init lParam
      Set JournalRecordlParam = c
   Else
      Debug.Print "Get JournalRecordlParam"
      pError 1
   End If
End Property
Private Sub pError(ByVal lErr As Long)
   Err.Raise vbObjectError + 1048 + lErr + 3000, App.EXEName & ".vbalWinHook", "Class only valid during HookCall with the same hook type."
End Sub

Private Sub Class_Terminate()
   ' Clear up?
   debugmsg "GHook:terminate"
End Sub
