VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCallWndProclParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Translates the lParam of a WH_CALLWNDPROC Hook."
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
' Decodes the lParam of a WH_CALLWNDPROC hook.
'
' ======================================================================================


Private m_tCWP As CWPSTRUCT
Private m_lParam As Long

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Gets the window handle that this message is being sent to."
   hWnd = m_tCWP.hWnd
End Property
Public Property Get lParam() As Long
Attribute lParam.VB_Description = "Gets the lParam for this message."
   lParam = m_tCWP.lParam
End Property
Public Property Get wParam() As Long
Attribute wParam.VB_Description = "Gets the wParam for this message."
   wParam = m_tCWP.wParam
End Property
Public Property Get Message() As Long
Attribute Message.VB_Description = "Gets the message."
   Message = m_tCWP.Message
End Property

Friend Sub Init(ByVal lParam As Long)
   m_lParam = lParam
   CopyMemory m_tCWP, ByVal lParam, Len(m_tCWP)
End Sub
