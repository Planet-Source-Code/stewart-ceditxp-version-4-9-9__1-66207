VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKeyboardlParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Translates the lParam of a WH_KEYBOARD Hook"
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
' Decodes the lParam of a WH_KEYBOARD hook.
'
' ======================================================================================


Private m_bKeyUp As Boolean
Private m_bAlt As Boolean

Public Property Get KeyUp() As Boolean
Attribute KeyUp.VB_Description = "Whether the Key is up for this keyboard message."
   KeyUp = m_bKeyUp
End Property
Public Property Get KeyDown() As Boolean
Attribute KeyDown.VB_Description = "Whether the Key is down for this keyboard message."
   KeyDown = Not (m_bKeyUp)
End Property
Public Property Get Alt() As Boolean
Attribute Alt.VB_Description = "Whether the Alt key is pressed for this keyboard message."
   Alt = m_bAlt
End Property

Friend Sub Init(ByVal lParam As Long)
   ' Key up or down:
   m_bKeyUp = ((lParam And &H80000000) = &H80000000)
   ' Alt pressed?
   m_bAlt = ((lParam And &H20000000) = &H20000000)
End Sub
