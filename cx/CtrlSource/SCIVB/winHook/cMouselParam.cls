VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMouselParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Translates the lParam of a WH_MOUSE Hook"
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
' Decodes the lParam of a WH_MOUSE hook.
'
' ======================================================================================

Private m_tMHS As MOUSEHOOKSTRUCT
Private m_lParam As Long

Public Property Get X() As Long
Attribute X.VB_Description = "X position of the mouse relative to the screen, in pixels."
   X = m_tMHS.pt.X
End Property
Public Property Get Y() As Long
Attribute Y.VB_Description = "Y position of the mouse relative to the screen, in pixels."
   Y = m_tMHS.pt.Y
End Property
Public Property Get ClientX(ByVal lhWNd As Long) As Long
Attribute ClientX.VB_Description = "X position of the mouse relative to the client window, in pixels."
Dim tP As POINTAPI
   LSet tP = m_tMHS.pt
   ScreenToClient lhWNd, tP
   ClientX = tP.X
End Property
Public Property Get ClientY(ByVal lhWNd As Long) As Long
Attribute ClientY.VB_Description = "Y position of the mouse relative to the client window, in pixels."
Dim tP As POINTAPI
   LSet tP = m_tMHS.pt
   ScreenToClient lhWNd, tP
   ClientY = tP.Y
End Property
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Window handle the mouse message will be sent to."
   hWnd = m_tMHS.hWnd
End Property
Public Property Get wHitTestCode() As Long
Attribute wHitTestCode.VB_Description = "Windows Hit Test Code for the mouse message."
   wHitTestCode = m_tMHS.wHitTestCode
End Property
Public Property Get dwExtraInfo() As Long
Attribute dwExtraInfo.VB_Description = "Extra information associated with the mouse message."
   dwExtraInfo = m_tMHS.dwExtraInfo
End Property
Friend Sub Init(ByVal lParam As Long)
   CopyMemory m_tMHS, ByVal lParam, Len(m_tMHS)
   m_lParam = lParam
End Sub

