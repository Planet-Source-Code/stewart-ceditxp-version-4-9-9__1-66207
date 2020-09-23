VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGetMsglParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Translates the lParam of a WH_GETMSG Hook"
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
' Decodes the lParam of a WH_GETMSG hook.
'
' ======================================================================================

Private m_lParam As Long
Private m_tMsg As Msg

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Gets/Sets the window handle this message is being retrieved by."
   hWnd = m_tMsg.hWnd
End Property
Public Property Let hWnd(ByVal lhWNd As Long)
   If m_lParam <> 0 And ValidlParamType = WH_MSGFILTER Then
      m_tMsg.hWnd = lhWNd
      CopyMemory ByVal m_lParam, m_tMsg, Len(m_tMsg)
   Else
      Debug.Print "Let hWnd"
      pError 1
   End If
End Property
Public Property Get lParam() As Long
Attribute lParam.VB_Description = "Gets/Sets the lParam member of the message."
   lParam = m_tMsg.lParam
End Property
Public Property Let lParam(ByVal lParam As Long)
   If m_lParam <> 0 And ValidlParamType = WH_MSGFILTER Then
      m_tMsg.lParam = lParam
      CopyMemory ByVal m_lParam, m_tMsg, Len(m_tMsg)
   Else
      Debug.Print "Let lParam"
      pError 1
   End If
End Property
Public Property Get wParam() As Long
Attribute wParam.VB_Description = "Gets/sets the wParam value for this message."
   wParam = m_tMsg.wParam
End Property
Public Property Let wParam(ByVal wParam As Long)
   If m_lParam <> 0 And ValidlParamType = WH_MSGFILTER Then
      m_tMsg.wParam = wParam
      CopyMemory ByVal m_lParam, m_tMsg, Len(m_tMsg)
   Else
      Debug.Print "Let wParam"
      pError 1
   End If
End Property
Public Property Get Message() As Long
Attribute Message.VB_Description = "Gets/sets the message."
   Message = m_tMsg.Message
End Property
Public Property Let Message(ByVal iMsg As Long)
   If m_lParam <> 0 And ValidlParamType = WH_MSGFILTER Then
      m_tMsg.Message = iMsg
      CopyMemory ByVal m_lParam, m_tMsg, Len(m_tMsg)
   Else
      Debug.Print "Let Message"
      pError 1
   End If
End Property
Public Property Get X() As Long
Attribute X.VB_Description = "Gets the x mouse position relative to the screen in pixels at the point when this message was posted."
   X = m_tMsg.pt.X
End Property
Public Property Get Y() As Long
Attribute Y.VB_Description = "Gets the y mouse position relative to the screen in pixels at the point when this message was posted."
   Y = m_tMsg.pt.Y
End Property
Private Sub pError(ByVal lErrNum As Long)
   Err.Raise vbObjectError + 1048 + lErrNum + 3000, App.EXEName, "Invalid object use."
End Sub
Friend Sub Init(ByVal lParam As Long)
   m_lParam = lParam
   CopyMemory m_tMsg, ByVal lParam, Len(m_tMsg)
End Sub
