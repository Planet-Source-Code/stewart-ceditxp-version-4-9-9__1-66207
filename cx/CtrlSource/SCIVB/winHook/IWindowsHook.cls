VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWindowsHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Client implementation required to interface with the vbAccelerator Hook Library."
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
' A user of the Windows Hook Library must implement
' this interface in order to receive hook calls.
'
' ======================================================================================


Public Function HookProc( _
      ByVal eType As EHTHookTypeConstants, _
      ByVal nCode As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long, _
      ByRef bConsume As Boolean _
   ) As Long
Attribute HookProc.VB_Description = "Fired whenever a Hook receives a notification.  Set bConsume to True if you wish to stop the notification from reaching your app."
   
End Function

