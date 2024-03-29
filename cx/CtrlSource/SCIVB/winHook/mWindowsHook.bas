Attribute VB_Name = "mWindowsHook"
Option Explicit

' ======================================================================================
' Name:     vbAccelerator Windows Hook Library
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     25 August 1999
'
' Requires: IWindowsHook.Cls
'
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Windows Hooks are the most powerful technique available to VB
' programmers for modifying the behaviour of windows.  The
' Windows Hook Library makes using a Hook a little easier to
' control by wrapping up some of the implementation details
' and allowing multiple consumers to share the same windows
' hook.
'
' This module plus the implementation class IWindowsHook.cls
' provide the minimal implementation of the Hook library.  Extract
' these two for direct compilation into your own projects.
'
' ======================================================================================


' ===========================================================================
' API Calls:
' ===========================================================================
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Public Type POINTAPI
   X As Long
   Y As Long
End Type
Public Type Msg '{     /* msg */
   hWnd As Long     '\\ The window whose Winproc will receive the message
   Message As Long  '\\ The message number
   wParam As Long
   lParam As Long
   time As Long     '\\ The time the message was posted
   pt As POINTAPI   '\\ The cursor position in screen coordinates
                    '\\  of the message
End Type
Public Type MOUSEHOOKSTRUCT '{ // ms
    pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
Public Type CWPSTRUCT
   lParam As Long
   wParam As Long
   Message As Long
   hWnd As Long
End Type
Public Type CWPRETSTRUCT
    lResult As Long
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type
Public Type EVENTMSG
   wMsg As Long
   lParamLow As Long
   lParamHigh As Long
   MsgTime As Long
   hWndMsg As Long
End Type

Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' ===========================================================================
' Implementation
' ===========================================================================
' Hook handles:
Private m_hHook([_WH_MIN] To [_WH_MAX]) As Long
' Hook consumers:
Private Type tHookConsumer
   lPtr As Long                     ' Pointer to consumer object
   eType As EHTHookTypeConstants    ' Type of hook
End Type
Private m_tHookConsumer() As tHookConsumer
Private m_iConsumerCount As Long
Private m_eValidItem As EHTHookTypeConstants
#Const debugmsg = 0

Public Sub debugmsg(ByVal sMsg As String)
   #If debugmsg = 1 Then
      MsgBox sMsg, vbInformation
   #Else
      Debug.Print sMsg
   #End If
End Sub

Public Property Get ValidlParamType() As EHTHookTypeConstants
   ValidlParamType = m_eValidItem
End Property

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must
End Property


Public Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
   
   ' Return the error message associated with LastDLLError:
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount <> 0 Then
      WinAPIError = Left$(sBuff, lCount)
   End If

End Function

Public Function InstallHook( _
      ByRef IHook As IWindowsHook, _
      ByVal eType As EHTHookTypeConstants _
   ) As Boolean
Dim hHook As Long
Dim lpFn As Long
Dim lErr As Long
Dim lPtr As Long
Dim i As Long
Dim bExists As Boolean
Dim iAvailSlot As Long
Dim hThread As Long
Dim hInstance As Long
      
   ' If Hook not already installed:
   If (m_hHook(eType) = 0) Then
      hThread = GetCurrentThreadId()
      hInstance = 0
      Select Case eType
      Case WH_CALLWNDPROC
         lpFn = HookAddress(AddressOf CallWndProc)
      Case WH_CALLWNDPROCRET
         lpFn = HookAddress(AddressOf CallWndProcRet)
      Case WH_MSGFILTER
         lpFn = HookAddress(AddressOf MessageProc)
      Case WH_MOUSE
         lpFn = HookAddress(AddressOf MouseProc)
      Case WH_KEYBOARD
         lpFn = HookAddress(AddressOf KeyboardProc)
      Case WH_GETMESSAGE
         lpFn = HookAddress(AddressOf GetMsgProc)
      Case WH_FOREGROUNDIDLE
         lpFn = HookAddress(AddressOf ForegroundIdleProc)
      Case WH_SHELL
         lpFn = HookAddress(AddressOf ShellProc)
      Case WH_CBT
         lpFn = HookAddress(AddressOf CBTProc)
      Case WH_JOURNALRECORD
         lpFn = HookAddress(AddressOf JournalRecordProc)
         hThread = 0
         hInstance = App.hInstance
      Case Else
         Err.Raise eehHookBase + 1, App.EXEName & ".cVBALHook", "Unsupported Hook Type."
      End Select
      ' Add the hook:
      If lpFn <> 0 Then
         hHook = SetWindowsHookEx(eType, lpFn, hInstance, hThread)
         ' If we succeeded then set up the hook type:
         If (hHook <> 0) Then
            ' Succeeded; store the handle so we can restore it
            ' again later:
            m_hHook(eType) = hHook
         Else
            ' Failed:
            lErr = Err.LastDllError
            Err.Raise vbObjectError + 1049, App.EXEName & ".mHook", WinAPIError(lErr)
         End If
      End If
   End If

   ' If have a hook function:
   If (m_hHook(eType) <> 0) Then
      ' Add the class to the hook receive list:
      lPtr = ObjPtr(IHook)
      For i = 1 To m_iConsumerCount
         With m_tHookConsumer(i)
            If .eType = eType And .lPtr = lPtr Then
               bExists = True
            ElseIf .lPtr = 0 And iAvailSlot = 0 Then
               iAvailSlot = i
            End If
         End With
      Next i
      If Not (bExists) Then
         If (iAvailSlot = 0) Then
            m_iConsumerCount = m_iConsumerCount + 1
            ReDim Preserve m_tHookConsumer(1 To m_iConsumerCount) As tHookConsumer
            iAvailSlot = m_iConsumerCount
         End If
         With m_tHookConsumer(iAvailSlot)
            .lPtr = lPtr
            .eType = eType
         End With
      End If
      ' Success:
      debugmsg "mWindowsHook: Number of attached: " & m_iConsumerCount
      InstallHook = True
   End If

End Function

Private Function HookAddress(ByVal lPtr As Long) As Long
   ' Work around for VB's poor AddressOf implementation:
   HookAddress = lPtr
End Function
Private Function CBTProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If nCode >= 0 Then
      ' Notification only:
      HookCall WH_CBT, nCode, wParam, lParam
   End If
   CBTProc = CallNextHookEx(m_hHook(WH_CBT), nCode, wParam, lParam)
End Function
Private Function ShellProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook isn't really much use when it only applies to
   ' the current thread, to be honest.
   If nCode >= 0 Then
      ' Notification only:
      HookCall WH_SHELL, nCode, wParam, lParam
   End If
   ShellProc = CallNextHookEx(m_hHook(WH_SHELL), nCode, wParam, lParam)
End Function
Private Function JournalRecordProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If nCode >= 0 Then
      HookCall WH_JOURNALRECORD, nCode, wParam, lParam
   End If
   JournalRecordProc = CallNextHookEx(m_hHook(WH_JOURNALRECORD), nCode, wParam, lParam)
End Function
Private Function ForegroundIdleProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook isn't particularly useful either; it continuously jabbers
   ' away saying that the foreground is idle almost all the time...
   If nCode >= 0 Then
      ' Notification only:
      HookCall WH_FOREGROUNDIDLE, nCode, wParam, lParam
   End If
   ForegroundIdleProc = CallNextHookEx(m_hHook(WH_FOREGROUNDIDLE), nCode, wParam, lParam)
End Function
Private Function MessageProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook allows you to intercept every message sent to every window
   ' in your application
   If nCode >= 0 Then
      If HookCall(WH_MSGFILTER, nCode, wParam, lParam) = 1 Then
         MessageProc = 0
         Exit Function
      End If
   End If
   MessageProc = CallNextHookEx(m_hHook(WH_MSGFILTER), nCode, wParam, lParam)
End Function
Private Function CallWndProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook is called just before the WindowProc is called for
   ' every window in your application.  The overhead of using this
   ' hook is very high, so only use it for short periods if possible.
   If nCode >= 0 Then
      ' Can discard the message.
      If HookCall(WH_CALLWNDPROC, nCode, wParam, lParam) = 1 Then
         ' not recommended though...
         CallWndProc = 0
         Exit Function
      End If
   End If
   CallWndProc = CallNextHookEx(m_hHook(WH_CALLWNDPROC), nCode, wParam, lParam)
End Function
Private Function CallWndProcRet(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' Same as CallWndProc, but it is called just before the
   ' WindowProc for every window in your application is about
   ' to be returned.  Again, overhead is very high for this hook.
   If nCode >= 0 Then
      ' notification:
      HookCall WH_CALLWNDPROCRET, nCode, wParam, lParam
   End If
   CallWndProcRet = CallNextHookEx(m_hHook(WH_CALLWNDPROC), nCode, wParam, lParam)
End Function

Private Function GetMsgProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook is fired whenever any window in your application
   ' is about to call PeekMessage or GetMessage.
   If (nCode >= 0) Then
      ' Can't discard the message, but you can modify
      ' the values:
      HookCall WH_GETMESSAGE, nCode, wParam, lParam
   End If
   GetMsgProc = CallNextHookEx(m_hHook(WH_GETMESSAGE), nCode, wParam, lParam)
End Function
Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook is called just before any mouse message is
   ' going to be posted to a window in your application:
   If (nCode >= 0) Then
      ' Can discard mouse events
      If (HookCall(WH_MOUSE, nCode, wParam, lParam) = 1) Then
         ' Not recommended; but you do it
         MouseProc = 1
         Exit Function
      End If
   End If
   MouseProc = CallNextHookEx(m_hHook(WH_MOUSE), nCode, wParam, lParam)
End Function
Private Function KeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This hook is called just before any WM_KEYDOWN or WM_KEYUP is
   ' going to be posted to a window in your application:
   If (nCode >= 0) Then
      ' Can discard keyboard events:
      If (HookCall(WH_KEYBOARD, nCode, wParam, lParam) = 1) Then
         ' Not recommended; but you do it
         KeyboardProc = 1
         Exit Function
      End If
   End If
   KeyboardProc = CallNextHookEx(m_hHook(WH_KEYBOARD), nCode, wParam, lParam)
End Function
Private Function HookCall(ByVal eType As EHTHookTypeConstants, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim oItem As IWindowsHook
Dim i As Long
Dim bConsume As Boolean
   ' Call the HookProc for any consumers attached to the DLL:
   For i = 1 To m_iConsumerCount
      If Not (m_tHookConsumer(i).lPtr = 0) And (m_tHookConsumer(i).eType = eType) Then
         Set oItem = ObjectFromPtr(m_tHookConsumer(i).lPtr)
         m_eValidItem = eType
         oItem.HookProc eType, nCode, wParam, lParam, bConsume
         m_eValidItem = 0
         If (bConsume) Then
            ' Note: consuming is not recommended unless you really
            ' have to
            HookCall = 1
            Exit Function
         End If
      End If
   Next i
   HookCall = 0
End Function

Public Function RemoveHook( _
      ByVal IHook As IWindowsHook, _
      ByVal eType As EHTHookTypeConstants _
   )
Dim i As Long
Dim lPtr As Long
Dim iRefCount As Long

   ' Remove the hook from the hook list:
   lPtr = ObjPtr(IHook)
   For i = 1 To m_iConsumerCount
      With m_tHookConsumer(i)
         If (.eType = eType) Then
            If (.lPtr = lPtr) Then
               .lPtr = 0
               .eType = -2
            ElseIf (.lPtr <> 0) Then
               iRefCount = iRefCount + 1
            End If
         End If
      End With
   Next i
   
   ' If no more clients on this hook then remove the hook:
   If (iRefCount = 0) Then
      If (m_hHook(eType) <> 0) Then
         UnhookWindowsHookEx m_hHook(eType)
         m_hHook(eType) = 0
      End If
   End If
   
End Function
