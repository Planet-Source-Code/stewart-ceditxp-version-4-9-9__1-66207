Attribute VB_Name = "modSystem"
Option Explicit

'////////////////////////////////////////////////////////////////////
'// Private/Public Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'////////////////////////////////////////////////////////////////////
'// Private/Public Variable Declarations
Private m_oTimers       As New Collection   ' Timers collection

'********************************************************************
'* Name: pEnumChildWindowProc
'* Description: Callback routine for enumerating MDI child windows.
'********************************************************************
Public Function pEnumChildWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sBuf As String
    Dim sClass As String
    Dim iPos As Long
   
    If Not lParam = 0 Then
        sBuf = String$(261, 0)
        GetClassName hWnd, sBuf, 260
        iPos = InStr(sBuf, vbNullChar)
        If iPos > 1 Then
            sClass = Left$(sBuf, iPos - 1)
            If InStr(sClass, "Form") > 0 Then
                Dim ctlTab As RevMDITabsCtl
                Dim oT As Object
                CopyMemory oT, lParam, 4
                Set ctlTab = oT
                CopyMemory oT, 0&, 4
                ctlTab.fAddMDIChildWindow hWnd
            End If
        End If
        pEnumChildWindowProc = 1
    End If
End Function

'********************************************************************
'* Name: TimerProc
'* Description: Timer callback method.
'********************************************************************
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    On Error Resume Next
    
    Dim oTimer As CTimer

    If hWnd = 0 Then
        ' Get timer object
        Set oTimer = m_oTimers.Item(CStr(idEvent))
        ' Raise timer event
        If Err.Number = 0 Then oTimer.RaiseTimerEvent
    End If
    
    Set oTimer = Nothing
End Sub

'********************************************************************
'* Name: AddTimer
'* Description: Add specified CTimer class into class collection.
'********************************************************************
Public Sub AddTimer(ByRef oTimer As CTimer, ByVal lTimerID As Long)
    On Error Resume Next
    
    m_oTimers.Add oTimer, CStr(lTimerID)
End Sub

'********************************************************************
'* Name: RemoveTimer
'* Description: Remove specified CTimer class from class collection.
'********************************************************************
Public Sub RemoveTimer(ByVal lTimerID As Long)
    On Error Resume Next
    
    m_oTimers.Remove CStr(lTimerID)
End Sub
