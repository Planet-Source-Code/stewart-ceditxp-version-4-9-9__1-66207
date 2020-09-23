VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Base error number constant
Private Const ERRBASE = vbObjectError + 1740

'////////////////////////////////////////////////////////////////////
'// Private/Public Win32 API Declarations
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'////////////////////////////////////////////////////////////////////
'// Private/Public Event Declarations
Public Event Timer()

'////////////////////////////////////////////////////////////////////
'// Private/Public Variable Declarations
Private m_lTimerID          As Long     ' Timer ID
Private m_lInterval         As Long     ' Timer interval
Private m_bEnabled          As Boolean  ' Timer enabled

'********************************************************************
'* Name: Interval
'* Description: Return/sets timer interval.
'********************************************************************
Public Property Get Interval() As Long
    On Error GoTo PROC_ERR_Interval

100     Interval = m_lInterval

PROC_EXIT:
        Exit Property
PROC_ERR_Interval:
        Err.Raise ERRBASE, "CTimer.Interval", "CTimer component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property
Public Property Let Interval(ByVal lValue As Long)
    On Error GoTo PROC_ERR_Interval

100     If m_lInterval = lValue Then Exit Property
    
102     If lValue > 0 Then
            ' First destroy timer
104         Enabled = False
            ' Then set new interval
106         m_lInterval = lValue
            ' Create new timer
108         Enabled = True
        Else
110         Enabled = False
        End If

PROC_EXIT:
        Exit Property
PROC_ERR_Interval:
        Err.Raise ERRBASE, "CTimer.Interval", "CTimer component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property

'********************************************************************
'* Name: Enabled
'* Description: Return/sets timer enabled state.
'********************************************************************
Public Property Get Enabled() As Boolean
    On Error GoTo PROC_ERR_Enabled

100     Enabled = m_bEnabled

PROC_EXIT:
        Exit Property
PROC_ERR_Enabled:
        Err.Raise ERRBASE, "CTimer.Enabled", "CTimer component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property
Public Property Let Enabled(ByVal bValue As Boolean)
    On Error GoTo PROC_ERR_Enabled

100     If m_bEnabled = bValue Then Exit Property
    
102     m_bEnabled = bValue
    
        ' If it is enabled
104     If m_bEnabled Then
            ' Create timer
106         m_lTimerID = SetTimer(0, 0, m_lInterval, AddressOf TimerProc)
            ' Add class reference
108         AddTimer Me, m_lTimerID
                    
        ' If it is disabled
        Else
            ' Destroy timer
110         KillTimer 0, m_lTimerID
            ' Remove class reference
112         RemoveTimer m_lTimerID
        End If

PROC_EXIT:
        Exit Property
PROC_ERR_Enabled:
        Err.Raise ERRBASE, "CTimer.Enabled", "CTimer component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property

'********************************************************************
'* Name: RaiseTimerEvent
'* Description: Raise timer event.
'********************************************************************
Public Sub RaiseTimerEvent()
    On Error Resume Next
    
    RaiseEvent Timer
End Sub

'********************************************************************
'* Name: Class_Initialize
'* Description: Class initialization routine.
'********************************************************************
Private Sub Class_Initialize()
    On Error Resume Next
    
    ' Set default values
    m_lInterval = 1000
End Sub

'********************************************************************
'* Name: Class_Terminate
'* Description: Class termination routine.
'********************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    
    ' Destroy timer
    Enabled = False
End Sub

