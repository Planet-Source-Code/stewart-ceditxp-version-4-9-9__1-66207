VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArielTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'Module     : ArielTimer
'Description: Simple Timer delay class with events
'Version    : V1.00 Sep 2000
'Release    : VB5/VB6
'Copyright  : © T De Lange, Ariel Systems, 2000
'e-mail     : tomdl@attglobal.net
'----------------------------------------------------------------------------------
'V1.00    Sep 00 Original version, based on code published by Riaan Aspeling
'         http://www.planet-source-code.com
'----------------------------------------------------------------------------------
'Implement the following in your project form:
'Declarations:
'Private WithEvents ArTimer As ArielTimer  'Must be private to raise events
'Form_Load():
'Set TimerObj = New ArielTimer
'TimerObj.Interval = <any long, in ms>
'To start the timer, issue the following:
'TimerObj.Enabled=True
'To stop the timer, issue the following:
'TimerObj.Enabled=False
'----------------------------------------------------------------------------------
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

Private mEnabled As Boolean     'Timer running?
Private mInterval As Long       'Timer interval, ms
Private mStart As Long          'Start timer mark
Private mName As String         'Timer Name
Public Event OnTimer()

'----------------------------------------------
'Windows API Declarations for ArielTimerClass
'----------------------------------------------
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Property Get Enabled() As Boolean
'----------------------------------------
'Read current Enabled status
'----------------------------------------
Enabled = mEnabled
'Call Running

End Property


Public Property Let Enabled(ByVal vEnabled As Boolean)
'-----------------------------------------------------
'Reset and Enable timer
'-----------------------------------------------------
mEnabled = vEnabled
mStart = 0
Call Running

End Property


Public Property Let Interval(vInterval As Long)
'-----------------------------------------------
'Set new interval & reset timer
'-----------------------------------------------
If vInterval > 0 Then
  mStart = 0
  mInterval = vInterval
  Call Running
Else
  mInterval = 0
  mStart = 0
  mEnabled = False
End If
  
End Property

Public Property Get Interval() As Long
'----------------------------------------
'Read current Interval
'----------------------------------------
Interval = mInterval

End Property

Public Property Let Name(vName As String)
'----------------------------------------
'Set current Name
'----------------------------------------
If vName <> "" Then
  mName = vName
End If

End Property

Public Property Get Name() As String
'----------------------------------------
'Read current Name
'----------------------------------------
Name = mName

End Property
Private Sub Running()
'-------------------------------------
'Internal timer loop raising event
'when required
'-------------------------------------
Dim Elapsed As Long

Do While mEnabled
  'Internal reset
  If mStart = 0 Then
    mStart = GetTickCount
  End If
  Elapsed = GetTickCount
  If (Elapsed - mStart) >= mInterval Then
    'Reset start
    mStart = GetTickCount
    'Raise event
    RaiseEvent OnTimer
  End If 'calls wont "spoof" on you!
  DoEvents
Loop

End Sub


Private Sub Class_Initialize()
'--------------------------------
'Initialise Timer
'--------------------------------
mEnabled = False
mInterval = 400
mName = "HoverTimer"

End Sub


Private Sub Class_Terminate()
'------------------------------------
'Terminate timer
'------------------------------------
mEnabled = False

End Sub

