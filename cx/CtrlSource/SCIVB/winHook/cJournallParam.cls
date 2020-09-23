VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJournallParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EJournalRecordMessage
   WM_CANCELJOURNAL = &H4B
End Enum

Private m_tEMSG As EVENTMSG
Private m_lParam As Long

Public Property Get hWnd() As Long
   hWnd = m_tEMSG.hWndMsg
End Property
Public Property Let hWnd(ByVal value As Long)
   m_tEMSG.hWndMsg = value
   updateValue
End Property
Public Property Get lParamHigh() As Long
   lParamHigh = m_tEMSG.lParamHigh
End Property
Public Property Let lParamHigh(ByVal value As Long)
   m_tEMSG.lParamHigh = value
   updateValue
End Property
Public Property Get lParamLow() As Long
   lParamLow = m_tEMSG.lParamLow
End Property
Public Property Let lParamLow(ByVal value As Long)
   m_tEMSG.lParamLow = value
   updateValue
End Property
Public Property Get MsgTime() As Long
   MsgTime = m_tEMSG.MsgTime
End Property
Public Property Let MsgTime(ByVal value As Long)
   m_tEMSG.MsgTime = value
   updateValue
End Property
Public Property Get Msg() As Long
   Msg = m_tEMSG.wMsg
End Property
Public Property Let Msg(ByVal value As Long)
   m_tEMSG.wMsg = value
   updateValue
End Property

Private Sub updateValue()
   CopyMemory ByVal m_lParam, m_tEMSG, Len(m_tEMSG)
End Sub

Friend Sub Init(ByVal lParam As Long)
   CopyMemory m_tEMSG, ByVal lParam, Len(m_tEMSG)
   m_lParam = lParam
End Sub
