VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ======================================================================================
' Name:     cTab.cls
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     7 January 2003
'
' Requires: -
'
' Copyright © 2003 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Public access to tab properties for the vbalDTab6 control.
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private m_lOwner As Long
Private m_hWnd As Long
Private m_lId As Long

Friend Sub fInit(ByVal lPtr As Long, ByVal lhWnd As Long, ByVal lId As Long)
   m_lOwner = lPtr
   m_hWnd = lhWnd
   m_lId = lId
End Sub

Private Property Get TabObject() As vbalDTabControl
   If Not (m_hWnd = 0) Then
      If (IsWindow(m_hWnd)) Then
         If Not (m_lOwner = 0) Then
            Dim oT As Object
            CopyMemory oT, m_lOwner, 4
            Set TabObject = oT
            CopyMemory oT, 0&, 4
            Exit Property
         End If
      End If
   End If
   Err.Raise 9, App.EXEName & ".vbalDTabControl"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Gets/sets the caption of this tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Caption = oT.fTabCaption(m_lId)
   End If
End Property
Public Property Let Caption(ByVal sCaption As String)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabCaption(m_lId) = sCaption
   End If
End Property
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Gets/sets the ToolTipText to be shown for this tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      ToolTipText = oT.fTabToolTipText(m_lId)
   End If
End Property
Public Property Let ToolTipText(ByVal sToolTipText As String)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabToolTipText(m_lId) = sToolTipText
   End If
End Property
Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Gets/sets the 0-based index of an icon to be shown on this tab.  Set to -1 if no icon is needed."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      IconIndex = oT.fTabIconIndex(m_lId)
   End If
End Property
Public Property Let IconIndex(ByVal lIconIndex As Long)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabIconIndex(m_lId) = lIconIndex
   End If
End Property
Public Property Get Key() As String
Attribute Key.VB_Description = "Gets the key of this tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Key = oT.fTabKey(m_lId)
   End If
End Property
Public Property Get ItemData() As Long
Attribute ItemData.VB_Description = "Gets/sets a long value associated with the tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      ItemData = oT.fTabItemData(m_lId)
   End If
End Property
Public Property Let ItemData(ByVal lItemData As Long)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabItemData(m_lId) = lItemData
   End If
End Property
Public Property Get Tag() As String
Attribute Tag.VB_Description = "Gets/sets a string tag associated with this tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Tag = oT.fTabTag(m_lId)
   End If
End Property
Public Property Let Tag(ByVal sTag As String)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabTag(m_lId) = sTag
   End If
End Property
Public Property Get CanClose() As Boolean
Attribute CanClose.VB_Description = "Gets/sets whether the close button (if shown) will be enabled for this tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      CanClose = oT.fTabCanClose(m_lId)
   End If
End Property
Public Property Let CanClose(ByVal bCanClose As Boolean)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabCanClose(m_lId) = bCanClose
   End If
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gets/sets whether this tab is enabled."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Enabled = oT.fTabEnabled(m_lId)
   End If
End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabEnabled(m_lId) = bEnabled
   End If
End Property
Public Property Let Panel(ByRef ctlThis As Object)
Attribute Panel.VB_Description = "Gets/sets any control you would like to be displayed in this tab."
   pSetPanel ctlThis
End Property
Public Property Set Panel(ByRef ctlThis As Object)
   pSetPanel ctlThis
End Property
Private Sub pSetPanel(ByRef ctlThis As Object)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabPanel(m_lId) = ctlThis
   End If
End Sub
Public Property Get Panel() As Object
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Set Panel = oT.fTabPanel(m_lId)
   End If
End Property
Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Gets/sets whether this tab is selected."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Selected = oT.fTabSelected(m_lId)
   End If
End Property
Public Property Let Selected(ByVal bState As Boolean)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabSelected(m_lId) = bState
   End If
End Property
Public Property Get Index() As Long
Attribute Index.VB_Description = "Gets/sets the index of this tab within the control."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Index = oT.fTabIndex(m_lId)
   End If
End Property
Public Property Let Index(ByVal lIndex As Long)
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fTabIndex(m_lId) = lIndex
   End If
End Property
