VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTabCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ======================================================================================
' Name:     cTabCollection.cls
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
' Strongly-typed collection of tab objects for vbalDTab control
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================


Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private m_lOwner As Long
Private m_hWnd As Long

Friend Sub Init(ByVal lPtr As Long, ByVal lhWnd As Long)
   m_lOwner = lPtr
   m_hWnd = lhWnd
End Sub

Private Property Get TabObject() As vbalDTabControl
   If Not (m_hWnd = 0) Then
      If (IsWindow(m_hWnd)) Then
         If Not (m_lOwner = 0) Then
            Dim oT As Object
            CopyMemory oT, m_lOwner, 4
            Set TabObject = oT
            CopyMemory oT, 0&, 4
         End If
      End If
   End If
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of tabs."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Count = oT.fTabCount
   End If
End Property

Public Property Get Item(Key As Variant) As cTab
Attribute Item.VB_Description = "Gets the tab with the specified Key (or at the specified position)."
Attribute Item.VB_MemberFlags = "200"
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Set Item = oT.fItem(Key)
   End If
End Property

Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Removes the specified tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      oT.fRemove Key
   End If
End Sub

Public Function Add( _
      Optional Key As Variant, _
      Optional KeyBefore As Variant, _
      Optional Caption As String, _
      Optional IconIndex As Long _
   ) As cTab
Attribute Add.VB_Description = "Adds or inserts a new tab."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Set Add = oT.fAdd(Key, KeyBefore, Caption, IconIndex)
   End If
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all tabs from the control."
Dim oT As vbalDTabControl
   Set oT = TabObject()
   If Not oT Is Nothing Then
      Dim i As Long
      For i = oT.fTabCount To 1 Step -1
         oT.fRemove i
      Next i
   End If
End Sub


