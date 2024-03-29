VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Type BITMAP '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Const WM_ACTIVATE = &H6
Private Const WM_SETCURSOR = &H20
Private Const WM_NCHITTEST = &H84&
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_SIZING = &H214

Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEWE = 32644&

Private Const R2_NOTXORPEN = 10  '  DPxn

Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)

Private Enum EHitTestAreas
   HTERROR = (-2)
   HTTRANSPARENT = (-1)
   HTNOWHERE = 0
   HTCLIENT = 1
   HTCAPTION = 2
   HTSYSMENU = 3
   HTGROWBOX = 4
   HTMENU = 5
   HTHSCROLL = 6
   HTVSCROLL = 7
   HTMINBUTTON = 8
   HTMAXBUTTON = 9
   HTLEFT = 10
   HTRIGHT = 11
   HTTOP = 12
   HTTOPLEFT = 13
   HTBOTTOM = 15
   HTBOTTOMLEFT = 16
   HTBOTTOMRIGHT = 17
   HTBORDER = 18
End Enum
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadCursorLong Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private m_hWnd As Long
Private m_hWndParent As Long
Private m_eAlign As AlignConstants
Private m_lSize As Long
Private m_lMinSize As Long
Private m_lMaxSize As Long
Private m_bFullDrag As Boolean
Private m_hCur As Long
Private m_tSplitR As RECT
Private m_tOffset As POINTAPI
Private m_lPattern(0 To 3) As Long
Private m_hBrush As Long
Private m_bInDrag As Boolean

Implements ISubclass

Public Property Get FullDrag() As Boolean
   FullDrag = m_bFullDrag
End Property
Public Property Let FullDrag(ByVal bState As Boolean)
   If Not (bState = m_bFullDrag) Then
      If Not (m_hWnd = 0) Then
         Dim hwnd As Long
         Dim hWndParent As Long
         hwnd = m_hWnd
         hWndParent = m_hWndParent
         Detach
      End If
      m_bFullDrag = bState
      If Not (hwnd = 0) Then
         pAttach hwnd, hWndParent
         m_hWnd = hwnd
         m_hWndParent = hWndParent
      End If
   End If
End Property
Public Property Get SplitSize() As Long
   SplitSize = m_lSize
End Property
Public Property Let SplitSize(ByVal lSize As Long)
   m_lSize = lSize
End Property
Public Property Get MinSize() As Long
   MinSize = m_lMinSize
End Property
Public Property Let MinSize(ByVal lMinSize As Long)
   m_lMinSize = lMinSize
End Property
Public Property Get MaxSize() As Long
   MaxSize = m_lMaxSize
End Property
Public Property Let MaxSize(ByVal lMaxSize As Long)
   m_lMaxSize = lMaxSize
End Property

Public Sub Attach(picThis As Object)
   Detach
   m_hWnd = picThis.hwnd
   m_hWndParent = picThis.Parent.hwnd
   m_eAlign = picThis.Align
   pAttach m_hWnd, m_hWndParent
End Sub
Private Sub pAttach(ByVal hwnd As Long, ByVal hWndParent As Long)
   If (m_bFullDrag) Then
      AttachMessage Me, hwnd, WM_NCHITTEST
      AttachMessage Me, hwnd, WM_SIZING
   Else
      AttachMessage Me, hwnd, WM_SETCURSOR
      AttachMessage Me, hwnd, WM_LBUTTONDOWN
      AttachMessage Me, hwnd, WM_MOUSEMOVE
      AttachMessage Me, hwnd, WM_LBUTTONUP
      AttachMessage Me, hWndParent, WM_ACTIVATE
      Select Case m_eAlign
      Case vbAlignTop, vbAlignBottom
         m_hCur = LoadCursorLong(0, IDC_SIZENS)
      Case Else
         m_hCur = LoadCursorLong(0, IDC_SIZEWE)
      End Select
      createBrush
   End If
End Sub
Public Sub Detach()
   If Not (m_hWnd = 0) Then
      If (m_bFullDrag) Then
         DetachMessage Me, m_hWnd, WM_NCHITTEST
         DetachMessage Me, m_hWnd, WM_SIZING
      Else
         DetachMessage Me, m_hWnd, WM_SETCURSOR
         DetachMessage Me, m_hWnd, WM_LBUTTONDOWN
         DetachMessage Me, m_hWnd, WM_MOUSEMOVE
         DetachMessage Me, m_hWnd, WM_LBUTTONUP
         DetachMessage Me, m_hWndParent, WM_ACTIVATE
      End If
      m_hWnd = 0
      m_hWndParent = 0
      If Not (m_hCur = 0) Then
         DestroyCursor m_hCur
      End If
      DestroyBrush
   End If
End Sub

Private Function createBrush() As Boolean
Dim tbm As BITMAP
Dim hBm As Long

   DestroyBrush
      
   ' Create a monochrome bitmap containing the desired pattern:
   tbm.bmType = 0
   tbm.bmWidth = 16
   tbm.bmHeight = 8
   tbm.bmWidthBytes = 2
   tbm.bmPlanes = 1
   tbm.bmBitsPixel = 1
   tbm.bmBits = VarPtr(m_lPattern(0))
   hBm = CreateBitmapIndirect(tbm)

   ' Make a brush from the bitmap bits
   m_hBrush = CreatePatternBrush(hBm)

   '// Delete the useless bitmap
   DeleteObject hBm

End Function
Private Sub DestroyBrush()
   If Not (m_hBrush = 0) Then
      DeleteObject m_hBrush
      m_hBrush = 0
   End If
End Sub

Private Function getHitCode() As EHitTestAreas
   
   Dim tR As RECT
   Dim tP As POINTAPI
   Dim eCode As EHitTestAreas
   
   GetClientRect m_hWnd, tR
   GetCursorPos tP
   ScreenToClient m_hWnd, tP
   Select Case m_eAlign
   Case vbAlignBottom
      tR.bottom = tR.top + m_lSize
      eCode = HTTOP
   Case vbAlignTop
      tR.top = tR.bottom - m_lSize
      eCode = HTBOTTOM
   Case vbAlignLeft
      tR.left = tR.right - m_lSize
      eCode = HTRIGHT
   Case vbAlignRight
      tR.right = tR.left + m_lSize
      eCode = HTLEFT
   End Select
   If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
      getHitCode = eCode
   Else
      getHitCode = HTNOWHERE
   End If

End Function

Private Sub pStartDrag()
Dim tR As RECT
   
   If Not (m_bInDrag) Then
      m_bInDrag = True
      SetCapture m_hWndParent
      GetWindowRect m_hWndParent, tR
      ClipCursorRect tR
      
      GetWindowRect m_hWnd, tR
      ' Draw the splitter:
      LSet m_tSplitR = tR
      Select Case m_eAlign
      Case vbAlignLeft
         m_tSplitR.left = m_tSplitR.right - m_lSize
      Case vbAlignRight
         m_tSplitR.right = m_tSplitR.left + m_lSize
      Case vbAlignTop
         m_tSplitR.top = m_tSplitR.bottom - m_lSize
      Case vbAlignBottom
         m_tSplitR.bottom = m_tSplitR.top + m_lSize
      End Select
      
      Dim tP As POINTAPI
      GetCursorPos tP
      m_tOffset.x = tP.x
      m_tOffset.y = tP.y
      
      pDrawSplitter
   End If
   
End Sub
Private Sub pDrag()

   If m_bInDrag Then
      ' Erase the last splitter:
      pDrawSplitter
      
      ' move to new position:
      getNewValidPosition
   
      ' Draw the new splitter:
      pDrawSplitter
   End If

End Sub
Private Sub pEndDrag()
   
   If m_bInDrag Then
      ' No longer in drag:
      m_bInDrag = False
      
      ' Clear cursor clipping and capture:
      ClipCursorClear 0&
      ReleaseCapture
      
      ' Erase the splitter:
      pDrawSplitter
      
      ' Move the window to the new position:
      Dim tR As RECT
      GetWindowRect m_hWnd, tR
      Select Case m_eAlign
      Case vbAlignLeft
         tR.right = m_tSplitR.right
      Case vbAlignRight
         tR.left = m_tSplitR.left
      Case vbAlignTop
         tR.bottom = m_tSplitR.bottom
      Case vbAlignBottom
         tR.top = m_tSplitR.top
      End Select
      SetWindowPos m_hWnd, 0, tR.left, tR.top, tR.right - tR.left, tR.bottom - tR.top, 0
      
   End If
   
End Sub
Private Sub pDrawSplitter()
Dim lhDC As Long
Dim hOldBrush As Long
   lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   hOldBrush = SelectObject(lhDC, m_hBrush)
   PatBlt lhDC, m_tSplitR.left, m_tSplitR.top, m_tSplitR.right - m_tSplitR.left, m_tSplitR.bottom - m_tSplitR.top, PATINVERT
   SelectObject lhDC, hOldBrush
   DeleteDC lhDC
End Sub
Private Sub getNewValidPosition()
   
   Dim tP As POINTAPI
   Dim tNewR As RECT
   Dim bNoOffset As Boolean
   
   GetCursorPos tP
   LSet tNewR = m_tSplitR
   Select Case m_eAlign
   Case vbAlignBottom, vbAlignTop
      OffsetRect tNewR, 0, (tP.y - m_tOffset.y)
   Case vbAlignLeft, vbAlignRight
      OffsetRect tNewR, (tP.x - m_tOffset.x), 0
   End Select
      
   ' Get the current size:
   Dim tRCurrent As RECT
   Dim tRMdi As RECT
   Dim lDiff As Long
   GetWindowRect m_hWnd, tRCurrent
   
   ' Get the size of the MDI Client area if any:
   Dim hWndMdi As Long
   hWndMdi = FindWindowEx(GetParent(m_hWnd), 0, "MDIClient", ByVal 0&)
   If Not (hWndMdi = 0) Then
      GetClientRect hWndMdi, tRMdi
   Else
      GetClientRect GetParent(m_hWnd), tRMdi
   End If

   Select Case m_eAlign
   Case vbAlignBottom
      If (tRCurrent.bottom - tNewR.top) < m_lMinSize Then
         tNewR.top = tRCurrent.bottom - m_lMinSize
         tNewR.bottom = tNewR.top + m_lSize
         bNoOffset = True
      ElseIf (tRCurrent.bottom - tNewR.top) > m_lMaxSize Then
         tNewR.top = tRCurrent.bottom - m_lMaxSize
         tNewR.bottom = tNewR.top + m_lSize
         bNoOffset = True
      End If
      If Not (hWndMdi = 0) Then
         lDiff = tNewR.top - tRCurrent.top
         If ((tRMdi.bottom - tRMdi.top + lDiff) < m_lSize) Then
            tNewR.top = m_tSplitR.top
            tNewR.bottom = m_tSplitR.bottom
            bNoOffset = True
         End If
      End If
   Case vbAlignTop
      If (tNewR.bottom - tRCurrent.top) < m_lMinSize Then
         tNewR.bottom = tRCurrent.top + m_lMinSize
         tNewR.top = tNewR.bottom - m_lSize
         bNoOffset = True
      ElseIf (tNewR.bottom - tRCurrent.top) > m_lMaxSize Then
         tNewR.bottom = tRCurrent.top + m_lMaxSize
         tNewR.top = tNewR.bottom - m_lSize
         bNoOffset = True
      End If
      If Not (hWndMdi = 0) Then
         lDiff = tRCurrent.bottom - tNewR.bottom
         If ((tRMdi.bottom - tRMdi.top + lDiff) < m_lSize) Then
            tNewR.top = m_tSplitR.top
            tNewR.bottom = m_tSplitR.bottom
            bNoOffset = True
         End If
      End If
   Case vbAlignRight
      If (tRCurrent.right - tNewR.left) < m_lMinSize Then
         tNewR.left = tRCurrent.right - m_lMinSize
         tNewR.right = tNewR.left + m_lSize
         bNoOffset = True
      ElseIf (tRCurrent.right - tNewR.left) > m_lMaxSize Then
         tNewR.left = tRCurrent.right - m_lMaxSize
         tNewR.right = tNewR.left + m_lSize
         bNoOffset = True
      End If
      If Not (hWndMdi = 0) Then
         lDiff = tNewR.left - tRCurrent.left
         If ((tRMdi.right - tRMdi.left + lDiff) < m_lSize) Then
            tNewR.left = m_tSplitR.left
            tNewR.right = m_tSplitR.right
            bNoOffset = True
         End If
      End If
   Case vbAlignLeft
      If (tNewR.right - tRCurrent.left) < m_lMinSize Then
         tNewR.right = tRCurrent.left + m_lMinSize
         tNewR.left = tNewR.right - m_lSize
         bNoOffset = True
      ElseIf (tNewR.right - tRCurrent.left) > m_lMaxSize Then
         tNewR.right = tRCurrent.left + m_lMaxSize
         tNewR.left = tNewR.right - m_lSize
         bNoOffset = True
      End If
      If Not (hWndMdi = 0) Then
         lDiff = tRCurrent.right - tNewR.right
         If ((tRMdi.right - tRMdi.left + lDiff) < m_lSize) Then
            tNewR.left = m_tSplitR.left
            tNewR.right = m_tSplitR.right
            bNoOffset = True
         End If
      End If
   End Select
   
   ' Validate position:
   LSet m_tSplitR = tNewR
   
   If Not (bNoOffset) Then
      m_tOffset.x = tP.x
      m_tOffset.y = tP.y
   End If

End Sub

Private Sub Class_Initialize()
   m_lSize = 4
   m_lMinSize = 16
   m_lMaxSize = &H7FFFFFFF
   m_bFullDrag = True

   Dim i As Long
   For i = 0 To 3
      m_lPattern(i) = &HAAAA5555
   Next i
End Sub

Private Sub Class_Terminate()
   Detach
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
   Select Case CurrentMessage
   Case WM_NCHITTEST, WM_SETCURSOR
      ISubClass_MsgResponse = emrConsume
   Case Else
      ISubClass_MsgResponse = emrPostProcess
   End Select
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
   Select Case iMsg
   
   Case WM_NCHITTEST
      ' Full Drag:
      Dim eCode As EHitTestAreas
      eCode = getHitCode()
      If (eCode = HTNOWHERE) Then
         ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      Else
         ISubClass_WindowProc = eCode
      End If
   
   Case WM_SIZING
      ' Full Drag sizing:
      Dim tR As RECT
      Dim tRCurrent As RECT
      Dim tRMdi As RECT
      Dim lDiff As Long
      
      ' Get the proposed size:
      CopyMemory tR, ByVal lParam, Len(tR)
      ' Get the current size:
      GetWindowRect hwnd, tRCurrent
      
      ' Get the size of the MDI Client area if any:
      Dim hWndMdi As Long
      hWndMdi = FindWindowEx(GetParent(hwnd), 0, "MDIClient", ByVal 0&)
      If Not (hWndMdi = 0) Then
         GetClientRect hWndMdi, tRMdi
      Else
         GetClientRect GetParent(hwnd), tRMdi
      End If
      
      ' Check if the object is too big or small
      Select Case m_eAlign
      Case vbAlignBottom
         If (tR.bottom - tR.top) < m_lMinSize Then
            tR.top = tR.bottom - m_lMinSize
         ElseIf (tR.bottom - tR.top) > m_lMaxSize Then
            tR.top = tR.bottom - m_lMaxSize
         End If
         If Not (hWndMdi = 0) Then
            lDiff = ((tRCurrent.bottom - tRCurrent.top) - (tR.bottom - tR.top))
            If ((tRMdi.bottom - tRMdi.top + lDiff) < m_lSize) Then
               tR.top = tRCurrent.top
            End If
         End If
      Case vbAlignTop
         If (tR.bottom - tR.top) < m_lMinSize Then
            tR.bottom = tR.top + m_lMinSize
         ElseIf (tR.bottom - tR.top) > m_lMaxSize Then
            tR.bottom = tR.top + m_lMaxSize
         End If
         If Not (hWndMdi = 0) Then
            lDiff = ((tRCurrent.bottom - tRCurrent.top) - (tR.bottom - tR.top))
            If ((tRMdi.bottom - tRMdi.top + lDiff) < m_lSize) Then
               tR.bottom = tRCurrent.bottom
            End If
         End If
      Case vbAlignRight
         If (tR.right - tR.left) < m_lMinSize Then
            tR.left = tR.right - m_lMinSize
         ElseIf (tR.right - tR.left) > m_lMaxSize Then
            tR.left = tR.right - m_lMaxSize
         End If
         If Not (hWndMdi = 0) Then
            lDiff = ((tRCurrent.right - tRCurrent.left) - (tR.right - tR.left))
            If ((tRMdi.right - tRMdi.left + lDiff) < m_lSize) Then
               tR.left = tRCurrent.left
            End If
         End If
      Case vbAlignLeft
         If (tR.right - tR.left) < m_lMinSize Then
            tR.right = tR.left + m_lMinSize
         ElseIf (tR.right - tR.left) > m_lMaxSize Then
            tR.right = tR.left + m_lMaxSize
         End If
         If Not (hWndMdi = 0) Then
            lDiff = ((tRCurrent.right - tRCurrent.left) - (tR.right - tR.left))
            If ((tRMdi.right - tRMdi.left + lDiff) < m_lSize) Then
               tR.right = tRCurrent.right
            End If
         End If
      End Select
      
      CopyMemory ByVal lParam, tR, Len(tR)
   
   Case WM_ACTIVATE
      ' non-full drag, check for switch focus:
      pEndDrag
      
   Case WM_SETCURSOR
      ' non-full drag
      eCode = getHitCode()
      Select Case eCode
      Case HTTOP, HTBOTTOM
         SetCursor m_hCur
         ISubClass_WindowProc = 1
      Case HTRIGHT, HTLEFT
         SetCursor m_hCur
         ISubClass_WindowProc = 1
      Case Else
         ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End Select
      
   Case WM_LBUTTONDOWN
      If Not (getHitCode() = HTNOWHERE) Then
         pStartDrag
      End If
   
   Case WM_MOUSEMOVE
      pDrag
   
   Case WM_LBUTTONUP
      pEndDrag
      
   End Select
   
End Function




