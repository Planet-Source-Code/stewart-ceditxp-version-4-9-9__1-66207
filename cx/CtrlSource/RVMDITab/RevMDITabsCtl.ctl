VERSION 5.00
Begin VB.UserControl RevMDITabsCtl 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   Picture         =   "RevMDITabsCtl.ctx":0000
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ToolboxBitmap   =   "RevMDITabsCtl.ctx":0CCA
   Begin VB.PictureBox picDefault 
      Height          =   315
      Left            =   840
      Picture         =   "RevMDITabsCtl.ctx":0FDC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   720
      Width           =   315
   End
End
Attribute VB_Name = "RevMDITabsCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
'*    Author      : Andrea Batina[Revelatek]
'*    Date        : 11/05/2004
'*
'*    Component   : RevMDITabsCtl
'*    Version     : 1.05
'*
'*    Description : Revelatek MDI Tabs control provides you with the ability to
'*                  have Visual Studio.NET, Office 2003 and Office 2000 style tabs.
'*
'*                  To use this control you don't have to write a single line of code
'*                  you just need to put it on MDI form and that's all!
'*
'*    Dependencies: None.
'*
'*    Credits     : Part of the code is borrowed from vbAccelerator MDITabs Control.
'*
'*    Copyright   : Copyright © 2004 Andrea Batina. All rights reserved.
'*
'*    Notes       : If you find any bugs then please report them to me at
'*                  <a_batina@hotmail.com> so that I can fix them ASAP.
'*
'*    History     :
'*           v1.00
'*                  11/05/2004 - Inital version.
'*           v1.01
'*                  11/18/2004 - Fixed minor bug. When closing the MDI child form
'*                               using "x" button QueryUnload and Unload events didn't
'*                               fire correctly. Thanks Guilect for reporting it.
'*           v1.02
'*                  11/18/2004 - Improved drawing code when all MDI child forms are
'*                               closed so that now border is drawn across whole border
'*                               side. Thanks to Phantom Man for suggesting this.
'*           v1.03
'*                  11/18/2004 - Fixed bug in pGetThemeName function. It was called in
'*                               Windows 2000 and raised error. Thanks to Guilect
'*                               for reporting it.
'*           v1.04
'*                  11/18/2004 - Improved drawing function. Old drawing functions drawed
'*                               all tabs including the ones that were not visible on screen
'*                               and produced flicker when there were 20 or more tabs.
'*                  11/19/2004 - No more flickering while switching between two MDI
'*                               child forms, thanks to Neal Rushforth.
'*                  11/19/2004 - Added optional drawing of focus rect.
'*                  11/19/2004 - Added support for drawing form icons.
'*           v1.05
'*                  11/19/2004 - Fixed bug that caused crash when dragging tab to empty space.
'*                               Thanks to The New iSoftware Company! for reporting it, and
'*                               thanks to Carles P.V. and Phantom Man for helping me fix it.
'*                  11/19/2004 - Improved focus rectangle drawing. Now when app lost focus
'*                               the focus rect is not drawn. Thanks to Carles P.V. for suggesting
'*                               it.
'*
'*******************************************************************************

'TODO:
'      - *v2* Add color properties so that user can customize controls color sheme.

Option Explicit

'////////////////////////////////////////////////////////////////////
'// Implement subclassing interface
Implements ISubclassingSink

'////////////////////////////////////////////////////////////////////
'// Private/Public Type Definitions
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cX As Long
    cY As Long
    flags As Long
End Type
Private Type NCCALCSIZE_PARAMS
    rgrc(0 To 2) As RECT
    lppos As Long
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

'////////////////////////////////////////////////////////////////////
'// Private/Public Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetFocus Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function DrawFocusRectAPI Lib "user32.dll" Alias "DrawFocusRect" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long

'////////////////////////////////////////////////////////////////////
'// Private/Public Constant Declarations
' Window Messages
Private Const WM_SETCURSOR = &H20
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCPAINT = &H85
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_CLOSE = &H10
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_GETICON As Long = &H7F
Private Const WM_ACTIVATE As Long = &H6

' Extended Window Style Messages
Private Const WS_EX_CLIENTEDGE = &H200

' GetWindowLong Messages
Private Const GWL_EXSTYLE = (-20)

' NonClientCalculateSize Messages
Private Const WVR_VALIDRECTS = &H400

' SetWindowPosition Messages
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_FRAMECHANGED = &H20

' RedrawWindow Messages
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ALLCHILDREN As Long = &H80

' Misc MEssages
Private Const PS_SOLID = 0
Private Const HTNOWHERE = 0
Private Const ALTERNATE = 1
Private Const IMAGE_ICON = 1
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Const LR_COPYFROMRESOURCE = &H4000

'////////////////////////////////////////////////////////////////////
'// Private/Public Enum Definitions
Public Enum EMTStyle
    mtsOfficeXP
    mtsOffice2000
    mtsOffice2003
End Enum

'////////////////////////////////////////////////////////////////////
'// Private/Public Event Declarations
Public Event TabBarClick(Button As Integer, x As Long, y As Long)
Public Event TabClick(TabHwnd As Long, Button As Integer, x As Long, y As Long)
Public Event ColorChanged(NewColor As OLE_COLOR)

'////////////////////////////////////////////////////////////////////
'// Private/Public Variable Declarations
Private m_oSubclass             As CEasySubclass_v1     ' Main subclasser object
Private m_oParentSubclass       As CEasySubclass_v1     ' Main subclasser object for parent(MDI form)
Private m_oMemDC                As CMemoryDC            ' Main memory dc object
Private m_oMemDCTabs            As CMemoryDC            ' Main tab memory dc object
Private WithEvents m_oTimer     As CTimer               ' Timer object
Attribute m_oTimer.VB_VarHelpID = -1

Private m_lMDIClient            As Long                 ' MDI client window handle
Private m_lTabHeight            As Long                 ' Tab control height
Private m_bHasFocus             As Boolean              ' Does tab control have focus?

' Tab drag variables
Private m_lDraggingTab          As Long                 ' Dragging tab ID
Private m_bJustReplaced         As Long
Private m_tJustReplacedPoint    As POINTAPI             ' Replace point coordinates

' Internal MDI child window variables
Private m_hWndTempChild         As Collection           ' Temporary MDI child windows collection
Private m_hWndChild             As Collection           ' MDI child windows collection
Private m_tTabR()               As RECT                 ' Tab items dimensions
Private m_lLastSelMDIChild      As Long                 ' Last selected MDI child window handle

' Drawing variables
Private m_tCloseBtnRect         As RECT                 ' Close, Next and Prev buttons positions
Private m_tNextBtnRect          As RECT                 ' Close, Next and Prev buttons positions
Private m_tPrevBtnRect          As RECT                 ' Close, Next and Prev buttons positions
Private m_lSelBtn               As Long                 ' Selected button ID
Private m_lPressedBtn           As Long                 ' Pressed button ID
Private m_lOffsetX              As Long                 ' Left offset position for tab drawing
Private m_lButtonsSize          As Long                 ' Buttons size
Private m_lColorTable(1 To 8)   As Long                 ' Office2003 style colors
Private m_oColorIndex           As Collection           '
Private m_lClrIndex             As Long

' Color variables
Private m_lClrBack              As Long                 ' Tab background color
Private m_lClrTabBack1          As Long                 ' Tab item background color 1
Private m_lClrTabFore           As Long                 ' Tab item fore color
Private m_lClrTabInactiveFore   As Long                 ' Tab item inactive fore color
Private m_lClrTabBorder         As Long                 ' Tab item light border color
Private m_lClrTabBorderDK       As Long                 ' Tab item dark border color
Private m_lClrOuterBorder       As Long                 ' Tab outer border color
Private m_lClrInnerBorder       As Long                 ' Tab inner border color
Private m_lClrBorder            As Long                 ' Tab border color
Private m_lClrButton            As Long                 ' Button-arrow color
Private m_lClrButtonBorder      As Long                 ' Button light border color
Private m_lClrButtonBorderDK    As Long                 ' Button dark border color
Private m_lClrTabSeparator      As Long                 ' Tab item separator (Xp Style Only)

' MDITabs Properties
Private m_fFont                 As StdFont              ' Font object
Private m_eStyle                As EMTStyle             ' Tab control drawing style
Private m_bDrawFocusRect        As Boolean              ' Should we draw focus rect?
Private m_bDrawIcons            As Boolean              ' Should we draw form icons?

' Visual Style Theme variables
Private m_sThemeName            As String               ' Current window theme name

'//////////////////////////////////////////////////////////////////////////////
'//// PUBLIC PROPERTIES
'//////////////////////////////////////////////////////////////////////////////
Public Property Get DrawIcons() As Boolean
    DrawIcons = m_bDrawIcons
End Property
Public Property Let DrawIcons(ByVal Value As Boolean)
    If Value <> m_bDrawIcons Then
        m_bDrawIcons = Value
        pDrawControl m_lMDIClient
    End If
    PropertyChanged "DrawIcons"
End Property
Public Property Get DrawFocusRect() As Boolean
    DrawFocusRect = m_bDrawFocusRect
End Property
Public Property Let DrawFocusRect(ByVal Value As Boolean)
    If Value <> m_bDrawFocusRect Then
        m_bDrawFocusRect = Value
        pDrawControl m_lMDIClient
    End If
    PropertyChanged "DrawFocusRect"
End Property
Public Property Get Style() As EMTStyle
    Style = m_eStyle
End Property
Public Property Let Style(ByVal Value As EMTStyle)
    If Value <> m_eStyle Then
        m_eStyle = Value
        pSetColors
        pDrawControl m_lMDIClient
    End If
    PropertyChanged "Style"
End Property
Public Property Get Font() As StdFont
    Set Font = m_fFont
End Property
Public Property Let Font(ByVal Value As StdFont)
    If Value <> m_fFont Then
        Set m_fFont = Value
        pDrawControl m_lMDIClient
    End If
    PropertyChanged "Font"
End Property
Public Property Set Font(ByVal Value As StdFont)
    If Value <> m_fFont Then
        Set m_fFont = Value
        pDrawControl m_lMDIClient
    End If
    PropertyChanged "Font"
End Property

'//////////////////////////////////////////////////////////////////////////////
'//// MAIN FUNCTIONS
'//////////////////////////////////////////////////////////////////////////////
'********************************************************************
'* Name: pDrawControl
'* Description: Draw MDI tabs.
'********************************************************************
Private Sub pDrawControl(ByVal lhWnd As Long)
    If m_oMemDC Is Nothing Then Exit Sub
    Select Case m_eStyle
        Case mtsOfficeXP
            pDrawTabsXPStyle lhWnd
        Case mtsOffice2000
            pDrawTabs2000Style lhWnd
        Case mtsOffice2003
            pDrawTabs2003Style lhWnd
    End Select
End Sub
Private Sub pDrawTabs2003Style(ByVal lhWnd As Long)
    Dim i As Long
    Dim tR As RECT
    Dim tT As RECT
    Dim lPenOld As Long
    Dim lPen As Long
    Dim tPA As POINTAPI
    Dim lCHDC As Long
    Dim hWndChild As Variant
    Dim lTabLeft As Long
    Dim lTabTop As Long
    Dim sCaption As String
    Dim lSelMDIChild As Long
    Dim lSelTabID As Long
    Dim lBtnOffset As Long
    Dim tRgnBox As RECT
    Dim tPoly() As POINTAPI
    Dim lPolyCount As Long
    Dim lRgn As Long
    Dim lHighlightClr As Long
    Dim bIntersect As Long
    
    ' Get MDI client window dc
    lCHDC = GetWindowDC(lhWnd)
    ' Get MDI client dimensions
    GetWindowRect lhWnd, tR
    OffsetRect tR, -tR.Left, -tR.Top
    ' Get active MDI Child handle
    pGetMDIChildWindows
    If m_hWndChild.Count > 0 Then
        lSelMDIChild = SendMessage(m_lMDIClient, WM_MDIGETACTIVE, 0, 0)
        ' For each MDI child window
        For Each hWndChild In m_hWndChild
            i = i + 1
            If lSelMDIChild = hWndChild Then Exit For
        Next
    End If
    
    If lSelMDIChild <> 0 Then
        lHighlightClr = m_lColorTable(m_oColorIndex(i))
    Else
        lHighlightClr = vbApplicationWorkspace
    End If
    
    With m_oMemDC
        ' Initialize memory dc
        .Width = Abs(tR.Right - tR.Left)
        .Height = Abs(tR.Bottom - tR.Top) + 1

        '============================================================
        '== BORDERS
        ' Draw outter border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vbButtonShadow))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Left, m_lTabHeight - 3
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vbButtonShadow))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Right - 1, m_lTabHeight - 3, tPA
        LineTo lCHDC, tR.Right - 1, tR.Bottom - 1
        LineTo lCHDC, tR.Left - 1, tR.Bottom - 1
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        ' Draw inner border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(lHighlightClr))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 1, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Left + 1, m_lTabHeight - 2
        LineTo lCHDC, tR.Right - 2, m_lTabHeight - 2
        MoveToEx lCHDC, tR.Left + 1, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Right - 1, tR.Bottom - 2
        MoveToEx lCHDC, tR.Right - 2, m_lTabHeight - 2, tPA
        LineTo lCHDC, tR.Right - 2, tR.Bottom - 2
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        ' Draw inner-inner border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vbButtonShadow))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 2, tR.Bottom - 3, tPA
        LineTo lCHDC, tR.Left + 2, m_lTabHeight - 1
        LineTo lCHDC, tR.Right - 3, m_lTabHeight - 1
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vb3DHighlight))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 2, tR.Bottom - 3, tPA
        LineTo lCHDC, tR.Right - 2, tR.Bottom - 3
        MoveToEx lCHDC, tR.Right - 3, m_lTabHeight - 1, tPA
        LineTo lCHDC, tR.Right - 3, tR.Bottom - 3
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        '============================================================
        ' Draw bar background
        .DrawGradient tR.Left, tR.Top, tR.Right, m_lTabHeight - 3, .LightenColor(.TranslateColor(vbButtonFace), .TranslateColor(vbWindowBackground), 205), vbButtonFace
        ' Draw bar highlight line
        .DrawLine tR.Left, m_lTabHeight - 3, tR.Right - 1, m_lTabHeight - 3, vbButtonShadow
        ' Draw tab highight line
        .DrawLine tR.Left + 1, m_lTabHeight - 2, tR.Right - 1, m_lTabHeight - 2, lHighlightClr
        .DrawLine tR.Left + 2, m_lTabHeight - 1, tR.Right - 2, m_lTabHeight - 1, vbButtonShadow
        ' Fix line end pixel
        .SetPixel tR.Right - 1, m_lTabHeight - 3, .TranslateColor(vbButtonShadow)
        ' left
        .DrawLine tR.Left, m_lTabHeight - 2, tR.Left, m_lTabHeight, vbButtonShadow
        .SetPixel tR.Left + 1, m_lTabHeight - 1, .TranslateColor(lHighlightClr)
        ' right
        .DrawLine tR.Right - 1, m_lTabHeight - 2, tR.Right - 1, m_lTabHeight, vbButtonShadow
        .SetPixel tR.Right - 2, m_lTabHeight - 1, .TranslateColor(lHighlightClr)
        '== END BORDERS
        '============================================================
    End With

    ' If there are open MDI child windows
    If m_hWndChild.Count > 0 Then
        With m_oMemDCTabs
            ' Initialize memory dc
            .Height = m_lTabHeight

            i = 0
            lTabLeft = 18
            ReDim m_tTabR(1 To m_hWndChild.Count) As RECT
            ' For each MDI child window
            For Each hWndChild In m_hWndChild
                i = i + 1
                ' Get MDI child caption
                sCaption = String$(256, vbNullChar)
                GetWindowText hWndChild, sCaption, 256
                sCaption = Left$(sCaption, InStr(1, sCaption, vbNullChar) - 1)
                ' If current tab is selected
                If lSelMDIChild = hWndChild Then
                    m_fFont.Bold = True
                    Set .Font = m_fFont
                    lSelTabID = i
                Else
                    m_fFont.Bold = False
                    Set .Font = m_fFont
                End If
                ' Save tab dimensions
                m_tTabR(i).Left = lTabLeft
                m_tTabR(i).Top = 3
                m_tTabR(i).Bottom = m_lTabHeight - 2 '                  gap=12(6*2)
                m_tTabR(i).Right = m_tTabR(i).Left + .TextWidth(sCaption) + 12
                If m_bDrawIcons Then
                    m_tTabR(i).Right = m_tTabR(i).Right + 20
                End If
                lTabLeft = m_tTabR(i).Right - 1
            Next

            If lSelMDIChild <> m_lLastSelMDIChild And lSelTabID > 0 Then
                RaiseEvent ColorChanged(m_lColorTable(m_oColorIndex(i)))
                m_lLastSelMDIChild = lSelMDIChild
                ' Ensure that a newly selected tab is scrolled into view 49=buttonbar size
                If m_tTabR(lSelTabID).Left - m_lOffsetX < tR.Left Then
                    m_lOffsetX = m_tTabR(lSelTabID).Left - 30
                ElseIf m_tTabR(lSelTabID).Right - m_lOffsetX > tR.Right - 49 Then
                    m_lOffsetX = m_lOffsetX + ((m_tTabR(lSelTabID).Right - m_lOffsetX) - (tR.Right - 49)) + 30
                    If m_tTabR(m_hWndChild.Count).Right - m_lOffsetX = (tR.Right - 49) - 30 Then
                        m_lOffsetX = m_lOffsetX - 30
                    End If
                End If
                If m_lOffsetX <= 30 Then m_lOffsetX = 0
            End If

            ' Initialize memory dc
            .Width = m_tTabR(m_hWndChild.Count).Right + 10
            .Cls vbButtonFace

            ' Draw control background
            .DrawGradient 0, 0, .Width, m_lTabHeight - 3, .LightenColor(.TranslateColor(vbButtonFace), .TranslateColor(vbWindowBackground), 205), vbButtonFace
            ' Draw highlight line
            .DrawLine 0, m_lTabHeight - 3, .Width, m_lTabHeight - 3, vbButtonShadow
            .DrawLine 0, m_lTabHeight - 2, .Width, m_lTabHeight - 2, lHighlightClr
            .DrawLine 0, m_lTabHeight - 1, .Width, m_lTabHeight - 1, vbButtonShadow
            i = 0
            ' For each MDI child window
            For Each hWndChild In m_hWndChild
                i = i + 1
                               
                ' See if rects intersect (if tab is visible)
                m_tTabR(i).Left = m_tTabR(i).Left - m_lOffsetX
                m_tTabR(i).Right = m_tTabR(i).Right - m_lOffsetX
                bIntersect = IntersectRect(tT, tR, m_tTabR(i))
                m_tTabR(i).Left = m_tTabR(i).Left + m_lOffsetX
                m_tTabR(i).Right = m_tTabR(i).Right + m_lOffsetX
                
                ' If it is visible then draw it
                If bIntersect Then
                    ' Get MDI child caption
                    sCaption = String$(256, vbNullChar)
                    GetWindowText hWndChild, sCaption, 256
                    sCaption = Left$(sCaption, InStr(1, sCaption, vbNullChar) - 1)
                
                    If i = 1 Or lSelMDIChild = hWndChild Then
                        ReDim tPoly(1 To 5) As POINTAPI
                        tPoly(1).x = m_tTabR(i).Left - 12
                        tPoly(1).y = m_tTabR(i).Bottom - 2
                        tPoly(2).x = m_tTabR(i).Left + 3
                        tPoly(2).y = m_tTabR(i).Top + 3
                        tPoly(3).x = m_tTabR(i).Left + 8
                        tPoly(3).y = m_tTabR(i).Top + 1
                        tPoly(4).x = m_tTabR(i).Right - 2
                        tPoly(4).y = m_tTabR(i).Top + 2
                        tPoly(5).x = m_tTabR(i).Right - 3
                        tPoly(5).y = m_tTabR(i).Bottom - 1
                        lPolyCount = 5
    
                        ' Draw tab border
                        .SetPixel m_tTabR(i).Left - 15, m_tTabR(i).Bottom - 2, .TranslateColor(vbButtonShadow)
                        .DrawLine m_tTabR(i).Left - 14, m_tTabR(i).Bottom - 2, m_tTabR(i).Left + 2, m_tTabR(i).Top + 2, vbButtonShadow
                        .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top + 2, m_tTabR(i).Left + 6, m_tTabR(i).Top + 1, vbButtonShadow
                        .DrawLine m_tTabR(i).Left + 6, m_tTabR(i).Top, m_tTabR(i).Right - 2, m_tTabR(i).Top, vbButtonShadow
                        .SetPixel m_tTabR(i).Right - 2, m_tTabR(i).Top + 1, .TranslateColor(vbButtonShadow)
                        .DrawLine m_tTabR(i).Right - 1, m_tTabR(i).Top + 2, m_tTabR(i).Right - 1, m_tTabR(i).Bottom, vbButtonShadow
                        If lSelMDIChild = hWndChild Then
                            lHighlightClr = vb3DHighlight
                        Else
                            lHighlightClr = vbButtonFace
                        End If
                        ' Draw tab highlight line
                        .DrawLine m_tTabR(i).Left - 13, m_tTabR(i).Bottom - 2, m_tTabR(i).Left + 2, m_tTabR(i).Top + 3, lHighlightClr
                        .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top + 3, m_tTabR(i).Left + 6, m_tTabR(i).Top + 2, lHighlightClr
                        .DrawLine m_tTabR(i).Left + 6, m_tTabR(i).Top + 1, m_tTabR(i).Right - 2, m_tTabR(i).Top + 1, lHighlightClr
                        .DrawLine m_tTabR(i).Right - 2, m_tTabR(i).Top + 2, m_tTabR(i).Right - 2, m_tTabR(i).Bottom, lHighlightClr
    
                        ' Erase MDI client window border
                        If lSelMDIChild = hWndChild Then
                            .DrawLine m_tTabR(i).Left - 15, m_tTabR(i).Bottom - 1, m_tTabR(i).Right, m_tTabR(i).Bottom - 1, m_lColorTable(m_oColorIndex(i))
                        Else
                            .DrawLine m_tTabR(i).Left - 15, m_tTabR(i).Bottom - 1, m_tTabR(i).Right, m_tTabR(i).Bottom - 1, vbButtonShadow
                        End If
                    Else
                        ReDim tPoly(1 To 5) As POINTAPI
                        tPoly(1).x = m_tTabR(i).Left + 1
                        tPoly(1).y = m_tTabR(i).Bottom - 1
                        tPoly(2).x = m_tTabR(i).Left + 1
                        tPoly(2).y = m_tTabR(i).Top + 4
                        tPoly(3).x = m_tTabR(i).Left + 8
                        tPoly(3).y = m_tTabR(i).Top + 1
                        tPoly(4).x = m_tTabR(i).Right - 2
                        tPoly(4).y = m_tTabR(i).Top + 2
                        tPoly(5).x = m_tTabR(i).Right - 3
                        tPoly(5).y = m_tTabR(i).Bottom - 1
                        lPolyCount = 5
    
                        ' Draw tab border
                        .SetPixel m_tTabR(i).Left + 1, m_tTabR(i).Top + 3, .TranslateColor(vbButtonShadow)
                        .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top + 2, m_tTabR(i).Left + 6, m_tTabR(i).Top + 1, vbButtonShadow
                        .DrawLine m_tTabR(i).Left + 6, m_tTabR(i).Top, m_tTabR(i).Right - 2, m_tTabR(i).Top, vbButtonShadow
                        .SetPixel m_tTabR(i).Right - 2, m_tTabR(i).Top + 1, .TranslateColor(vbButtonShadow)
                        .DrawLine m_tTabR(i).Right - 1, m_tTabR(i).Top + 2, m_tTabR(i).Right - 1, m_tTabR(i).Bottom, vbButtonShadow
                        ' Draw tab highlight line
                        .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top + 3, m_tTabR(i).Left + 6, m_tTabR(i).Top + 2, vbButtonFace
                        .DrawLine m_tTabR(i).Left + 6, m_tTabR(i).Top + 1, m_tTabR(i).Right - 2, m_tTabR(i).Top + 1, vbButtonFace
                        .DrawLine m_tTabR(i).Right - 2, m_tTabR(i).Top + 2, m_tTabR(i).Right - 2, m_tTabR(i).Bottom, vbButtonFace
                    End If
    
                    ' Draw tab gradient background
                    lRgn = CreatePolygonRgn(tPoly(1), lPolyCount, ALTERNATE)
                    SelectClipRgn .hdc, lRgn
                    GetRgnBox lRgn, tRgnBox
                    .DrawGradient tRgnBox.Left, tRgnBox.Top, tRgnBox.Right, tRgnBox.Bottom, m_lColorTable(m_oColorIndex(i)), .LightenColor(m_lColorTable(m_oColorIndex(i)), vbWhite, 240)
                    SelectClipRgn .hdc, 0
                    DeleteObject lRgn
                    
                    ' If current tab is selected
                    If lSelMDIChild = hWndChild Then
                        ' Set text color to default (black) color
                        .ForeColor = .TranslateColor(vbWindowText)
                        m_fFont.Bold = True
                        Set .Font = m_fFont
                        ' Should we draw focus rect?
                        If m_bDrawFocusRect And m_bHasFocus Then
                            LSet tT = m_tTabR(i)
                            With tT
                                .Left = .Left + 3
                                .Top = .Top + 3
                                .Right = .Right - 4
                                .Bottom = .Bottom - 2
                            End With
                            ' Draw focus rect
                            DrawFocusRectAPI .hdc, tT
                        End If
                    Else
                        ' Set text color to lighter black color
                        .ForeColor = .TranslateColor(vbWindowText)
                        m_fFont.Bold = False
                        Set .Font = m_fFont
                    End If
                    ' Draw tab caption
                    lTabTop = m_tTabR(i).Top + (((m_tTabR(i).Bottom - m_tTabR(i).Top) / 2) - (.TextHeight(sCaption) / 2))
                    If m_bDrawIcons Then
                        .DrawText sCaption, m_tTabR(i).Left + 20, lTabTop, m_tTabR(i).Right, m_tTabR(i).Bottom, DT_CENTER
                        .DrawPicture pGetSmallIcon(CLng(hWndChild)), m_tTabR(i).Left + 5, m_tTabR(i).Top + 2
                    Else
                        .DrawText sCaption, m_tTabR(i).Left, lTabTop, m_tTabR(i).Right, m_tTabR(i).Bottom, DT_CENTER
                    End If
                End If
            Next
        End With

        With m_oMemDC
            '============================================================
            '== PREV, NEXT AND CLOSE BUTTONS
            lBtnOffset = 2
            ' Draw background for buttons
            LSet tT = tR
            tT.Left = tT.Right - 47
            tT.Top = tR.Top '+ 1
            tT.Bottom = m_lTabHeight - 2
            .DrawGradient tT.Left, tT.Top, tT.Right, tT.Bottom, .LightenColor(.TranslateColor(vbButtonFace), .TranslateColor(vbWindowBackground), 205), vbButtonFace
            .DrawLine tT.Left, tT.Bottom - 1, tT.Right, tT.Bottom - 1, vbButtonShadow
            m_lButtonsSize = 7 - lBtnOffset
            ' Draw close button
            m_tCloseBtnRect.Left = tT.Right - (19 - lBtnOffset)
            m_tCloseBtnRect.Top = 5
            m_tCloseBtnRect.Right = m_tCloseBtnRect.Left + 14
            m_tCloseBtnRect.Bottom = m_tCloseBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            ' Draw X sign using Marlett font
            Dim fFont As StdFont
            Set fFont = New StdFont
            fFont.Name = "Marlett"
            fFont.Size = 7
            Set .Font = fFont
            ' Draw button border
            If m_lSelBtn = 3 Then
                ' If it is pressed
                If m_lPressedBtn = 3 Then
                    .FillRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, .AlphaBlend(vbHighlight, .AlphaBlend(vbWindowBackground, vbButtonFace, 220), 85)
                    .ForeColor = vbWhite
                    .Draw3DRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, vbHighlight, vbHighlight
                    ' If it is selected
                Else
                    .FillRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, .AlphaBlend(vbHighlight, .AlphaBlend(vbWindowBackground, vbButtonFace, 220), 85)
                    .ForeColor = vbWindowText
                    .Draw3DRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, vbHighlight, vbHighlight
                End If
            End If
            ' If it is pressed
            If m_lPressedBtn = 3 Then
                ' Draw X sign
                .DrawText "r", m_tCloseBtnRect.Left + 4, m_tCloseBtnRect.Top + 4, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom
            Else
                ' Draw X sign
                .DrawText "r", m_tCloseBtnRect.Left + 3, m_tCloseBtnRect.Top + 3, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom
            End If
            '========================================
            ' Draw NEXT button
            m_tNextBtnRect.Left = m_tCloseBtnRect.Left - 14
            m_tNextBtnRect.Top = 5
            m_tNextBtnRect.Right = m_tNextBtnRect.Left + 14
            m_tNextBtnRect.Bottom = m_tNextBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            LSet tT = m_tNextBtnRect
            ' Draw button border
            If m_lSelBtn = 2 Then
                ' If it is pressed
                If m_lPressedBtn = 2 Then
                    .FillRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, .AlphaBlend(vbHighlight, .AlphaBlend(vbWindowBackground, vbButtonFace, 220), 85)
                    .Draw3DRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, vbHighlight, vbHighlight
                    ' If it is selected
                Else
                    .FillRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, .AlphaBlend(vbHighlight, .AlphaBlend(vbWindowBackground, vbButtonFace, 220), 85)
                    .Draw3DRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, vbHighlight, vbHighlight
                End If
            End If
            tT.Left = tT.Left + 4
            tT.Top = tT.Top + 2
            tT.Bottom = tT.Top + 9
            ' If it is pressed
            If m_lPressedBtn = 2 Then
                tT.Left = tT.Left + 1
                tT.Top = tT.Top + 1
                tT.Bottom = tT.Bottom + 1
            End If
            ' Draw arrow
            For i = 0 To 4
                .DrawLine tT.Left + i, tT.Top + i, tT.Left + i, tT.Bottom - i, vbWindowText
            Next
            If Not pIsNextButtonEnabled Then
                ' Draw empty arrows
                For i = 1 To 3
                    .DrawLine tT.Left + i, tT.Top + i + 1, tT.Left + i, tT.Bottom - i - 1, vbButtonFace
                Next
            End If
            '========================================
            ' Draw PREV button
            m_tPrevBtnRect.Left = m_tNextBtnRect.Left - 14
            m_tPrevBtnRect.Top = 5
            m_tPrevBtnRect.Right = m_tPrevBtnRect.Left + 14
            m_tPrevBtnRect.Bottom = m_tPrevBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            LSet tT = m_tPrevBtnRect
            ' Draw button border
            If m_lSelBtn = 1 Then
                ' If it is pressed
                If m_lPressedBtn = 1 Then
                    .FillRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, .AlphaBlend(vbHighlight, .AlphaBlend(vbWindowBackground, vbButtonFace, 220), 85)
                    .Draw3DRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, vbHighlight, vbHighlight
                    ' If it is selected
                Else
                    .FillRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, .AlphaBlend(vbHighlight, .AlphaBlend(vbWindowBackground, vbButtonFace, 220), 85)
                    .Draw3DRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, vbHighlight, vbHighlight
                End If
            End If
            tT.Top = tT.Top + 2
            tT.Bottom = tT.Top + 9
            tT.Right = tT.Right - 6
            ' If it is pressed
            If m_lPressedBtn = 1 Then
                tT.Right = tT.Right + 1
                tT.Top = tT.Top + 1
                tT.Bottom = tT.Bottom + 1
            End If
            ' Draw arrow
            For i = 4 To 0 Step -1
                .DrawLine tT.Right - i, tT.Top + i, tT.Right - i, tT.Bottom - i, vbWindowText
            Next
            If Not pIsPrevButtonEnabled Then
                ' Draw empty arrows
                For i = 3 To 1 Step -1
                    .DrawLine tT.Right - i, tT.Top + i + 1, tT.Right - i, tT.Bottom - i - 1, vbButtonFace
                Next
            End If
            '== END PREV, NEXT AND CLOSE BUTTONS
            '============================================================
        End With
    Else
        With m_oMemDC
            ' No open MDI child windows
            m_oMemDCTabs.Cls vbApplicationWorkspace
            .Cls vbApplicationWorkspace
            ' Draw LEFT control border
            .DrawLine 0, 0, 0, m_lTabHeight, vbButtonShadow
            .DrawLine 1, 0, 1, m_lTabHeight, lHighlightClr
            .DrawLine 2, 0, 2, m_lTabHeight, vbButtonShadow
            m_oMemDCTabs.DrawLine 0, 0, 0, m_lTabHeight, vbButtonShadow
            ' Draw RIGHT control border
            .DrawLine tR.Right - 1, 0, tR.Right - 1, m_lTabHeight, vbButtonShadow
            .DrawLine tR.Right - 2, 0, tR.Right - 2, m_lTabHeight, lHighlightClr
            .DrawLine tR.Right - 3, 0, tR.Right - 3, m_lTabHeight, vb3DHighlight
        End With
        m_lLastSelMDIChild = 0
    End If

    '============================================================
    ' Transfer image from memory dc into MDI client window dc
    m_oMemDCTabs.BitBlt m_oMemDC.hdc, 2, 2, Abs(tR.Right - tR.Left) - m_lButtonsSize, m_lTabHeight - 2, m_lOffsetX, 2
    m_oMemDC.BitBlt lCHDC, , , , m_lTabHeight

    ReleaseDC lhWnd, lCHDC
End Sub
Private Sub pDrawTabs2000Style(ByVal lhWnd As Long)
    Dim i As Long
    Dim tR As RECT
    Dim tT As RECT
    Dim lPenOld As Long
    Dim lPen As Long
    Dim tPA As POINTAPI
    Dim lCHDC As Long
    Dim hWndChild As Variant
    Dim lTabLeft As Long
    Dim lTabTop As Long
    Dim sCaption As String
    Dim lSelMDIChild As Long
    Dim lSelTabID As Long
    Dim lTextTop As Long
    Dim lBtnOffset As Long
    Dim bIntersect As Long
    
    ' Get MDI client window dc
    lCHDC = GetWindowDC(lhWnd)
    ' Get MDI client dimensions
    GetWindowRect lhWnd, tR
    OffsetRect tR, -tR.Left, -tR.Top
    ' Get active MDI Child handle
    pGetMDIChildWindows
    If m_hWndChild.Count > 0 Then lSelMDIChild = SendMessage(m_lMDIClient, WM_MDIGETACTIVE, 0, 0)

    With m_oMemDC
        ' Initialize memory dc
        .Width = Abs(tR.Right - tR.Left)
        .Height = Abs(tR.Bottom - tR.Top) + 1

        '============================================================
        '== BORDERS
        ' Draw outter border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vb3DHighlight))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Left, m_lTabHeight - 3
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vb3DDKShadow))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Right - 1, m_lTabHeight - 3, tPA
        LineTo lCHDC, tR.Right - 1, tR.Bottom - 1
        LineTo lCHDC, tR.Left - 1, tR.Bottom - 1
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        ' Draw inner border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(m_lClrTabBack1))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 1, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Left + 1, m_lTabHeight - 2
        LineTo lCHDC, tR.Right - 2, m_lTabHeight - 2
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vbButtonShadow))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 1, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Right - 1, tR.Bottom - 2
        MoveToEx lCHDC, tR.Right - 2, m_lTabHeight - 2, tPA
        LineTo lCHDC, tR.Right - 2, tR.Bottom - 2
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        ' Draw inner-inner border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vbButtonShadow))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 2, tR.Bottom - 3, tPA
        LineTo lCHDC, tR.Left + 2, m_lTabHeight - 1
        LineTo lCHDC, tR.Right - 3, m_lTabHeight - 1
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(vb3DHighlight))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 2, tR.Bottom - 3, tPA
        LineTo lCHDC, tR.Right - 2, tR.Bottom - 3
        MoveToEx lCHDC, tR.Right - 3, m_lTabHeight - 1, tPA
        LineTo lCHDC, tR.Right - 3, tR.Bottom - 3
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        '============================================================
        ' Draw bar background
        .FillRect tR.Left, tR.Top, tR.Right, m_lTabHeight - 3, m_lClrBack
        ' Draw bar highlight line
        .DrawLine tR.Left, m_lTabHeight - 3, tR.Right - 1, m_lTabHeight - 3, vb3DHighlight
        ' Fix line end pixel
        .SetPixel tR.Right - 1, m_lTabHeight - 3, .TranslateColor(vb3DDKShadow)
        '== END BORDERS
        '============================================================
    End With

    ' If there are open MDI child windows
    If m_hWndChild.Count > 0 Then
        With m_oMemDCTabs
            ' Initialize memory dc
            .Height = m_lTabHeight
            .Width = 10

            i = 0
            lTabLeft = 2
            ReDim m_tTabR(1 To m_hWndChild.Count) As RECT
            ' For each MDI child window
            For Each hWndChild In m_hWndChild
                i = i + 1
                ' Get MDI child caption
                sCaption = String$(256, vbNullChar)
                GetWindowText hWndChild, sCaption, 256
                sCaption = Left$(sCaption, InStr(1, sCaption, vbNullChar) - 1)
                ' If current tab is selected
                If lSelMDIChild = hWndChild Then
                    m_fFont.Bold = True
                    Set .Font = m_fFont
                    lSelTabID = i
                Else
                    m_fFont.Bold = False
                    Set .Font = m_fFont
                End If
                ' Save tab dimensions
                m_tTabR(i).Left = lTabLeft
                m_tTabR(i).Top = 3
                m_tTabR(i).Bottom = m_lTabHeight - 2 '                  gap=12(6*2)
                m_tTabR(i).Right = m_tTabR(i).Left + .TextWidth(sCaption) + 12
                If m_bDrawIcons Then
                    m_tTabR(i).Right = m_tTabR(i).Right + 20
                End If
                lTabLeft = m_tTabR(i).Right
            Next

            If lSelMDIChild <> m_lLastSelMDIChild And lSelTabID > 0 Then
                RaiseEvent ColorChanged(m_lClrBack)
                m_lLastSelMDIChild = lSelMDIChild
                ' Ensure that a newly selected tab is scrolled into view 49=buttonbar size
                If m_tTabR(lSelTabID).Left - m_lOffsetX < tR.Left Then
                    m_lOffsetX = m_tTabR(lSelTabID).Left - 30
                ElseIf m_tTabR(lSelTabID).Right - m_lOffsetX > tR.Right - 49 Then
                    m_lOffsetX = m_lOffsetX + ((m_tTabR(lSelTabID).Right - m_lOffsetX) - (tR.Right - 49)) + 30
                    If m_tTabR(m_hWndChild.Count).Right - m_lOffsetX = (tR.Right - 49) - 30 Then
                        m_lOffsetX = m_lOffsetX - 30
                    End If
                End If
                If m_lOffsetX <= 30 Then m_lOffsetX = 0
            End If

            ' Initialize memory dc
            .Width = m_tTabR(m_hWndChild.Count).Right + 10
            .Height = m_lTabHeight
            .Cls m_lClrBack
            ' Draw highlight line
            .DrawLine 0, m_lTabHeight - 3, .Width, m_lTabHeight - 3, vb3DHighlight
            i = 0
            ' For each MDI child window
            For Each hWndChild In m_hWndChild
                i = i + 1
                
                ' See if rects intersect (if tab is visible)
                m_tTabR(i).Left = m_tTabR(i).Left - m_lOffsetX
                m_tTabR(i).Right = m_tTabR(i).Right - m_lOffsetX
                bIntersect = IntersectRect(tT, tR, m_tTabR(i))
                m_tTabR(i).Left = m_tTabR(i).Left + m_lOffsetX
                m_tTabR(i).Right = m_tTabR(i).Right + m_lOffsetX
                
                ' If it is visible then draw it
                If bIntersect Then
                    ' Get MDI child caption
                    sCaption = String$(256, vbNullChar)
                    GetWindowText hWndChild, sCaption, 256
                    sCaption = Left$(sCaption, InStr(1, sCaption, vbNullChar) - 1)
                    ' If current tab is selected
                    If lSelMDIChild = hWndChild Then
                        ' Set text color to default (black) color
                        .ForeColor = .TranslateColor(vbWindowText)
                        m_fFont.Bold = True
                        Set .Font = m_fFont
    
                        m_tTabR(i).Left = m_tTabR(i).Left - 2
                        m_tTabR(i).Right = m_tTabR(i).Right + 2
    
                        ' Draw tab background
                        .FillRect m_tTabR(i).Left, m_tTabR(i).Top, m_tTabR(i).Right, m_tTabR(i).Bottom, m_lClrTabBack1 'vbButtonFace
                        .DrawLine m_tTabR(i).Left + 1, m_tTabR(i).Top - 1, m_tTabR(i).Right - 1, m_tTabR(i).Top - 1, m_lClrTabBack1 'vbButtonFace
                        ' Draw tab border
                        .DrawLine m_tTabR(i).Left, m_tTabR(i).Top, m_tTabR(i).Left, m_tTabR(i).Bottom, vb3DHighlight
                        .SetPixel m_tTabR(i).Left + 1, m_tTabR(i).Top - 1, .TranslateColor(vb3DHighlight)
                        .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top - 2, m_tTabR(i).Right - 2, m_tTabR(i).Top - 2, vb3DHighlight
                        .SetPixel m_tTabR(i).Right - 2, m_tTabR(i).Top - 1, .TranslateColor(vb3DDKShadow)
                        .DrawLine m_tTabR(i).Right - 1, m_tTabR(i).Top, m_tTabR(i).Right - 1, m_tTabR(i).Bottom - 1, vb3DDKShadow
                        .DrawLine m_tTabR(i).Right - 2, m_tTabR(i).Top, m_tTabR(i).Right - 2, m_tTabR(i).Bottom - 1, vbButtonShadow
                        ' Fix highlight line
                        .DrawLine m_tTabR(i).Right - 2, m_tTabR(i).Bottom - 1, m_tTabR(i).Right, m_tTabR(i).Bottom - 1, vb3DHighlight
    
                        lTextTop = -1
                        
                        ' Should we draw focus rect?
                        If m_bDrawFocusRect Then
                            LSet tT = m_tTabR(i)
                            With tT
                                .Left = .Left + 4
                                .Top = .Top + 2
                                .Right = .Right - 5
                                .Bottom = .Bottom - 3
                            End With
                            ' Draw focus rect
                            DrawFocusRectAPI .hdc, tT
                        End If
                    Else
                        ' Set text color to lighter black color
                        .ForeColor = .TranslateColor(vbWindowText)
                        m_fFont.Bold = False
                        Set .Font = m_fFont
    
                        If lSelTabID = i - 1 Then
                            ' Draw tab background
                            .FillRect m_tTabR(i).Left + 2, m_tTabR(i).Top + 2, m_tTabR(i).Right, m_tTabR(i).Bottom - 1, vbButtonFace
                            .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top + 1, m_tTabR(i).Right - 1, m_tTabR(i).Top + 1, vbButtonFace
                            ' Draw tab border
                            .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top, m_tTabR(i).Right - 2, m_tTabR(i).Top, vb3DHighlight
                            .SetPixel m_tTabR(i).Right - 2, m_tTabR(i).Top + 1, .TranslateColor(vb3DDKShadow)
                            .DrawLine m_tTabR(i).Right - 1, m_tTabR(i).Top + 2, m_tTabR(i).Right - 1, m_tTabR(i).Bottom - 1, vb3DDKShadow
                            .DrawLine m_tTabR(i).Right - 2, m_tTabR(i).Top + 2, m_tTabR(i).Right - 2, m_tTabR(i).Bottom - 1, vbButtonShadow
                        Else
                            ' Draw tab background
                            .FillRect m_tTabR(i).Left, m_tTabR(i).Top + 2, m_tTabR(i).Right, m_tTabR(i).Bottom - 1, vbButtonFace
                            .DrawLine m_tTabR(i).Left + 1, m_tTabR(i).Top + 1, m_tTabR(i).Right - 1, m_tTabR(i).Top + 1, vbButtonFace
                            ' Draw tab border
                            .DrawLine m_tTabR(i).Left, m_tTabR(i).Top + 2, m_tTabR(i).Left, m_tTabR(i).Bottom, vb3DHighlight
                            .SetPixel m_tTabR(i).Left + 1, m_tTabR(i).Top + 1, .TranslateColor(vb3DHighlight)
                            .DrawLine m_tTabR(i).Left + 2, m_tTabR(i).Top, m_tTabR(i).Right - 2, m_tTabR(i).Top, vb3DHighlight
                            .SetPixel m_tTabR(i).Right - 2, m_tTabR(i).Top + 1, .TranslateColor(vb3DDKShadow)
                            .DrawLine m_tTabR(i).Right - 1, m_tTabR(i).Top + 2, m_tTabR(i).Right - 1, m_tTabR(i).Bottom - 1, vb3DDKShadow
                            .DrawLine m_tTabR(i).Right - 2, m_tTabR(i).Top + 2, m_tTabR(i).Right - 2, m_tTabR(i).Bottom - 1, vbButtonShadow
                        End If
    
                        lTextTop = 0
                    End If
                    ' Draw tab caption
                    lTabTop = m_tTabR(i).Top + (((m_tTabR(i).Bottom - m_tTabR(i).Top) / 2) - (.TextHeight(sCaption) / 2))
                    If m_bDrawIcons Then
                        .DrawText sCaption, m_tTabR(i).Left + 20, lTabTop + lTextTop, m_tTabR(i).Right, m_tTabR(i).Bottom, DT_CENTER
                        .DrawPicture pGetSmallIcon(CLng(hWndChild)), m_tTabR(i).Left + 5, m_tTabR(i).Top + 2
                    Else
                        .DrawText sCaption, m_tTabR(i).Left, lTabTop + lTextTop, m_tTabR(i).Right, m_tTabR(i).Bottom, DT_CENTER
                    End If
                End If
            Next
        End With
        
        With m_oMemDC
            '============================================================
            '== PREV, NEXT AND CLOSE BUTTONS
            lBtnOffset = 2
            ' Draw background for buttons
            LSet tT = tR
            tT.Left = tT.Right - 47
            tT.Top = tR.Top + 1
            tT.Bottom = m_lTabHeight - 2
            .FillRect tT.Left, tT.Top, tT.Right, tT.Bottom, m_lClrBack
            .DrawLine tT.Left, tT.Bottom - 1, tT.Right, tT.Bottom - 1, vb3DHighlight
            m_lButtonsSize = 7 - lBtnOffset
            ' Draw close button
            m_tCloseBtnRect.Left = tT.Right - (19 - lBtnOffset)
            m_tCloseBtnRect.Top = 5
            m_tCloseBtnRect.Right = m_tCloseBtnRect.Left + 14
            m_tCloseBtnRect.Bottom = m_tCloseBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            ' Draw X sign using Marlett font
            Dim fFont As StdFont
            Set fFont = New StdFont
            fFont.Name = "Marlett"
            fFont.Size = 7
            Set .Font = fFont
            .ForeColor = m_lClrButton
            ' Draw button border
            If m_lSelBtn = 3 Then
                ' If it is pressed
                If m_lPressedBtn = 3 Then
                    .Draw3DRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, m_lClrButtonBorderDK, m_lClrButtonBorder
                    ' If it is selected
                Else
                    .Draw3DRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, m_lClrButtonBorder, m_lClrButtonBorderDK
                End If
            End If
            ' If it is pressed
            If m_lPressedBtn = 3 Then
                ' Draw X sign
                .DrawText "r", m_tCloseBtnRect.Left + 4, m_tCloseBtnRect.Top + 4, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom
            Else
                ' Draw X sign
                .DrawText "r", m_tCloseBtnRect.Left + 3, m_tCloseBtnRect.Top + 3, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom
            End If
            '========================================
            ' Draw NEXT button
            m_tNextBtnRect.Left = m_tCloseBtnRect.Left - 14
            m_tNextBtnRect.Top = 5
            m_tNextBtnRect.Right = m_tNextBtnRect.Left + 14
            m_tNextBtnRect.Bottom = m_tNextBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            LSet tT = m_tNextBtnRect
            ' Draw button border
            If m_lSelBtn = 2 Then
                ' If it is pressed
                If m_lPressedBtn = 2 Then
                    .Draw3DRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, m_lClrButtonBorderDK, m_lClrButtonBorder
                    ' If it is selected
                Else
                    .Draw3DRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, m_lClrButtonBorder, m_lClrButtonBorderDK
                End If
            End If
            tT.Left = tT.Left + 4
            tT.Top = tT.Top + 2
            tT.Bottom = tT.Top + 9
            ' If it is pressed
            If m_lPressedBtn = 2 Then
                tT.Left = tT.Left + 1
                tT.Top = tT.Top + 1
                tT.Bottom = tT.Bottom + 1
            End If
            ' Draw arrow
            For i = 0 To 4
                .DrawLine tT.Left + i, tT.Top + i, tT.Left + i, tT.Bottom - i, m_lClrButton
            Next
            If Not pIsNextButtonEnabled Then
                ' Draw empty arrows
                For i = 1 To 3
                    .DrawLine tT.Left + i, tT.Top + i + 1, tT.Left + i, tT.Bottom - i - 1, m_lClrBack
                Next
            End If
            '========================================
            ' Draw PREV button
            m_tPrevBtnRect.Left = m_tNextBtnRect.Left - 14
            m_tPrevBtnRect.Top = 5
            m_tPrevBtnRect.Right = m_tPrevBtnRect.Left + 14
            m_tPrevBtnRect.Bottom = m_tPrevBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            LSet tT = m_tPrevBtnRect
            ' Draw button border
            If m_lSelBtn = 1 Then
                ' If it is pressed
                If m_lPressedBtn = 1 Then
                    .Draw3DRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, m_lClrButtonBorderDK, m_lClrButtonBorder
                    ' If it is selected
                Else
                    .Draw3DRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, m_lClrButtonBorder, m_lClrButtonBorderDK
                End If
            End If
            tT.Top = tT.Top + 2
            tT.Bottom = tT.Top + 9
            tT.Right = tT.Right - 6
            ' If it is pressed
            If m_lPressedBtn = 1 Then
                tT.Right = tT.Right + 1
                tT.Top = tT.Top + 1
                tT.Bottom = tT.Bottom + 1
            End If
            ' Draw arrow
            For i = 4 To 0 Step -1
                .DrawLine tT.Right - i, tT.Top + i, tT.Right - i, tT.Bottom - i, m_lClrButton
            Next
            If Not pIsPrevButtonEnabled Then
                ' Draw empty arrows
                For i = 3 To 1 Step -1
                    .DrawLine tT.Right - i, tT.Top + i + 1, tT.Right - i, tT.Bottom - i - 1, m_lClrBack
                Next
            End If
            '== END PREV, NEXT AND CLOSE BUTTONS
            '============================================================
        End With
        
        '============================================================
        ' Transfer image from memory dc into MDI client window dc
        m_oMemDCTabs.BitBlt m_oMemDC.hdc, , , Abs(tR.Right - tR.Left) - m_lButtonsSize, m_lTabHeight - 2, m_lOffsetX
        m_oMemDC.BitBlt lCHDC, , , , m_lTabHeight - 2
    Else
        With m_oMemDC
            ' No open MDI child windows
            m_oMemDCTabs.Cls vbApplicationWorkspace
            .Cls vbApplicationWorkspace
            ' Draw LEFT control border
            .DrawLine 0, 0, 0, m_lTabHeight, vb3DHighlight
            .DrawLine 1, 0, 1, m_lTabHeight, m_lClrTabBack1
            .DrawLine 2, 0, 2, m_lTabHeight, vbButtonShadow
            ' Draw RIGHT control border
            .DrawLine tR.Right - 1, 0, tR.Right - 1, m_lTabHeight, vb3DDKShadow
            .DrawLine tR.Right - 2, 0, tR.Right - 2, m_lTabHeight, vbButtonShadow
            .DrawLine tR.Right - 3, 0, tR.Right - 3, m_lTabHeight, vb3DHighlight
        End With
        
        m_lLastSelMDIChild = 0
        
        '============================================================
        ' Transfer image from memory dc into MDI client window dc
        m_oMemDC.BitBlt lCHDC, , , , m_lTabHeight
    End If
    
    ReleaseDC lhWnd, lCHDC
End Sub
Private Sub pDrawTabsXPStyle(ByVal lhWnd As Long)
    Dim i As Long
    Dim tR As RECT
    Dim tT As RECT
    Dim lPenOld As Long
    Dim lPen As Long
    Dim tPA As POINTAPI
    Dim lCHDC As Long
    Dim hWndChild As Variant
    Dim lTabLeft As Long
    Dim lTabTop As Long
    Dim sCaption As String
    Dim lSelMDIChild As Long
    Dim lSelTabID As Long
    Dim lBtnOffset As Long
    Dim bIntersect As Long
    
    ' Get MDI client window dc
    lCHDC = GetWindowDC(lhWnd)
    ' Get MDI client dimensions
    GetWindowRect lhWnd, tR
    OffsetRect tR, -tR.Left, -tR.Top
    ' Get active MDI Child handle
    pGetMDIChildWindows
    If m_hWndChild.Count > 0 Then lSelMDIChild = SendMessage(m_lMDIClient, WM_MDIGETACTIVE, 0, 0)
    
    With m_oMemDC
        ' Initialize memory dc
        .Width = Abs(tR.Right - tR.Left)
        .Height = Abs(tR.Bottom - tR.Top) + 1

        '============================================================
        '== BORDERS
        ' Draw outter border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(m_lClrOuterBorder))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left, tR.Top, tPA
        LineTo lCHDC, tR.Right - 1, tR.Top
        LineTo lCHDC, tR.Right - 1, tR.Bottom - 1
        LineTo lCHDC, tR.Left, tR.Bottom - 1
        LineTo lCHDC, tR.Left, tR.Top
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        ' Draw inner border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(m_lClrBorder))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 1, tR.Top + 1, tPA
        LineTo lCHDC, tR.Left + 1, tR.Bottom - 2
        MoveToEx lCHDC, tR.Right - 2, tR.Top + 1, tPA
        LineTo lCHDC, tR.Right - 2, tR.Bottom - 2
        MoveToEx lCHDC, tR.Left + 1, tR.Bottom - 2, tPA
        LineTo lCHDC, tR.Right - 1, tR.Bottom - 2
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        ' Draw inner-inner border line
        lPen = CreatePen(PS_SOLID, 1, .TranslateColor(m_lClrInnerBorder))
        lPenOld = SelectObject(lCHDC, lPen)
        MoveToEx lCHDC, tR.Left + 2, m_lTabHeight, tPA
        LineTo lCHDC, tR.Left + 2, tR.Bottom - 3
        MoveToEx lCHDC, tR.Right - 3, m_lTabHeight, tPA
        LineTo lCHDC, tR.Right - 3, tR.Bottom - 3
        MoveToEx lCHDC, tR.Left + 2, tR.Bottom - 3, tPA
        LineTo lCHDC, tR.Right - 2, tR.Bottom - 3
        SelectObject lCHDC, lPenOld
        DeleteObject lPen
        '========================================
        ' Draw left border
        .DrawLine tR.Left, tR.Top, tR.Left, m_lTabHeight, m_lClrOuterBorder
        .DrawLine tR.Left + 1, tR.Top + 1, tR.Left + 1, m_lTabHeight, m_lClrBorder
        ' Draw top border
        .DrawLine tR.Left, tR.Top, tR.Right, tR.Top, m_lClrOuterBorder
        ' Draw right border
        .DrawLine tR.Right - 1, tR.Top, tR.Right - 1, m_lTabHeight, m_lClrOuterBorder
        .DrawLine tR.Right - 2, tR.Top + 1, tR.Right - 2, m_lTabHeight, m_lClrBorder
        '============================================================
        ' Draw bar background
        .FillRect tR.Left + 2, tR.Top + 1, tR.Right - 2, m_lTabHeight - 3, m_lClrBack
        ' Draw bar highlight line
        .DrawLine tR.Left + 2, m_lTabHeight - 3, tR.Right - 2, m_lTabHeight - 3, m_lClrTabBorder
        .DrawLine tR.Left + 2, m_lTabHeight - 1, tR.Right - 2, m_lTabHeight - 1, m_lClrInnerBorder
        .DrawLine tR.Left + 2, m_lTabHeight - 2, tR.Right - 2, m_lTabHeight - 2, m_lClrBorder
        '== END BORDERS
        '============================================================
    End With
    
    ' If there are open MDI child windows
    If m_hWndChild.Count > 0 Then
        With m_oMemDCTabs
            ' Initialize memory dc
            .Height = m_lTabHeight
            .Width = 10
            
            i = 0
            lTabLeft = 4
            ReDim m_tTabR(1 To m_hWndChild.Count) As RECT
            ' For each MDI child window
            For Each hWndChild In m_hWndChild
                i = i + 1
                ' Get MDI child caption
                sCaption = String$(256, vbNullChar)
                GetWindowText hWndChild, sCaption, 256
                sCaption = Left$(sCaption, InStr(1, sCaption, vbNullChar) - 1)
                ' If current tab is selected
                If lSelMDIChild = hWndChild Then
                    m_fFont.Bold = True
                    Set .Font = m_fFont
                    lSelTabID = i
                Else
                    m_fFont.Bold = False
                    Set .Font = m_fFont
                End If
                ' Save tab dimensions
                m_tTabR(i).Left = lTabLeft
                m_tTabR(i).Top = 3
                m_tTabR(i).Bottom = m_lTabHeight - 2 '                  gap=12(6*2)
                m_tTabR(i).Right = m_tTabR(i).Left + .TextWidth(sCaption) + 12
                If m_bDrawIcons Then
                    m_tTabR(i).Right = m_tTabR(i).Right + 20
                End If
                lTabLeft = m_tTabR(i).Right
            Next
            
            If lSelMDIChild <> m_lLastSelMDIChild And lSelTabID > 0 Then
                RaiseEvent ColorChanged(m_lClrBack)
                m_lLastSelMDIChild = lSelMDIChild
                ' Ensure that a newly selected tab is scrolled into view 49=buttonbar size
                If m_tTabR(lSelTabID).Left - m_lOffsetX < tR.Left Then
                    m_lOffsetX = m_tTabR(lSelTabID).Left - 30
                ElseIf m_tTabR(lSelTabID).Right - m_lOffsetX > tR.Right - 49 Then
                    m_lOffsetX = m_lOffsetX + ((m_tTabR(lSelTabID).Right - m_lOffsetX) - (tR.Right - 49)) + 30
                    If m_tTabR(m_hWndChild.Count).Right - m_lOffsetX = (tR.Right - 49) - 30 Then
                        m_lOffsetX = m_lOffsetX - 30
                    End If
                End If
                If m_lOffsetX <= 30 Then m_lOffsetX = 0
            End If
                
            ' Initialize memory dc
            .Width = m_tTabR(m_hWndChild.Count).Right + 10
            .Height = m_lTabHeight
            .Cls m_lClrBack
            ' Draw highlight line
            .DrawLine 0, m_lTabHeight - 3, .Width, m_lTabHeight - 3, m_lClrTabBorder
            i = 0
            ' For each MDI child window
            For Each hWndChild In m_hWndChild
                i = i + 1
                
                ' See if rects intersect (if tab is visible)
                m_tTabR(i).Left = m_tTabR(i).Left - m_lOffsetX
                m_tTabR(i).Right = m_tTabR(i).Right - m_lOffsetX
                bIntersect = IntersectRect(tT, tR, m_tTabR(i))
                m_tTabR(i).Left = m_tTabR(i).Left + m_lOffsetX
                m_tTabR(i).Right = m_tTabR(i).Right + m_lOffsetX
                
                ' If it is visible then draw it
                If bIntersect Then
                    ' Get MDI child caption
                    sCaption = String$(256, vbNullChar)
                    GetWindowText hWndChild, sCaption, 256
                    sCaption = Left$(sCaption, InStr(1, sCaption, vbNullChar) - 1)
                    ' If current tab is selected
                    If lSelMDIChild = hWndChild Then
                        ' Set text color to default (black) color
                        .ForeColor = .TranslateColor(m_lClrTabFore)
                        m_fFont.Bold = True
                        Set .Font = m_fFont
    
                        ' Draw tab background
                        .FillRect m_tTabR(i).Left, m_tTabR(i).Top, m_tTabR(i).Right, m_tTabR(i).Bottom, m_lClrTabBack1
                        ' Draw tab border
                        .DrawLine m_tTabR(i).Left, m_tTabR(i).Top, m_tTabR(i).Left, m_tTabR(i).Bottom, m_lClrTabBorder
                        .DrawLine m_tTabR(i).Left, m_tTabR(i).Top, m_tTabR(i).Right, m_tTabR(i).Top, m_lClrTabBorder
                        .DrawLine m_tTabR(i).Right - 1, m_tTabR(i).Top + 1, m_tTabR(i).Right - 1, m_tTabR(i).Bottom, m_lClrTabBorderDK
                        
                        ' Should we draw focus rect?
                        If m_bDrawFocusRect Then
                            LSet tT = m_tTabR(i)
                            With tT
                                .Left = .Left + 3
                                .Top = .Top + 3
                                .Right = .Right - 3
                                .Bottom = .Bottom - 2
                            End With
                            ' Draw focus rect
                            DrawFocusRectAPI .hdc, tT
                        End If
                    Else
                        ' Set text color to lighter black color
                        .ForeColor = m_lClrTabInactiveFore
                        m_fFont.Bold = False
                        Set .Font = m_fFont
    
                        .DrawLine m_tTabR(i).Right, m_tTabR(i).Top + 2, m_tTabR(i).Right, m_tTabR(i).Bottom - 2, m_lClrTabSeparator
                    End If
                    ' Draw tab caption
                    lTabTop = m_tTabR(i).Top + (((m_tTabR(i).Bottom - m_tTabR(i).Top) / 2) - (.TextHeight(sCaption) / 2))
                    If m_bDrawIcons Then
                        .DrawText sCaption, m_tTabR(i).Left + 20, lTabTop, m_tTabR(i).Right, m_tTabR(i).Bottom, DT_CENTER
                        .DrawPicture pGetSmallIcon(CLng(hWndChild)), m_tTabR(i).Left + 5, m_tTabR(i).Top + 2
                    Else
                        .DrawText sCaption, m_tTabR(i).Left, lTabTop, m_tTabR(i).Right, m_tTabR(i).Bottom, DT_CENTER
                    End If
                End If
            Next
        End With
        
        With m_oMemDC
            '============================================================
            '== PREV, NEXT AND CLOSE BUTTONS
            ' Draw background for buttons
            LSet tT = tR
            tT.Left = tT.Right - 47
            tT.Top = tR.Top + 1
            tT.Bottom = m_lTabHeight - 2
            .FillRect tT.Left, tT.Top, tT.Right, tT.Bottom, m_lClrBack
            .DrawLine tT.Right - 1, tT.Top, tT.Right - 1, tT.Bottom, m_lClrOuterBorder
            .DrawLine tT.Left, tT.Bottom - 1, tT.Right - 2, tT.Bottom - 1, m_lClrTabBorder
            .FillRect tT.Right - 2, tT.Top, tT.Right - 1, tT.Bottom, m_lClrBorder
            m_lButtonsSize = 7 - lBtnOffset
            ' Draw close button
            m_tCloseBtnRect.Left = tT.Right - (19 - lBtnOffset)
            m_tCloseBtnRect.Top = 5
            m_tCloseBtnRect.Right = m_tCloseBtnRect.Left + 14
            m_tCloseBtnRect.Bottom = m_tCloseBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            ' Draw X sign using Marlett font
            Dim fFont As StdFont
            Set fFont = New StdFont
            fFont.Name = "Marlett"
            fFont.Size = 7
            Set .Font = fFont
            .ForeColor = m_lClrButton
            ' Draw button border
            If m_lSelBtn = 3 Then
                ' If it is pressed
                If m_lPressedBtn = 3 Then
                    .Draw3DRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, m_lClrButtonBorderDK, m_lClrButtonBorder
                    ' If it is selected
                Else
                    .Draw3DRect m_tCloseBtnRect.Left, m_tCloseBtnRect.Top, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom, m_lClrButtonBorder, m_lClrButtonBorderDK
                End If
            End If
            ' If it is pressed
            If m_lPressedBtn = 3 Then
                ' Draw X sign
                .DrawText "r", m_tCloseBtnRect.Left + 4, m_tCloseBtnRect.Top + 4, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom
            Else
                ' Draw X sign
                .DrawText "r", m_tCloseBtnRect.Left + 3, m_tCloseBtnRect.Top + 3, m_tCloseBtnRect.Right, m_tCloseBtnRect.Bottom
            End If
            '========================================
            ' Draw NEXT button
            m_tNextBtnRect.Left = m_tCloseBtnRect.Left - 14
            m_tNextBtnRect.Top = 5
            m_tNextBtnRect.Right = m_tNextBtnRect.Left + 14
            m_tNextBtnRect.Bottom = m_tNextBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            LSet tT = m_tNextBtnRect
            ' Draw button border
            If m_lSelBtn = 2 Then
                ' If it is pressed
                If m_lPressedBtn = 2 Then
                    .Draw3DRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, m_lClrButtonBorderDK, m_lClrButtonBorder
                    ' If it is selected
                Else
                    .Draw3DRect m_tNextBtnRect.Left, m_tNextBtnRect.Top, m_tNextBtnRect.Right, m_tNextBtnRect.Bottom, m_lClrButtonBorder, m_lClrButtonBorderDK
                End If
            End If
            tT.Left = tT.Left + 4
            tT.Top = tT.Top + 2
            tT.Bottom = tT.Top + 9
            ' If it is pressed
            If m_lPressedBtn = 2 Then
                tT.Left = tT.Left + 1
                tT.Top = tT.Top + 1
                tT.Bottom = tT.Bottom + 1
            End If
            ' Draw arrow
            For i = 0 To 4
                .DrawLine tT.Left + i, tT.Top + i, tT.Left + i, tT.Bottom - i, m_lClrButton
            Next
            If Not pIsNextButtonEnabled Then
                ' Draw empty arrows
                For i = 1 To 3
                    .DrawLine tT.Left + i, tT.Top + i + 1, tT.Left + i, tT.Bottom - i - 1, m_lClrBack
                Next
            End If
            '========================================
            ' Draw PREV button
            m_tPrevBtnRect.Left = m_tNextBtnRect.Left - 14
            m_tPrevBtnRect.Top = 5
            m_tPrevBtnRect.Right = m_tPrevBtnRect.Left + 14
            m_tPrevBtnRect.Bottom = m_tPrevBtnRect.Top + 15
            m_lButtonsSize = m_lButtonsSize + 14
            LSet tT = m_tPrevBtnRect
            ' Draw button border
            If m_lSelBtn = 1 Then
                ' If it is pressed
                If m_lPressedBtn = 1 Then
                    .Draw3DRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, m_lClrButtonBorderDK, m_lClrButtonBorder
                    ' If it is selected
                Else
                    .Draw3DRect m_tPrevBtnRect.Left, m_tPrevBtnRect.Top, m_tPrevBtnRect.Right, m_tPrevBtnRect.Bottom, m_lClrButtonBorder, m_lClrButtonBorderDK
                End If
            End If
            tT.Top = tT.Top + 2
            tT.Bottom = tT.Top + 9
            tT.Right = tT.Right - 6
            ' If it is pressed
            If m_lPressedBtn = 1 Then
                tT.Right = tT.Right + 1
                tT.Top = tT.Top + 1
                tT.Bottom = tT.Bottom + 1
            End If
            ' Draw arrow
            For i = 4 To 0 Step -1
                .DrawLine tT.Right - i, tT.Top + i, tT.Right - i, tT.Bottom - i, m_lClrButton
            Next
            If Not pIsPrevButtonEnabled Then
                ' Draw empty arrows
                For i = 3 To 1 Step -1
                    .DrawLine tT.Right - i, tT.Top + i + 1, tT.Right - i, tT.Bottom - i - 1, m_lClrBack
                Next
            End If
            '== END PREV, NEXT AND CLOSE BUTTONS
            '============================================================
        End With
    Else
        With m_oMemDC
            ' No open MDI child windows
            m_oMemDCTabs.Cls vbApplicationWorkspace
            .Cls vbApplicationWorkspace
            ' Draw LEFT control border
            .DrawLine 0, 0, 0, m_lTabHeight, m_lClrOuterBorder
            .DrawLine 1, 0, 1, m_lTabHeight, m_lClrBorder
            .DrawLine 2, 0, 2, m_lTabHeight, m_lClrInnerBorder
            m_oMemDCTabs.DrawLine 0, 0, 0, m_lTabHeight, m_lClrInnerBorder
            ' Draw RIGHT control border
            .DrawLine tR.Right - 1, 0, tR.Right - 1, m_lTabHeight, m_lClrOuterBorder
            .DrawLine tR.Right - 2, 0, tR.Right - 2, m_lTabHeight, m_lClrBorder
            .DrawLine tR.Right - 3, 0, tR.Right - 3, m_lTabHeight, m_lClrInnerBorder
        End With
        m_lLastSelMDIChild = 0
    End If
    
    '============================================================
    ' Transfer image from memory dc into MDI client window dc
    m_oMemDCTabs.BitBlt m_oMemDC.hdc, 2, 2, Abs(tR.Right - tR.Left) - m_lButtonsSize, m_lTabHeight - 4, m_lOffsetX, 2
    m_oMemDC.BitBlt lCHDC, , , , m_lTabHeight
    
    ReleaseDC lhWnd, lCHDC
End Sub
'********************************************************************
'* Name: pGetMDIChildWindows
'* Description: Get all open MDI child windows and put them into collection.
'********************************************************************
Private Sub pGetMDIChildWindows()
    Dim i As Long
    Dim lR As Long
    Dim hWndNow As Variant
    
    ' create new temporary collection
    Set m_hWndTempChild = New Collection
    ' Enumerate all MDI child windows
    lR = EnumChildWindows(m_lMDIClient, AddressOf pEnumChildWindowProc, ObjPtr(Me))
    ' For each MDI child window
    For Each hWndNow In m_hWndTempChild
        ' If it is a new window then add it to child windows collection
        If Not pKeyExists(m_hWndChild, "H" & hWndNow) Then
            m_hWndChild.Add hWndNow, "H" & hWndNow
            
            m_lClrIndex = m_lClrIndex + 1
            If m_lClrIndex > 8 Then m_lClrIndex = 1
            m_oColorIndex.Add m_lClrIndex, "H" & hWndNow
        End If
    Next
    ' For each MDI window in child window collection
    For i = m_hWndChild.Count To 1 Step -1
        ' If current window is closed then remove it from collection
        If Not pKeyExists(m_hWndTempChild, "H" & m_hWndChild(i)) Then
            m_hWndChild.Remove i
            m_oColorIndex.Remove i
        End If
    Next
End Sub
'********************************************************************
'* Name: pIsWithinTab
'* Description: Return selected tab or button based on mouse coordinates.
'********************************************************************
Private Function pIsWithinTab(ByRef bInButton As Long) As Long
    Dim i As Long
    Dim tPA As POINTAPI
    
    ' Get cursor position
    GetCursorPos tPA
    ScreenToClient m_lMDIClient, tPA
    ' Compensate for nonclient area size
    tPA.y = tPA.y + m_lTabHeight
    tPA.x = tPA.x + 2  '2=border
    
    ' First check to see if it is within close,next or prev button
    If tPA.x >= m_tCloseBtnRect.Left And tPA.x <= m_tCloseBtnRect.Right And _
       tPA.y >= m_tCloseBtnRect.Top And tPA.y <= m_tCloseBtnRect.Bottom Then bInButton = 3: pIsWithinTab = -1: Exit Function
    If tPA.x >= m_tNextBtnRect.Left And tPA.x <= m_tNextBtnRect.Right And tPA.y >= m_tNextBtnRect.Top And tPA.y <= m_tNextBtnRect.Bottom Then
        If pIsNextButtonEnabled Then bInButton = 2
        pIsWithinTab = -1
        Exit Function
    End If
    If tPA.x >= m_tPrevBtnRect.Left And tPA.x <= m_tPrevBtnRect.Right And tPA.y >= m_tPrevBtnRect.Top And tPA.y <= m_tPrevBtnRect.Bottom Then
        If pIsPrevButtonEnabled Then bInButton = 1
        pIsWithinTab = -1
        Exit Function
    End If
    bInButton = 0
    
    tPA.x = tPA.x + m_lOffsetX
    For i = 1 To m_hWndChild.Count
        With m_tTabR(i)
            If tPA.x >= .Left And tPA.x <= .Right And tPA.y >= .Top And tPA.y <= .Bottom Then
                pIsWithinTab = i
                Exit For
            End If
        End With
    Next
End Function
'********************************************************************
'* Name: pIsPrevButtonEnabled
'* Description: Return true if we can scroll to left.
'********************************************************************
Private Function pIsPrevButtonEnabled() As Boolean
    pIsPrevButtonEnabled = (m_lOffsetX > 0)
End Function
'********************************************************************
'* Name: pIsNextButtonEnabled
'* Description: Return true if we can scroll to right.
'********************************************************************
Private Function pIsNextButtonEnabled() As Boolean
    If m_hWndChild.Count > 0 Then
        pIsNextButtonEnabled = ((m_tTabR(m_hWndChild.Count).Right - m_lOffsetX + 1) > m_tPrevBtnRect.Left)
    End If
End Function
'********************************************************************
'* Name: pReplaceTab
'* Description: Switch places of two provided tabs.
'********************************************************************
Private Sub pReplaceTab(ByVal lDragging As Long, ByVal lCandidate As Long)
    Dim i As Long
    Dim tCol As New Collection
    Dim tColClr As New Collection
    
    ' If we are replacing tab before the dragging one
    If lCandidate < lDragging Then
        ' Add all tabs to temporary collection which are before the replace candidate
        For i = 1 To lCandidate - 1
            If i <> lDragging Then
                tCol.Add m_hWndChild(i), "H" & m_hWndChild(i)
                tColClr.Add m_oColorIndex(i), "H" & m_hWndChild(i)
            End If
        Next
        ' Add dragging tab to replace candidates place
        tCol.Add m_hWndChild(lDragging), "H" & m_hWndChild(lDragging)
        tColClr.Add m_oColorIndex(lDragging), "H" & m_hWndChild(lDragging)
        ' Save dragging tab ID
        m_lDraggingTab = tCol.Count
        ' Add all tabs to temporary collection which are after the replace candidate
        For i = lCandidate To m_hWndChild.Count
            If i <> lDragging Then
                tCol.Add m_hWndChild(i), "H" & m_hWndChild(i)
                tColClr.Add m_oColorIndex(i), "H" & m_hWndChild(i)
            End If
        Next
    Else
        ' Add all tabs to temporary collection until the replace candidate
        For i = 1 To lCandidate
            If i <> lDragging Then
                tCol.Add m_hWndChild(i), "H" & m_hWndChild(i)
                tColClr.Add m_oColorIndex(i), "H" & m_hWndChild(i)
            End If
        Next
        ' Add the dragged one
        tCol.Add m_hWndChild(lDragging), "H" & m_hWndChild(lDragging)
        tColClr.Add m_oColorIndex(lDragging), "H" & m_hWndChild(lDragging)
        ' Save dragging tab ID
        m_lDraggingTab = tCol.Count
        ' Add all tabs to temporary collection which are after the replace candidate
        For i = lCandidate + 1 To m_hWndChild.Count
            If i <> lDragging Then
                tCol.Add m_hWndChild(i), "H" & m_hWndChild(i)
                tColClr.Add m_oColorIndex(i), "H" & m_hWndChild(i)
            End If
        Next
    End If
    ' Replace MDI child collection with temporary one
    Set m_hWndChild = tCol
    Set m_oColorIndex = tColClr
    ' Redraw control
    pDrawControl m_lMDIClient
    m_bJustReplaced = True
    ' Save cursor coordinates
    GetCursorPos m_tJustReplacedPoint
End Sub
Private Sub pScrollPrev()
    m_lOffsetX = m_lOffsetX - 33
    If m_lOffsetX < 0 Then m_lOffsetX = 0
    ' Redraw control
    pDrawControl m_lMDIClient
End Sub
Private Sub pScrollNext()
    m_lOffsetX = m_lOffsetX + 33
    pFixOffset
    If m_lOffsetX < 0 Then m_lOffsetX = 0
    ' Redraw control
    pDrawControl m_lMDIClient
End Sub
Private Sub pFixOffset()
    Dim lMaxRight As Long
    Dim lSize As Long
    Dim tR As RECT
    
    ' If there are open MDI child forms
    If m_hWndChild.Count > 0 Then
        ' Get dimensions of MDI client window
        GetWindowRect m_lMDIClient, tR
        ' Set MaxRight to length of all tab items together
        lMaxRight = m_tTabR(m_hWndChild.Count).Right
        ' set size to MDI client window width
        lSize = tR.Right - tR.Left
        ' Decrease size by button width
        lSize = lSize - 49
        ' If MaxRight is bigger then mdi client window width
        If (lMaxRight > lSize) Then
            ' If current view is to small to display tab control then
            If (lMaxRight - m_lOffsetX < lSize) Then m_lOffsetX = lMaxRight - lSize
            ' If everything can fit into mdi client window then set offsetx to 0
        ElseIf (lSize > lMaxRight) Then
            m_lOffsetX = 0
        End If
    End If
End Sub
'********************************************************************
'* Name: pSetColors
'* Description: Setup default colors for current theme.
'********************************************************************
Private Sub pSetColors()
    If m_oMemDC Is Nothing Then Exit Sub

    Dim lMax As Long, lAmount As Long
    Dim lR As Long, lG As Long, lB As Long
            
    ' Get theme name
    pGetWindowThemeName
    
    With m_oMemDC
        If m_eStyle = mtsOffice2000 Then
            m_lClrTabSeparator = vbBlack
            m_lClrButton = vbWindowText
            m_lClrButtonBorder = vb3DHighlight
            m_lClrButtonBorderDK = vbButtonShadow
            m_lClrBack = vbButtonFace
            m_lClrTabBack1 = vbButtonFace
            
        ElseIf m_eStyle = mtsOffice2003 Then
            m_lClrTabSeparator = vbBlack
            
        ElseIf m_eStyle = mtsOfficeXP Then
            lR = .GetRGB(.TranslateColor(vbButtonFace), 1)
            lG = .GetRGB(.TranslateColor(vbButtonFace), 2)
            lB = .GetRGB(.TranslateColor(vbButtonFace), 3)
            If lR > lG Then
                If lR > lB Then
                    lMax = lR
                Else
                    lMax = lB
                End If
            Else
                If lG > lB Then
                    lMax = lG
                Else
                    lMax = lB
                End If
            End If
            If lMax = 0 Then
                m_lClrBack = RGB(35, 35, 35)
            Else
                If 255 - lMax > 35 Then
                    lAmount = lMax + 35
                Else
                    lAmount = (255 - lMax) + lMax
                End If
                m_lClrBack = RGB(lR * lAmount / lMax, lG * lAmount / lMax, lB * lAmount / lMax)
            End If
            m_lClrTabBack1 = vbButtonFace
            m_lClrTabFore = vbWindowText
            m_lClrTabInactiveFore = .AlphaBlend(vbWindowText, vbButtonShadow, 100)
            m_lClrTabBorder = vb3DHighlight
            m_lClrTabBorderDK = vbWindowText
            m_lClrOuterBorder = vbButtonShadow
            m_lClrInnerBorder = vbButtonShadow
            m_lClrBorder = vbButtonFace
            m_lClrButton = vb3DDKShadow
            m_lClrButtonBorder = &H80000016
            m_lClrButtonBorderDK = vbWindowText
            m_lClrTabSeparator = vbButtonShadow
        End If
    End With
End Sub
'********************************************************************
'* Name: pGetSmallIcon
'* Description: Returns specified window icon.
'********************************************************************
Private Function pGetSmallIcon(lHandle As Long) As StdPicture
    Dim hIcon As Long
    
    ' Try to get small icon
    hIcon = SendMessage(lHandle, WM_GETICON, ICON_SMALL, ByVal 0&)
    If hIcon = 0 Then
        ' Try to get big icon
        hIcon = SendMessage(lHandle, WM_GETICON, ICON_BIG, ByVal 0&)
        If hIcon <> 0 Then
            ' Try to load small icon from big one
            hIcon = CopyImage(hIcon, IMAGE_ICON, 16, 16, LR_COPYFROMRESOURCE)
            If hIcon <> 0 Then
                ' Convert icon to picture
                Set pGetSmallIcon = m_oMemDC.IconToPicture(hIcon)
            Else
                ' Get default icon
                Set pGetSmallIcon = picDefault.Picture
            End If
        Else
            ' Get default icon
            Set pGetSmallIcon = picDefault.Picture
        End If
    Else
        ' Convert icon to picture
        Set pGetSmallIcon = m_oMemDC.IconToPicture(hIcon)
    End If
End Function
'********************************************************************
'* Name: pGetWindowThemeName
'* Description: Returns window visual style theme name.
'* NOTE: Only valid for WinXP and above.
'********************************************************************
Private Function pGetWindowThemeName() As Long
    Dim lTheme As Long
    Dim lpOS As Long
    Dim lPtrColorName As Long
    Dim lPtrThemeFile As Long
    Dim tOVI As OSVERSIONINFO
    
    ' Check if windows is XP
    tOVI.dwOSVersionInfoSize = Len(tOVI)
    GetVersionEx tOVI
    ' Default theme name
    m_sThemeName = "Classic"
    ' If it is Windows Xp or higher
    If tOVI.dwMajorVersion > 5 And tOVI.dwMinorVersion >= 1 Then
        ' Try to open theme
        lTheme = OpenThemeData(UserControl.hWnd, StrPtr("ExplorerBar"))
        ' Success
        If Not lTheme = 0 Then
            ' Initialize theme filename and theme name variables
            ReDim bThemeFile(0 To 520) As Byte
            lPtrThemeFile = VarPtr(bThemeFile(0))
            ReDim bColorName(0 To 520) As Byte
            lPtrColorName = VarPtr(bColorName(0))
            ' Get theme name
            lpOS = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
            
            ' Return theme name (NormalColor,HomeStead,Metallic)
            m_sThemeName = bColorName
            lpOS = InStr(m_sThemeName, vbNullChar)
            If lpOS > 1 Then m_sThemeName = Left$(m_sThemeName, lpOS - 1)
            
            ' Cleanup
            CloseThemeData lTheme
        End If
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////
'//// HELPER FUNCTIONS
'//////////////////////////////////////////////////////////////////////////////
'********************************************************************
'* Name: pKeyExists
'* Description: Return true if specified item exists in collection.
'********************************************************************
Private Function pKeyExists(ByVal c As Collection, ByVal Key As String) As Boolean
    On Error Resume Next
    Dim oItem As Variant
    oItem = c(Key)
    If Err.Number = 0 Then pKeyExists = True
End Function

'//////////////////////////////////////////////////////////////////////////////
'//// FRIEND FUNCTIONS
'//////////////////////////////////////////////////////////////////////////////
'********************************************************************
'* Name: fAddMDIChildWindow
'* Description: Add new MDI child window into temporary collection.
'********************************************************************
Friend Sub fAddMDIChildWindow(ByVal hWnd As Long)
    m_hWndTempChild.Add hWnd, "H" & hWnd
End Sub

Public Sub RedrawControl()
  pDrawControl m_lMDIClient
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//// SUBCLASS EVENTS
'//////////////////////////////////////////////////////////////////////////////
Private Sub ISubclassingSink_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim lMsg As Long, lBtn As Long, lTab As Long
    Dim lReplaceCandidate As Long
    Dim tP As POINTAPI
    Dim i As Long
    Dim iMouseBtn As Integer
    Dim lColor As OLE_COLOR
    Dim tR As RECT
    
    Select Case uMsg
        Case &HAE, &H7F '7f=WM_GETICON (works in XpSp2, AE doesn't)
            pDrawControl m_lMDIClient
        Case WM_MDIACTIVATE
            lReturn = m_oSubclass.CallOrigWndProc(uMsg, wParam, lParam)
            pDrawControl m_lMDIClient
        Case WM_NCPAINT
            lReturn = m_oSubclass.CallOrigWndProc(uMsg, wParam, lParam)
            pDrawControl m_lMDIClient
        Case WM_SETCURSOR
            lReturn = m_oSubclass.CallOrigWndProc(uMsg, wParam, lParam)
            ' If mouse pointer is over nonclient area
            If (lParam And &HFFFF&) = HTNOWHERE Then
                lMsg = (lParam And &H7FFF0000) \ &H10000
                Select Case lMsg
                    Case WM_MOUSEMOVE
                        ' Get selected tab or button
                        lTab = pIsWithinTab(lBtn)
                        ' If selected button is changed ERROR ERROR ERROR
                        If lBtn <> m_lSelBtn Then
                            ' If haven't captured window yet, capture it
                            If m_lSelBtn = 0 Then
                                SetCapture m_lMDIClient
                            End If
                            ' Redraw control
                            m_lSelBtn = lBtn
                            pDrawControl m_lMDIClient
                        End If
                    Case WM_LBUTTONDOWN
                        ' Get selected tab or button
                        lTab = pIsWithinTab(lBtn)
                        ' If we clciked on tab
                        If lBtn > 0 Then
                            ' If haven't captured window yet, capture it
                            If m_lSelBtn = 0 Then
                                SetCapture m_lMDIClient
                            End If
                            m_lPressedBtn = lBtn
                            m_lSelBtn = lBtn
                            ' Redraw control
                            pDrawControl m_lMDIClient
                        Else
                            If lTab > 0 Then
                                ' Save dragging tab ID
                                m_lDraggingTab = lTab
                                m_bJustReplaced = True
                                ' Save cursor position
                                GetCursorPos m_tJustReplacedPoint
                                ' Capture mdi client window messages
                                SetCapture m_lMDIClient
                                ' Added by NR Start
                                SendMessage m_lMDIClient, WM_SETREDRAW, 0, 0
                                ' Added by NR End
                                ' If active child form is not the dragged form(tab)
                                If SendMessageLong(m_lMDIClient, WM_MDIGETACTIVE, 0, 0) <> m_hWndChild(m_lDraggingTab) Then
                                    SendMessageLong m_lMDIClient, WM_MDIACTIVATE, m_hWndChild(m_lDraggingTab), 0
                                End If
                                ' Added by NR Start
                                SendMessage m_lMDIClient, WM_SETREDRAW, 1, 0
                                RedrawWindow m_lMDIClient, 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN
                                SetFocus m_lMDIClient
                                ' Added by NR End
                                ' Redraw control
                                pDrawControl m_lMDIClient
                            End If
                        End If
                    Case WM_LBUTTONUP, WM_RBUTTONUP
                        ' If we are not dragging tab
                        If m_lDraggingTab = 0 Then
                            ' See which button is pressed
                            If lMsg = WM_LBUTTONUP Then
                                iMouseBtn = vbLeftButton
                            ElseIf lMsg = WM_RBUTTONUP Then
                                iMouseBtn = vbRightButton
                            End If
                            GetCursorPos tP
                            ScreenToClient m_lMDIClient, tP
                            If iMouseBtn = vbRightButton Then
                                lTab = pIsWithinTab(lBtn)
                                If lTab > 0 Then
                                    If SendMessageLong(m_lMDIClient, WM_MDIGETACTIVE, 0, 0) <> m_hWndChild(lTab) Then
                                        SendMessageLong m_lMDIClient, WM_MDIACTIVATE, m_hWndChild(lTab), 0
                                    End If
                                    RaiseEvent TabClick(m_hWndChild(lTab), iMouseBtn, tP.x, tP.y)
                                    Select Case m_eStyle
                                        Case mtsOfficeXP, mtsOffice2000
                                            lColor = vbButtonFace
                                        Case mtsOffice2003
                                            lColor = m_lColorTable(m_oColorIndex(lTab))
                                    End Select
                                    RaiseEvent ColorChanged(lColor)
                                Else
                                    RaiseEvent TabBarClick(iMouseBtn, tP.x, tP.y)
                                End If
                            Else
                                RaiseEvent TabBarClick(iMouseBtn, tP.x, tP.y)
                            End If
                        End If
                End Select
            End If
        Case WM_MOUSEMOVE
            If m_lDraggingTab > 0 Then
                ' Get selected tab or button
                lTab = pIsWithinTab(lBtn)
                ' get cursor position
                GetCursorPos tP
                ' If we just replaced tab
                If m_bJustReplaced Then
                    ' If dragging tab is not the currently active tab
                    If m_lDraggingTab <> lTab Then
                        ' If we are dragging tab outside of control
                        If Abs(tP.x - m_tJustReplacedPoint.x) > (m_tTabR(m_lDraggingTab).Right - m_tTabR(m_lDraggingTab).Left) / 2 Then
                            m_bJustReplaced = False
                        Else
                            Exit Sub
                        End If
                    Else
                        m_bJustReplaced = False
                    End If
                End If
                ' Convert screen mouse coordinates to client coordinates
                ScreenToClient m_lMDIClient, tP
                ' Increase left mouse coordinate by offset value
                tP.x = tP.x + m_lOffsetX
                tP.y = tP.y + m_lTabHeight
                ' If we are inside of tab
                If (tP.y > m_tTabR(1).Top) And (tP.y < m_tTabR(1).Bottom) Then
                    ' If we are in front of first tab then replace first tab
                    If tP.x < m_tTabR(1).Left And tP.x > m_tTabR(1).Left - 3 Then
                        lReplaceCandidate = 1
                        ' If we are behind the last tab then replace last tab
                    ElseIf tP.x > m_tTabR(m_hWndChild.Count).Right And tP.x < m_tTabR(m_hWndChild.Count).Right - 3 Then
                        lReplaceCandidate = m_hWndChild.Count
                    Else
                        ' Check over which tab are we and replace active tab
                        For i = 1 To m_hWndChild.Count
                            ' We are inside the tab then replace it
                            If (tP.x >= m_tTabR(i).Left) And (tP.x <= m_tTabR(i).Right) Then
                                lReplaceCandidate = i
                                Exit For
                            End If
                        Next
                    End If
                    ' If we got tab to replace
                    If lReplaceCandidate > 0 Then
                        ' If this is not the same tab as dragged one then replace it
                        If lReplaceCandidate <> m_lDraggingTab Then
                            pReplaceTab m_lDraggingTab, lReplaceCandidate
                        End If
                    End If
                End If
            End If

            ' If we got selected button
            If m_lSelBtn > 0 Then
                ' Get selected tab or button
                lTab = pIsWithinTab(lBtn)
                ' If selected button is changed
                If lBtn <> m_lSelBtn Then
                    If lBtn = 0 Then
                        If m_lPressedBtn = 0 Then
                            ReleaseCapture
                        End If
                    End If
                    ' Redraw control
                    m_lSelBtn = lBtn
                    pDrawControl m_lMDIClient
                End If
            End If
        Case WM_LBUTTONDOWN
            ' Get selected tab or button
            lTab = pIsWithinTab(lBtn)
            If lBtn > 0 Then
                SetCapture m_lMDIClient
                m_lPressedBtn = lBtn
                m_lSelBtn = lBtn
                Select Case lBtn
                    Case 1
                        ' Scroll left
                        If pIsPrevButtonEnabled Then
                            pScrollPrev
                            m_oTimer.Enabled = True
                        End If
                    Case 2
                        ' Scroll right
                        If pIsNextButtonEnabled Then
                            pScrollNext
                            m_oTimer.Enabled = True
                        End If
                End Select
                pDrawControl m_lMDIClient
            End If
        Case WM_LBUTTONUP, WM_RBUTTONUP
            m_oTimer.Enabled = False
            ' If we are dragging tab
            If m_lDraggingTab > 0 Then
                ' Get selected tab or button
                lTab = pIsWithinTab(lBtn)
                ' Activate dragged tab form
                If SendMessageLong(m_lMDIClient, WM_MDIGETACTIVE, 0, 0) <> m_hWndChild(m_lDraggingTab) Then
                    SendMessageLong m_lMDIClient, WM_MDIACTIVATE, m_hWndChild(m_lDraggingTab), 0
                End If
                ReleaseCapture
                ' See which button is pressed
                If uMsg = WM_LBUTTONUP Then
                    iMouseBtn = vbLeftButton
                ElseIf uMsg = WM_RBUTTONUP Then
                    iMouseBtn = vbRightButton
                End If
                If lTab > 0 Then
                    GetCursorPos tP
                    ScreenToClient m_lMDIClient, tP
                    RaiseEvent TabClick(m_hWndChild(lTab), iMouseBtn, tP.x, tP.y)
                    Select Case m_eStyle
                        Case mtsOfficeXP, mtsOffice2000
                            lColor = vbButtonFace
                        Case mtsOffice2003
                            lColor = m_lColorTable(m_oColorIndex(lTab))
                    End Select
                    RaiseEvent ColorChanged(lColor)
                End If
                m_lDraggingTab = 0
            End If
            ' If user pressed button
            If m_lPressedBtn > 0 Then
                ReleaseCapture
                m_lPressedBtn = 0
                m_lSelBtn = 0
                ' Get selected tab or button
                lTab = pIsWithinTab(lBtn)
                ' If user clicked on close button then close active form
                If lBtn = 3 Then
                    ' Added by NR Start
                    'If UBound(m_tTabR()) > 1 Then
                    '    SendMessage m_lMDIClient, WM_SETREDRAW, 0, 0
                    'End If
                    ' Added by NR End
                    SendMessage SendMessage(m_lMDIClient, WM_MDIGETACTIVE, 0, 0), WM_CLOSE, 0, 0
                    ' Added by NR Start
                    If UBound(m_tTabR()) > 1 Then
                        SendMessage m_lMDIClient, WM_SETREDRAW, 1, 0
                    End If
                    ' Added by NR End
                End If
                ReleaseCapture
                pDrawControl m_lMDIClient
            End If
            '============================================
            '= Detect system color change message so that
            '= we can adjust colors for selected menu style.
            '============================================
        Case WM_SYSCOLORCHANGE
            pSetColors
    End Select
End Sub
Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    Dim tNCR As NCCALCSIZE_PARAMS
    Dim tWP As WINDOWPOS
    
    Select Case uMsg
        Case WM_NCCALCSIZE
            If wParam <> 0 Then
                ' Increase nonclient area size so that we have room
                ' for drawing the tabs.
                CopyMemory tNCR, ByVal lParam, Len(tNCR)
                CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)
                With tNCR.rgrc(0)
                    .Left = tWP.x + 3
                    .Top = tWP.y + m_lTabHeight
                    .Right = tWP.x + tWP.cX - 3
                    .Bottom = tWP.y + tWP.cY - 3
                End With
                LSet tNCR.rgrc(1) = tNCR.rgrc(0)
                CopyMemory ByVal lParam, tNCR, Len(tNCR)
                bHandled = True
                lReturn = WVR_VALIDRECTS
            End If
        Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED
            pDrawControl m_lMDIClient
        Case WM_ACTIVATE
            m_bHasFocus = (wParam And &HFFFF&)
            pDrawControl m_lMDIClient
    End Select
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//// TIMER EVENTS
'//////////////////////////////////////////////////////////////////////////////
Private Sub m_oTimer_Timer()
    If m_lSelBtn = 1 And pIsPrevButtonEnabled Then
        pScrollPrev
        pDrawControl m_lMDIClient
    ElseIf m_lSelBtn = 2 And pIsNextButtonEnabled Then
        pScrollNext
        pDrawControl m_lMDIClient
    End If
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//// USERCONTROL EVENTS
'//////////////////////////////////////////////////////////////////////////////
Private Sub UserControl_Initialize()
    m_lTabHeight = 25
    m_bDrawIcons = True
    
    m_lColorTable(1) = RGB(138, 168, 228)
    m_lColorTable(2) = RGB(255, 219, 117)
    m_lColorTable(3) = RGB(189, 205, 159)
    m_lColorTable(4) = RGB(240, 158, 159)
    m_lColorTable(5) = RGB(186, 166, 225)
    m_lColorTable(6) = RGB(154, 191, 180)
    m_lColorTable(7) = RGB(247, 182, 131)
    m_lColorTable(8) = RGB(216, 171, 192)
    
    pSetColors
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim lStyle As Long
    
    With PropBag
        m_eStyle = .ReadProperty("Style", mtsOffice2003)
        Set m_fFont = .ReadProperty("Font", UserControl.Font)
        m_bDrawFocusRect = .ReadProperty("DrawFocusRect", False)
        m_bDrawIcons = .ReadProperty("DrawIcons", True)
    End With
    
    ' If we are in run-time mode
    If UserControl.Ambient.UserMode Then
        ' Create a new instance of object
        Set m_oMemDC = New CMemoryDC
        Set m_oMemDCTabs = New CMemoryDC
        Set m_oSubclass = New CEasySubclass_v1
        Set m_oParentSubclass = New CEasySubclass_v1
        Set m_oTimer = New CTimer
        m_oTimer.Interval = 50
        m_oTimer.Enabled = False
        Set m_hWndChild = New Collection
        Set m_oColorIndex = New Collection
        ' Find MDI client window handle
        m_lMDIClient = FindWindowEx(UserControl.Parent.hWnd, 0, "MDIClient", vbNullString)
        If m_lMDIClient <> 0 Then
            ' Change MDI client window border style
            lStyle = GetWindowLong(m_lMDIClient, GWL_EXSTYLE)
            lStyle = lStyle And Not WS_EX_CLIENTEDGE
            SetWindowLong m_lMDIClient, GWL_EXSTYLE, lStyle
            ' Refresh window
            SetWindowPos m_lMDIClient, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
            ' Subclass MDI client window
            With m_oSubclass
                .Subclass m_lMDIClient, Me, False, True
                .AddBeforeMsgs WM_NCCALCSIZE, WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED
                .AddAfterMsgs WM_MDIACTIVATE, WM_SETCURSOR, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONUP, WM_SYSCOLORCHANGE
            End With
            With m_oParentSubclass
                .Subclass GetParent(m_lMDIClient), Me, False, True
                .AddAfterMsgs WM_NCPAINT, &HAE, &H7F
                .AddBeforeMsgs WM_ACTIVATE
            End With
        End If
    End If
       
    pSetColors
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Style", m_eStyle, mtsOffice2003
        .WriteProperty "Font", m_fFont, UserControl.Font
        .WriteProperty "DrawFocusRect", m_bDrawFocusRect, False
        .WriteProperty "DrawIcons", m_bDrawIcons, True
    End With
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Width = 32 * Screen.TwipsPerPixelX
    UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub
Private Sub UserControl_Terminate()
    On Error Resume Next
    Dim lStyle As Long
    If m_lMDIClient <> 0 Then
        ' Unsubclass window
        m_oSubclass.UnSubclass
        m_oParentSubclass.UnSubclass
        Set m_oSubclass = Nothing
        Set m_oParentSubclass = Nothing
        Set m_oMemDC = Nothing
        Set m_oMemDCTabs = Nothing
        Set m_oTimer = Nothing
        ' Restore window style
        lStyle = GetWindowLong(m_lMDIClient, GWL_EXSTYLE)
        lStyle = m_lMDIClient Or WS_EX_CLIENTEDGE
        SetWindowLong m_lMDIClient, GWL_EXSTYLE, lStyle
    End If
End Sub
