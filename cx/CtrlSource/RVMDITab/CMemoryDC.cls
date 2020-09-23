VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'*    Author      : Andrea Batina[Revelatek]
'*    Date        : 24 February 2004
'*
'*    Component   : CMemoryDC
'*    Description : Memory device contex drawing class used for flicker-free
'*                  drawing.
'*
'*    Dependencies: None.
'*
'*    Credits     : This class contains some functions which I found on internet.
'*                  If you see your function in here please notify me so that I can
'*                  put your name on it.
'*
'*    Copyright   : Copyright © 2004 Revelatek. All rights reserved.
'*
'*    History     :
'*                  24.02.2004 - Started creating new class from scratch.
'*
'*******************************************************************************

Option Explicit

' Base error number constant
Private Const ERRBASE = 5000

'////////////////////////////////////////////////////////////////////
'// Private/Public Type Definitions
Private Type POINTAPI
    x                   As Long
    y                   As Long
End Type
'Private Type RECT
'    Left                As Long
'    Top                 As Long
'    Right               As Long
'    Bottom              As Long
'End Type
Private Type DRAWTEXTPARAMS
    cbSize              As Long
    iTabLength          As Long
    iLeftMargin         As Long
    iRightMargin        As Long
    uiLengthDrawn       As Long
End Type
Private Type TEXTMETRIC
    tmHeight            As Long
    tmAscent            As Long
    tmDescent           As Long
    tmInternalLeading   As Long
    tmExternalLeading   As Long
    tmAveCharWidth      As Long
    tmMaxCharWidth      As Long
    tmWeight            As Long
    tmOverhang          As Long
    tmDigitizedAspectX  As Long
    tmDigitizedAspectY  As Long
    tmFirstChar         As Byte
    tmLastChar          As Byte
    tmDefaultChar       As Byte
    tmBreakChar         As Byte
    tmItalic            As Byte
    tmUnderlined        As Byte
    tmStruckOut         As Byte
    tmPitchAndFamily    As Byte
    tmCharSet           As Byte
End Type
Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(1 To 32) As Byte
End Type
Private Type SIZEAPI
    cX                  As Long
    cY                  As Long
End Type
Private Type PointSng   ' Internal Point structure
    x                   As Single   ' Uses Singles for more precision.
    y                   As Single
End Type
Private Type RGBType
    r                   As Byte
    g                   As Byte
    b                   As Byte
    A                   As Byte
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type
Private Type PictDesc
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type

'////////////////////////////////////////////////////////////////////
'// Private/Public Enum Definitions
Public Enum EMemDCDrawText
    DT_LEFT = &H0               ' Aligns text to the left.
    DT_TOP = &H0                ' Justifies the text to the top of the rectangle.
    DT_CENTER = &H1             ' Centers text horizontally in the rectangle.
    DT_RIGHT = &H2              ' Aligns text to the right.
    DT_VCENTER = &H4            ' Centers text vertically. This value is used only with the DT_SINGLELINE value.
    DT_BOTTOM = &H8             ' Justifies the text to the bottom of the rectangle. This value is used only with the DT_SINGLELINE value.
    DT_WORDBREAK = &H10         ' Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-line feed sequence also breaks the line.<br>If this is not specified, output is on one line.
    DT_SINGLELINE = &H20        ' Displays text on a single line only. Carriage returns and line feeds do not break the line.
    DT_EXPANDTABS = &H40        ' Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value.
    DT_TABSTOP = &H80           ' Sets tab stops. Bits 15–8 (high-order byte of the low-order word) of the uFormat parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value.
    DT_NOCLIP = &H100           ' Draws without clipping. DrawText is somewhat faster when DT_NOCLIP is used.
    DT_EXTERNALLEADING = &H200  ' Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.
    DT_CALCRECT = &H400         ' Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawText returns the height of the formatted text but does not draw the text.
    DT_NOPREFIX = &H800         ' Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off
    DT_INTERNAL = &H1000        ' Uses the system font to calculate text metrics.
    DT_EDITCONTROL = &H2000&    ' Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line.
    DT_PATH_ELLIPSIS = &H4000&  ' For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_END_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_END_ELLIPSIS = &H8000&   ' For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_MODIFYSTRING = &H10000   ' Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
    DT_RTLREADING = &H20000     ' Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right.
    DT_WORD_ELLIPSIS = &H40000  ' Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

'////////////////////////////////////////////////////////////////////
'// Private/Public Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function APIFillRect Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function APIBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hdcDest As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function ApiDrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Private Declare Function APIGetPixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function APISetPixel Lib "gdi32" Alias "SetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'////////////////////////////////////////////////////////////////////
'// Private/Public Constant Declarations
' Pen Style Messages
Private Const PS_SOLID = 0
Private Const PS_DOT = 2

' Draw State Messages
Private Const DST_BITMAP = &H4
Private Const DST_COMPLEX = &H0
Private Const DST_ICON = &H3
Private Const DSS_NORMAL = &H0
Private Const DSS_MONO = &H80

' Misc Messages
Private Const LOGPIXELSY = 90
Private Const FW_NORMAL = 400
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180

'////////////////////////////////////////////////////////////////////
'// Private/Public Variable Declarations
Private m_lhDC              As Long     ' Memory dc handle
Private m_lWidth            As Long     ' Memory dc width
Private m_lHeight           As Long     ' Memory dc height
Private m_hBmp              As Long     ' Memory dc new compatible bitmap
Private m_hBmpOld           As Long     ' Memory dc old compatible bitmap
Private m_lOrginalFont      As Long     ' Memory dc original font
Private m_lMemoryFont       As Long     ' Memory dc new font

Public Function BlendColors(ByVal lClrFirst As Long, ByVal lClrSecond As Long, ByVal lAmount As Currency)
    Dim lRed As Long, lGreen As Long, lBlue As Long
    
    lRed = Int(GetRGB(lClrFirst, 1) * lAmount + GetRGB(lClrSecond, 1)) * (255 - lAmount) / 255
    lGreen = Int(GetRGB(lClrFirst, 2) * lAmount + GetRGB(lClrSecond, 2)) * (255 - lAmount) / 255
    lBlue = Int(GetRGB(lClrFirst, 3) * lAmount + GetRGB(lClrSecond, 3)) * (255 - lAmount) / 255
    
    If lRed > 255 Then lRed = 255
    If lGreen > 255 Then lGreen = 255
    If lBlue > 255 Then lBlue = 255
    
    BlendColors = RGB(lRed, lGreen, lBlue)
End Function

'********************************************************************
'* Name: LightenColor
'* Description: Lightens provided color by specified amount.
'********************************************************************
Public Function LightenColor(ByVal lClrLight As Long, ByVal lClrDark As Long, ByVal lAmount As Long) As Long
    Dim lVal As Long
    Dim lRed As Long, lGreen As Long, lBlue As Long
    
    If lAmount < 100 Then
        lVal = 100
    Else
        lVal = 1000
    End If
    
    lRed = (GetRGB(lClrDark, 1) * (lVal - lAmount) + lVal / 2 + GetRGB(lClrLight, 1) * lAmount) / lVal
    lGreen = (GetRGB(lClrDark, 2) * (lVal - lAmount) + lVal / 2 + GetRGB(lClrLight, 2) * lAmount) / lVal
    lBlue = (GetRGB(lClrDark, 3) * (lVal - lAmount) + lVal / 2 + GetRGB(lClrLight, 3) * lAmount) / lVal
    
    LightenColor = RGB(lRed, lGreen, lBlue)
End Function

'********************************************************************
'* Name: hDC
'* Description: Returns memory dc handle.
'********************************************************************
Public Property Get hdc() As Long
    hdc = m_lhDC
End Property

'********************************************************************
'* Name: Width
'* Description: Get\Set memory dc width.
'********************************************************************
Public Property Get Width() As Long
    Width = m_lWidth
End Property
Public Property Let Width(ByVal lW As Long)
    If lW > m_lWidth Then pCreateMemDC lW, m_lHeight
End Property

'********************************************************************
'* Name: Height
'* Description: Get\Set memory dc height.
'********************************************************************
Public Property Get Height() As Long
    Height = m_lHeight
End Property
Public Property Let Height(ByVal lH As Long)
    If lH > m_lHeight Then pCreateMemDC m_lWidth, lH
End Property

'********************************************************************
'* Name: pCreateMemDC
'* Description: Create memory dc.
'********************************************************************
Private Sub pCreateMemDC(ByVal lW As Long, ByVal lH As Long)
    On Error GoTo PROC_ERR_pCreateMemDC

        Dim lhDC As Long
    
        ' Destroy memory dc
100     pDestroyMemDC
    
        ' Create a new dc based on desktop
102     lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Create compatible dc with desktop one
104     m_lhDC = CreateCompatibleDC(lhDC)
        ' Create a new bitmap on which we will draw
106     m_hBmp = CreateCompatibleBitmap(lhDC, lW, lH)
        ' Select bitmap into new dc
108     m_hBmpOld = SelectObject(m_lhDC, m_hBmp)
        ' If there was an error
110     If m_hBmpOld = 0 Then
112         pDestroyMemDC
        Else
            ' Save dc width and height
114         m_lWidth = lW
116         m_lHeight = lH
        End If
        ' Set background mode to transparent
118     SetBkMode m_lhDC, 1
        ' Delete temp dc
120     DeleteDC lhDC

PROC_EXIT:
        Exit Sub
PROC_ERR_pCreateMemDC:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pCreateMemDC", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: pDestroyMemDC
'* Description: Delete all GDI object to free memory and to prevent
'* memory leaks.
'********************************************************************
Private Sub pDestroyMemDC()
    On Error GoTo PROC_ERR_pDestroyMemDC

100     If m_hBmpOld <> 0 Then
102         SelectObject m_lhDC, m_hBmpOld
104         m_hBmpOld = 0
        End If
106     If m_hBmp <> 0 Then
108         DeleteObject m_hBmp
110         m_hBmp = 0
        End If
112     If m_lOrginalFont <> 0 Then
114         Call SelectObject(m_lhDC, m_lOrginalFont)
116         m_lOrginalFont = 0
        End If
118     If m_lMemoryFont <> 0 Then
120         Call DeleteObject(m_lMemoryFont)
122         m_lMemoryFont = 0
        End If
124     m_lWidth = 0
126     m_lHeight = 0
128     If m_lhDC <> 0 Then
130         DeleteDC m_lhDC
132         m_lhDC = 0
        End If

PROC_EXIT:
        Exit Sub
PROC_ERR_pDestroyMemDC:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pDestroyMemDC", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: IsCreated
'* Description: Returns true if we already created the memory dc.
'********************************************************************
Public Function IsCreated() As Boolean
    IsCreated = m_lhDC <> 0
End Function

'********************************************************************
'* Name: Cls
'* Description: Clear memory dc.
'********************************************************************
Public Sub Cls(Optional ByVal lColor As Long = -1)
    FillRect 0, 0, Width, Height, lColor
End Sub

'********************************************************************
'* Name: BitBlt
'* Description: Transfers image from memory dc to hdcDest dc.
'********************************************************************
Public Sub BitBlt(ByVal hdcDest As Long, _
       Optional ByVal xDest As Long, _
       Optional ByVal yDest As Long, _
       Optional ByVal lWidth As Long, _
       Optional ByVal lHeight As Long, _
       Optional ByVal xSrc As Long, _
       Optional ByVal ySrc As Long, _
       Optional ByVal dwRop As RasterOpConstants = vbSrcCopy, _
       Optional ByVal lhdcSrc As Long = 0)

    On Error GoTo PROC_ERR_BitBlt

100     If Not IsCreated Then Exit Sub
    
        ' Setup properties
102     If lWidth = 0 Then lWidth = m_lWidth
104     If lHeight = 0 Then lHeight = m_lHeight
        ' BitBlt
106     If lhdcSrc <> 0 Then
108         APIBitBlt hdcDest, xDest, yDest, lWidth, lHeight, lhdcSrc, xSrc, ySrc, dwRop
        Else
110         APIBitBlt hdcDest, xDest, yDest, lWidth, lHeight, m_lhDC, xSrc, ySrc, dwRop
        End If

PROC_EXIT:
        Exit Sub
PROC_ERR_BitBlt:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.BitBlt", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: Class_Terminate
'* Description: Class object termination.
'********************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    pDestroyMemDC
End Sub

'********************************************************************
'* Name: TranslateColor
'* Description: Translates color from OLE_COLOR type to system color.
'********************************************************************
Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    On Error Resume Next
    
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function

'********************************************************************
'* Name: FillRect
'* Description: Fill rect with specified color.
'********************************************************************
Public Sub FillRect( _
       Optional ByVal lLeft As Long, _
       Optional ByVal lTop As Long, _
       Optional ByVal lRight As Long = -1, _
       Optional ByVal lBottom As Long = -1, _
       Optional ByVal lColor As Long = -1, _
       Optional ByVal lhDC As Long)
    
    On Error GoTo PROC_ERR_FillRect
    
100     If Not IsCreated Then Exit Sub
    
        Dim rc As RECT
        Dim hbrFill As Long
    
        ' Create solid brush with specified color
102     If lColor <> -1 Then hbrFill = CreateSolidBrush(TranslateColor(lColor))
        ' Setup rectangle dimensions
104     With rc
106         .Left = lLeft
108         .Top = lTop
110         .Right = IIf(lRight < lLeft, Width, lRight)
112         .Bottom = IIf(lBottom < lTop, Height, lBottom)
        End With
            ' Fill rect
        If lhDC <> 0 Then
            APIFillRect lhDC, rc, hbrFill
        Else
114         APIFillRect m_lhDC, rc, hbrFill
        End If
        ' Delete brush object
116     If lColor <> -1 Then DeleteObject hbrFill

PROC_EXIT:
        Exit Sub
PROC_ERR_FillRect:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.FillRect", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: DrawLine
'* Description: Draw line on memory dc.
'********************************************************************
Public Sub DrawLine(lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, lColor As OLE_COLOR, Optional lWidth As Long = 1)
    On Error GoTo PROC_ERR_DrawLine
        
        If Not IsCreated() Then Exit Sub
        
        Dim LrPos As POINTAPI
        Dim LhPen As Long
        Dim LhOldPen As Long
    
        ' If user didn't specified pen width then set it to 1
100     If lWidth = 0 Then
102         LhPen = CreatePen(PS_DOT, 1, TranslateColor(lColor))
        ' Else create pen with specified width
        Else
104         LhPen = CreatePen(PS_SOLID, lWidth, TranslateColor(lColor))
        End If
        ' Put new pen on dc and save the old one
106     If LhPen Then LhOldPen = SelectObject(m_lhDC, LhPen)

        ' Set first line point(X1,Y2)
108     MoveToEx m_lhDC, lX1, lY1, LrPos
        ' Draw line to second point (X2,Y2)
110     LineTo m_lhDC, lX2, lY2
        ' Restore old pen and delete the new one
112     If LhPen Then
114         Call SelectObject(m_lhDC, LhOldPen)
116         Call DeleteObject(LhPen)
        End If

PROC_EXIT:
        Exit Sub
PROC_ERR_DrawLine:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.DrawLine", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: GetRGB
'* Description: Extract Red, Green or Blue color from RGB color.
'********************************************************************
Public Function GetRGB(lRGB As Long, lNum As Integer) As Integer
    On Error GoTo PROC_ERR_GetRGB

100     If lNum > 0 And lNum < 4 And lRGB > -1 And lRGB < 16777216 Then
102         GetRGB = lRGB \ 256 ^ (lNum - 1) And 255
        End If

PROC_EXIT:
        Exit Function
PROC_ERR_GetRGB:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.GetRGB", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Function

'********************************************************************
'* Name: DrawPicture
'* Description: Draw specified picture on memory dc.
'********************************************************************
Public Sub DrawPicture(ByRef m_PictureF As StdPicture, _
       ByVal x As Long, _
       ByVal y As Long, _
       Optional lW As Long = 16, _
       Optional lH As Long = 16, _
       Optional bShadow As Boolean = False, _
       Optional lShadowColor As Long = -1, _
       Optional lMaskColor As Long = -1)

    On Error GoTo PROC_ERR_DrawPicture
        
        If Not IsCreated() Then Exit Sub
        
        Dim lFlags As Long
        Dim hBrush As Long
                          
        ' Get image type
100     Select Case m_PictureF.Type
            Case vbPicTypeBitmap
102             lFlags = DST_BITMAP
                
                Dim hdcPaint As Long, hbmOrig As Long
                hdcPaint = CreateCompatibleDC(m_lhDC)
                hbmOrig = SelectObject(hdcPaint, m_PictureF.handle)
                If lMaskColor = -1 Then
                    APIBitBlt m_lhDC, x, y, lW, lH, hdcPaint, 0, 0, vbSrcCopy
                Else
                    pvTransBlt m_lhDC, x, y, lW, lH, hdcPaint, 0, 0, lMaskColor
                End If
                Call SelectObject(hdcPaint, hbmOrig)
                Call DeleteDC(hdcPaint)
                Exit Sub
            
104         Case vbPicTypeIcon
106             lFlags = DST_ICON
108         Case Else
110             lFlags = DST_COMPLEX
        End Select
        ' If shadow is set then create shadow brush
112     If bShadow Then
114         If lShadowColor <> -1 Then
116             hBrush = CreateSolidBrush(lShadowColor)
            Else
118             hBrush = CreateSolidBrush(TranslateColor(vbButtonShadow))
            End If
        End If
        ' Draw icon
120     DrawState m_lhDC, IIf(bShadow, hBrush, 0), 0, m_PictureF.handle, 0, x, y, lW, lH, lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
        
        ' Delete shadow brush object
122     If bShadow Then DeleteObject hBrush

PROC_EXIT:
        Exit Sub
PROC_ERR_DrawPicture:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.DrawPicture", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub
Private Sub pvTransBlt( _
                       ByVal hdcDest As Long, _
                       ByVal xDest As Long, _
                       ByVal yDest As Long, _
                       ByVal nWidth As Long, _
                       ByVal nHeight As Long, _
                       ByVal hdcSrc As Long, _
                       Optional ByVal xSrc As Long = 0, _
                       Optional ByVal ySrc As Long = 0, _
                       Optional ByVal clrMask As OLE_COLOR = vbMagenta, _
                       Optional ByVal hPal As Long = 0)

    Dim hdcMask             As Long ' hDC of the created mask image
    Dim hdcColor            As Long ' hDC of the created color image
    Dim hBmMask             As Long ' Bitmap handle to the mask image
    Dim hbmColor            As Long ' Bitmap handle to the color image
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hpalOld             As Long
    Dim hdcScreen           As Long
    Dim hdcScnBuffer        As Long ' Buffer to do all work on
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    Dim hPalBufferOld       As Long
    Dim lMaskColor          As Long
    Dim hpalHalftone        As Long

    hdcScreen = GetDC(0&)
    ' Validate palette
    If hPal = 0 Then
        hpalHalftone = CreateHalftonePalette(hdcScreen)
        hPal = hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    ' Create a color bitmap to server as a copy of the destination
    ' Do all work on this bitmap and then copy it back over the destination
    ' when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    ' Copy the destination to the screen buffer
    APIBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy
    ' Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    ' hdcSrc, because this will create a DIB section if the original bitmap
    ' is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Now create a monochrome bitmap for the mask
    hBmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    ' First, blt the source bitmap onto the cover.  We do this first
    ' and then use it instead of the source bitmap
    ' because the source bitmap may be
    ' a DIB section, which behaves differently than a bitmap.
    ' (Specifically, copying from a DIB section to a monochrome bitmap
    ' does a nearest-color selection rather than painting based on the
    ' backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hpalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    ' In case hdcSrc contains a monochrome bitmap, we must set the destination
    ' foreground/background colors according to those currently set in hdcSrc
    ' (because Windows will associate these colors with the two monochrome colors)
    Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
    Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
    Call APIBitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
    ' Paint the mask.  What we want is white at the transparent color
    ' from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hBmMask)
    ' When ApiBitBlt'ing from color to monochrome, Windows sets to 1
    ' all pixels that match the background color of the source DC.  All
    ' other bits are set to 0.
    Call SetBkColor(hdcColor, lMaskColor)
    Call SetTextColor(hdcColor, vbWhite)
    Call APIBitBlt(hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy)
    ' Paint the rest of the cover bitmap.
    '
    ' What we want here is black at the transparent color, and
    ' the original colors everywhere else.  To do this, we first
    ' paint the original onto the cover (which we already did), then we
    ' AND the inverse of the mask onto that using the DSna ternary raster
    ' operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    ' Operation Codes", "Ternary Raster Operations", or search in MSDN
    ' for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    ' When ApiBitBlt'ing from monochrome to color, Windows transforms all white
    ' bits (1) to the background color of the destination hDC.  All black (0)
    ' bits are transformed to the foreground color.
    Call SetTextColor(hdcColor, vbBlack)
    Call SetBkColor(hdcColor, vbWhite)
    Call APIBitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, &H220326)
    ' Paint the Mask to the Screen buffer
    Call APIBitBlt(hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd)
    ' Paint the Color to the Screen buffer
    Call APIBitBlt(hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint)
    ' Copy the screen buffer to the screen
    Call APIBitBlt(hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy)
    ' All done!
    Call DeleteObject(SelectObject(hdcColor, hbmColorOld))
    Call SelectPalette(hdcColor, hpalOld, True)
    Call RealizePalette(hdcColor)
    Call DeleteDC(hdcColor)
    Call DeleteObject(SelectObject(hdcScnBuffer, hbmScnBufferOld))
    Call SelectPalette(hdcScnBuffer, hPalBufferOld, 0)
    Call RealizePalette(hdcScnBuffer)
    Call DeleteDC(hdcScnBuffer)
    Call DeleteObject(SelectObject(hdcMask, hbmMaskOld))
    Call DeleteDC(hdcMask)
    Call ReleaseDC(0&, hdcScreen)
    If hpalHalftone <> 0 Then
        Call DeleteObject(hpalHalftone)
    End If
End Sub
'********************************************************************
'* Name: Draw3DRect
'* Description: Draw 3D rectangle.
'********************************************************************
Public Sub Draw3DRect( _
       ByVal lLeft As Long, _
       ByVal lTop As Long, _
       ByVal lRight As Long, _
       ByVal lBottom As Long, _
       ByVal oTopLeftColor As OLE_COLOR, _
       ByVal oBottomRightColor As OLE_COLOR)

    On Error GoTo PROC_ERR_Draw3DRect
        
        If Not IsCreated() Then Exit Sub
        
        Dim hPen As Long
        Dim hPenOld As Long
        Dim tP As POINTAPI
        Dim rcItem As RECT

        ' Setup rectangle dimensions
100     With rcItem
102         .Left = lLeft
104         .Top = lTop
106         .Right = lRight
108         .Bottom = lBottom
        End With
    
        ' Create pen with specified color
110     hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
        ' Save old pen
112     hPenOld = SelectObject(m_lhDC, hPen)
        ' Draw top-left of rectangle
114     MoveToEx m_lhDC, rcItem.Left, rcItem.Bottom - 1, tP
116     LineTo m_lhDC, rcItem.Left, rcItem.Top
118     LineTo m_lhDC, rcItem.Right - 1, rcItem.Top
        ' Restore old pen
120     SelectObject m_lhDC, hPenOld
        ' Delete pen
122     DeleteObject hPen
        ' If there is right border
124     If rcItem.Left <> rcItem.Right Then
            ' Create pen
126         hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
            ' Save old pen
128         hPenOld = SelectObject(m_lhDC, hPen)
            ' Draw bottom-right of rectangle
130         LineTo m_lhDC, rcItem.Right - 1, rcItem.Bottom - 1
132         LineTo m_lhDC, rcItem.Left - 1, rcItem.Bottom - 1
            ' Restore old pen
134         SelectObject m_lhDC, hPenOld
            ' Delete pen
136         DeleteObject hPen
        End If

PROC_EXIT:
        Exit Sub
PROC_ERR_Draw3DRect:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.Draw3DRect", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: DrawText
'* Description: Draw text on memory device contex.
'********************************************************************
Public Sub DrawText( _
            ByVal sText As String, _
            ByRef lLeft As Long, _
            ByRef lTop As Long, _
            ByRef lRight As Long, _
            ByRef lBottom As Long, _
            Optional ByVal lFlags As EMemDCDrawText)

    On Error GoTo PROC_ERR_DrawText
    
100     If Not IsCreated() Then Exit Sub
    
        Dim rc As RECT
        Dim wTextParams As DRAWTEXTPARAMS

        ' Setup text box dimensions
102     With rc
104         .Left = lLeft
106         .Top = lTop
108         .Right = lRight
110         .Bottom = lBottom
        End With
112     wTextParams.cbSize = Len(wTextParams)
        
        ' If windows 9x
        If Not pIsWinNT Then
            ApiDrawTextEx m_lhDC, sText, Len(sText), rc, lFlags, wTextParams
        Else
            DrawTextW m_lhDC, StrPtr(sText), -1, rc, lFlags
        End If

        ' Return text dimensions
116     With rc
118         lLeft = .Left
120         lTop = .Top
122         lRight = .Right
124         lBottom = .Bottom
        End With

PROC_EXIT:
        Exit Sub
PROC_ERR_DrawText:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.DrawText", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: Font
'* Description: Return\sets memory dc font.
'********************************************************************
Public Property Get Font() As StdFont
    On Error GoTo PROC_ERR_Font

100     If Not IsCreated() Then Exit Property
    
        Dim tTM             As TEXTMETRIC
        Dim sFaceName       As String * 80
    
        ' Get text metrics
102     GetTextMetrics m_lhDC, tTM
        ' Return font name
104     GetTextFace m_lhDC, 79, sFaceName
        ' Create a new font object
106     Set Font = New StdFont
        ' And fill it with memory dc font properties
108     With Font
110         .Name = sFaceName
112         .Bold = (tTM.tmWeight >= FW_NORMAL)
114         .Charset = tTM.tmCharSet
116         .Italic = (tTM.tmItalic <> 0)
118         .Strikethrough = (tTM.tmStruckOut <> 0)
120         .Underline = (tTM.tmUnderlined <> 0)
122         .Weight = tTM.tmWeight
124         .Size = (tTM.tmHeight - tTM.tmInternalLeading) * 72 / tTM.tmDigitizedAspectY
        End With

PROC_EXIT:
        Exit Property
PROC_ERR_Font:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.Font", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property
Public Property Set Font(ByVal oValue As StdFont)
    On Error GoTo PROC_ERR_Font

100     If Not IsCreated() Then Exit Property

        Dim tFont           As LOGFONT
    
        ' Fill LOGFONT type with new font properties
102     With tFont
104         CopyMemory .lfFaceName(1), ByVal oValue.Name, Len(oValue.Name) + 1
106         .lfCharSet = oValue.Charset
108         .lfItalic = (-oValue.Italic)
110         .lfStrikeOut = (-oValue.Strikethrough)
112         .lfUnderline = (-oValue.Underline)
114         .lfWeight = oValue.Weight
116         .lfHeight = -(oValue.Size * GetDeviceCaps(m_lhDC, LOGPIXELSY) \ 72)
        End With
        ' Select new font into memory dc and delete the previous one
118     If m_lMemoryFont <> 0 Then
120         Call SelectObject(m_lhDC, m_lOrginalFont)
122         Call DeleteObject(m_lMemoryFont)
        End If
124     m_lMemoryFont = CreateFontIndirect(tFont)
126     m_lOrginalFont = SelectObject(m_lhDC, m_lMemoryFont)

PROC_EXIT:
        Exit Property
PROC_ERR_Font:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.Font", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property

'********************************************************************
'* Name: ForeColor
'* Description: Return\sets text fore color.
'********************************************************************
Public Property Get ForeColor() As Long
    On Error GoTo PROC_ERR_ForeColor

100     If IsCreated() Then ForeColor = GetTextColor(m_lhDC)

PROC_EXIT:
        Exit Property
PROC_ERR_ForeColor:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.ForeColor", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property
Public Property Let ForeColor(ByVal lValue As Long)
    On Error GoTo PROC_ERR_ForeColor

100     If IsCreated() Then SetTextColor m_lhDC, TranslateColor(lValue)

PROC_EXIT:
        Exit Property
PROC_ERR_ForeColor:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.ForeColor", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Property

'********************************************************************
'* Name: TextWidth
'* Description: Return provided text width using current font.
'********************************************************************
Public Function TextWidth(ByVal sText As String) As Long
    On Error GoTo PROC_ERR_TextWidth

100     If Not IsCreated() Then Exit Function
    
        Dim tSA As SIZEAPI
102     GetTextExtentPoint m_lhDC, sText, Len(sText), tSA
104     TextWidth = tSA.cX

PROC_EXIT:
        Exit Function
PROC_ERR_TextWidth:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.TextWidth", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Function

'********************************************************************
'* Name: TextHeight
'* Description: Return provided text height using current font.
'********************************************************************
Public Function TextHeight(ByVal sText As String) As Long
    On Error GoTo PROC_ERR_TextHeight

100     If Not IsCreated() Then Exit Function
    
        Dim tSA As SIZEAPI
102     GetTextExtentPoint m_lhDC, sText, Len(sText), tSA
104     TextHeight = tSA.cY

PROC_EXIT:
        Exit Function
PROC_ERR_TextHeight:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.TextHeight", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Function

'********************************************************************
'* Name: GetPixel
'* Description: Return current pixel color.
'********************************************************************
Public Function GetPixel(ByVal x As Long, ByVal y As Long, Optional lhDC As Long = -1) As Long
    On Error GoTo PROC_ERR_GetPixel

100     If Not IsCreated() Then Exit Function
    
        ' If user specified from which dc to return pixel color then
        ' return it from that one otherwise from memory dc
102     If lhDC <> -1 Then
104         GetPixel = APIGetPixel(lhDC, x, y)
        Else
106         GetPixel = APIGetPixel(m_lhDC, x, y)
        End If

PROC_EXIT:
        Exit Function
PROC_ERR_GetPixel:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.GetPixel", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Function

'********************************************************************
'* Name: SetPixel
'* Description: Set current pixel color.
'********************************************************************
Public Sub SetPixel(ByVal x As Long, ByVal y As Long, ByVal lColor As Long)
    On Error GoTo PROC_ERR_SetPixel

100     If Not IsCreated() Then Exit Sub
    
        ' Set pixel color
102     APISetPixel m_lhDC, x, y, lColor

PROC_EXIT:
        Exit Sub
PROC_ERR_SetPixel:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.SetPixel", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: ShiftColor
'* Description: Returns shifted color(added or removed certain color)
'* from specified color.
'********************************************************************
Public Function ShiftColor(ByVal lColor As Long, ByVal lValue As Long, Optional bIsXP As Boolean = False) As Long
    Dim lRed As Long
    Dim lBlue As Long
    Dim lGreen As Long
    
    lColor = TranslateColor(lColor)
    
    If bIsXP Then
        lBlue = ((lColor \ &H10000) Mod &H100)
        lBlue = lBlue + ((lBlue * lValue) \ &HC0)
    Else
        lBlue = ((lColor \ &H10000) Mod &H100) + lValue
    End If
    lGreen = ((lColor \ &H100) Mod &H100) + lValue
    lRed = (lColor And &HFF) + lValue
    
    ' Keep color values in range of 0-255
    If lRed < 0 Then
        lRed = 0
    ElseIf lRed > 255 Then
        lRed = 255
    End If
    If lGreen < 0 Then
        lGreen = 0
    ElseIf lGreen > 255 Then
        lGreen = 255
    End If
    If lBlue < 0 Then
        lBlue = 0
    ElseIf lBlue > 255 Then
        lBlue = 255
    End If
    
    ' Return shifted color
    ShiftColor = RGB(lRed, lGreen, lBlue)
End Function

'********************************************************************
'* Name: AlphaBlend
'* Description: Returns alpha blend of two specified colors.
'********************************************************************
Public Function AlphaBlend(ByVal lColorOne As Long, ByVal lColorTwo As Long, ByVal lAlpha As Long) As OLE_COLOR
    Dim tClrFore         As RGBType
    Dim tClrBack         As RGBType
    
    ' Translate both colors
    lColorOne = TranslateColor(lColorOne)
    lColorTwo = TranslateColor(lColorTwo)
    ' Split second color into RGB elements
    With tClrBack
        .r = GetRGB(lColorTwo, 1)
        .g = GetRGB(lColorTwo, 2)
        .b = GetRGB(lColorTwo, 3)
    End With
    With tClrFore
        ' Split the first color also
        .r = GetRGB(lColorOne, 1)
        .g = GetRGB(lColorOne, 2)
        .b = GetRGB(lColorOne, 3)
        ' Alpha blend each color item
        .r = (.r * lAlpha + tClrBack.r * (255 - lAlpha)) / 255
        .g = (.g * lAlpha + tClrBack.g * (255 - lAlpha)) / 255
        .b = (.b * lAlpha + tClrBack.b * (255 - lAlpha)) / 255
    End With
    ' Return blended color
    AlphaBlend = RGB(tClrFore.r, tClrFore.g, tClrFore.b)
End Function

'********************************************************************
'* Name: DrawGradient
'* Description: Draws gradient between two colors.
'********************************************************************
Public Sub DrawGradient( _
       ByVal lLeft As Long, _
       ByVal lTop As Long, _
       ByVal lRight As Long, _
       ByVal lBottom As Long, _
       ByVal lStartColor As Long, _
       ByVal lEndColor As Long, _
       Optional ByVal sngAngle As Single = 90)

    On Error GoTo PROC_ERR_DrawGradient
    
        Dim lhDC        As Long
        Dim lDesktophDC As Long
        Dim lBitmap     As Long
        Dim lBitmapOld  As Long
        Dim lWidth      As Long
        Dim lHeight     As Long
        Dim bDone       As Boolean
        Dim iIncX       As Integer
        Dim iIncY       As Integer
        Dim lIdx        As Long
        Dim lRet        As Long
        Dim hPen        As Long
        Dim hOldPen     As Long
        Dim laColors()  As Long
        Dim fMovX       As Single
        Dim fMovY       As Single
        Dim fDist       As Single
        Dim fAngle      As Single
        Dim fLongSide   As Single
        Dim uTmpPt      As POINTAPI
        Dim uaPts()     As POINTAPI
        Dim uaTmpPts()  As PointSng
        Dim mfAngle     As Single

        ' Calculate angle
100     mfAngle = CDbl(sngAngle) - Int(Int(CDbl(sngAngle) / 360#) * 360#)
        ' Calculate width and height
102     lWidth = lRight - lLeft
104     lHeight = lBottom - lTop
        ' Translate start and end color
106     lStartColor = TranslateColor(lStartColor)
108     lEndColor = TranslateColor(lEndColor)
    
        ' Create a new dc based on desktop
110     lDesktophDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Create compatible dc with desktop one
112     lhDC = CreateCompatibleDC(lDesktophDC)
        ' Create a new bitmap on which we will draw
114     lBitmap = CreateCompatibleBitmap(lDesktophDC, lWidth, lHeight)
        ' Select bitmap into new dc
116     lBitmapOld = SelectObject(lhDC, lBitmap)
        ' Delete temp dc
118     DeleteDC lDesktophDC

        'Start with center of rect
120     ReDim uaTmpPts(2)
122     uaTmpPts(2).x = Int(lWidth / 2)
124     uaTmpPts(2).y = Int(lHeight / 2)

        'Calc distance to furthest edge as if rect were square
126     fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
128     fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2

        'Create points to the left and the right at a 0º angle (horizontal)
130     uaTmpPts(0).x = uaTmpPts(2).x - fDist
132     uaTmpPts(0).y = uaTmpPts(2).y
134     uaTmpPts(1).x = uaTmpPts(2).x + fDist
136     uaTmpPts(1).y = uaTmpPts(2).y

        'Lines will be drawn perpendicular to mfAngle so
        'add 90º and correct for 360º wrap
138     fAngle = CDbl(mfAngle + 90) - Int(Int(CDbl(mfAngle + 90) / 360#) * 360#)

        'Rotate second and third points to fAngle
140     Call pRotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
142     Call pRotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)

        'We now have a line that crosses the center and
        'two sides of the rect at the correct angle.

        'Calc the starting quadrant, direction of and amount of first move
        '(fMovX, fMovY moves line from center to starting edge)
        'and direction of each incremental move (iIncX, iIncY).
144     Select Case mfAngle
            Case 0 To 90
                'Left Bottom
146             If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                    'Move line to left edge; Draw left to right
148                 fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, -uaTmpPts(0).x, -uaTmpPts(1).x)
150                 fMovY = 0
152                 iIncX = 1
154                 iIncY = 0
                Else
                    'Move line to bottom edge; Draw bottom to top
156                 fMovX = 0
158                 fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, lHeight - uaTmpPts(1).y, lHeight - uaTmpPts(0).y)
160                 iIncX = 0
162                 iIncY = -1
                End If
164         Case 90 To 180
                'Right Bottom
166             If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                    'Move line to right edge; Draw right to left
168                 fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, lWidth - uaTmpPts(1).x, lWidth - uaTmpPts(0).x)
170                 fMovY = 0
172                 iIncX = -1
174                 iIncY = 0
                Else
                    'Move line to bottom edge; Draw bottom to top
176                 fMovX = 0
178                 fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, lHeight - uaTmpPts(1).y, lHeight - uaTmpPts(0).y)
180                 iIncX = 0
182                 iIncY = -1
                End If
184         Case 180 To 270
                'Right Top
186             If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                    'Move line to right edge; Draw right to left
188                 fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, lWidth - uaTmpPts(1).x, lWidth - uaTmpPts(0).x)
190                 fMovY = 0
192                 iIncX = -1
194                 iIncY = 0
                Else
                    'Move line to top edge; Draw top to bottom
196                 fMovX = 0
198                 fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, -uaTmpPts(0).y, -uaTmpPts(1).y)
200                 iIncX = 0
202                 iIncY = 1
                End If
204         Case Else   '(270 to 360)
                'Left Top
206             If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                    'Move line to left edge; Draw left to right
208                 fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, -uaTmpPts(0).x, -uaTmpPts(1).x)
210                 fMovY = 0
212                 iIncX = 1
214                 iIncY = 0
                Else
                    'Move line to top edge; Draw top to bottom
216                 fMovX = 0
218                 fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, -uaTmpPts(0).y, -uaTmpPts(1).y)
220                 iIncX = 0
222                 iIncY = 1
                End If
        End Select

        'At this point we could calculate where the lines will cross the rect edges, but
        'this would slow things down. The picObj clipping region will take care of this.

        'Start with 1000 points and add more if needed. This increases
        'speed by not re-dimming the array in each loop.
224     ReDim uaPts(999)

        'Set the first two points in the array
226     uaPts(0).x = uaTmpPts(0).x + fMovX
228     uaPts(0).y = uaTmpPts(0).y + fMovY
230     uaPts(1).x = uaTmpPts(1).x + fMovX
232     uaPts(1).y = uaTmpPts(1).y + fMovY

234     lIdx = 2
        'Create the rest of the points by incrementing both points
        'on each line iIncX, iIncY from the previous line's points.
        'Where we stop depends on the direction of travel.
        'We'll continue until both points in a set reach the end.
236     While Not bDone
238         uaPts(lIdx).x = uaPts(lIdx - 2).x + iIncX
240         uaPts(lIdx).y = uaPts(lIdx - 2).y + iIncY
242         lIdx = lIdx + 1
244         Select Case True
                Case iIncX > 0  'Moving Left to Right
246                 bDone = uaPts(lIdx - 1).x > lWidth And uaPts(lIdx - 2).x > lWidth
248             Case iIncX < 0  'Moving Right to Left
250                 bDone = uaPts(lIdx - 1).x < 0 And uaPts(lIdx - 2).x < 0
252             Case iIncY > 0  'Moving Top to Bottom
254                 bDone = uaPts(lIdx - 1).y > lHeight And uaPts(lIdx - 2).y > lHeight
256             Case iIncY < 0  'Moving Bottom to Top
258                 bDone = uaPts(lIdx - 1).y < 0 And uaPts(lIdx - 2).y < 0
            End Select
260         If (lIdx Mod 1000) = 0 Then
262             ReDim Preserve uaPts(UBound(uaPts) + 1000)
            End If
        Wend

        'Free excess memory (may have 1001 points dimmed to 2000)
264     ReDim Preserve uaPts(lIdx - 1)

        'Create the array of colors blending from mlColor1 to mlColor2
266     lRet = pBlendColors(lStartColor, lEndColor, lIdx / 2, laColors)

        'Now draw each line in it's own color
268     For lIdx = 0 To UBound(uaPts) - 1 Step 2
            'Move to next point
270         lRet = MoveToEx(lhDC, uaPts(lIdx).x, uaPts(lIdx).y, uTmpPt)
            'Create the colored pen and select it into the DC
272         hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
274         hOldPen = SelectObject(lhDC, hPen)
            'Draw the line
276         lRet = LineTo(lhDC, uaPts(lIdx + 1).x, uaPts(lIdx + 1).y)
            'Get the pen back out of the DC and destroy it
278         lRet = SelectObject(lhDC, hOldPen)
280         lRet = DeleteObject(hPen)
        Next

        ' Transfer bitmap from temporary memorydc to memorydc
282     APIBitBlt m_lhDC, lLeft, lTop, lWidth, lHeight, lhDC, 0, 0, vbSrcCopy
    
        ' Free the memory
284     Erase laColors
286     Erase uaPts
288     Erase uaTmpPts
    
        ' Delete temporary memory DC
290     If lBitmapOld <> 0 Then SelectObject lhDC, lBitmapOld
292     If lBitmap <> 0 Then DeleteObject lBitmap
294     If lhDC <> 0 Then DeleteDC lhDC

PROC_EXIT:
        Exit Sub
PROC_ERR_DrawGradient:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.DrawGradient", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: pBlendColors
'* Description: Creates an array of colors blending from Color1 to
'* Color2 in lSteps number of steps. Returns the count and fills
'* the laRetColors() array.
'* Credits: Kath-Rock
'********************************************************************
Private Function pBlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long
    On Error GoTo PROC_ERR_pBlendColors

        Dim lIdx    As Long
        Dim lRed    As Long
        Dim lGrn    As Long
        Dim lBlu    As Long
        Dim fRedStp As Single
        Dim fGrnStp As Single
        Dim fBluStp As Single

        'Stop possible error
100     If lSteps < 2 Then lSteps = 2
    
        'Extract Red, Blue and Green values from the start and end colors.
102     lRed = (lColor1 And &HFF&)
104     lGrn = (lColor1 And &HFF00&) / &H100
106     lBlu = (lColor1 And &HFF0000) / &H10000
    
        'Find the amount of change for each color element per color change.
108     fRedStp = pDiv(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
110     fGrnStp = pDiv(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
112     fBluStp = pDiv(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
        'Create the colors
114     ReDim laRetColors(lSteps - 1)
116     laRetColors(0) = lColor1            'First Color
118     laRetColors(lSteps - 1) = lColor2   'Last Color
120     For lIdx = 1 To lSteps - 2          'All Colors between
122         laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
               (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
               (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
        Next
    
        'Return number of colors in array
124     pBlendColors = lSteps

PROC_EXIT:
        Exit Function
PROC_ERR_pBlendColors:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pBlendColors", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Function

'********************************************************************
'* Name: pRotatePoint
'* Description:
'* Credits: Kath-Rock
'********************************************************************
Private Sub pRotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)
    On Error GoTo PROC_ERR_pRotatePoint
        
        Dim fDX         As Single
        Dim fDY         As Single
        Dim fRadians    As Single

100     fRadians = fDegrees * RADS
102     fDX = uRotatePt.x - uAxisPt.x
104     fDY = uRotatePt.y - uAxisPt.y
106     uRotatePt.x = uAxisPt.x + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
108     uRotatePt.y = uAxisPt.y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))

PROC_EXIT:
        Exit Sub
PROC_ERR_pRotatePoint:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pRotatePoint", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Sub

'********************************************************************
'* Name: pDiv
'* Description: Divides dNumer by dDenom if dDenom <> 0. Eliminates
'* 'Division By Zero' error.
'* Credits: Kath-Rock
'********************************************************************
Private Function pDiv(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    On Error GoTo PROC_ERR_pDiv

100     If dDenom <> 0 Then pDiv = dNumer / dDenom

PROC_EXIT:
        Exit Function
PROC_ERR_pDiv:
        Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pDiv", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
        Resume PROC_EXIT
End Function

Public Sub TileArea( _
       ByVal x As Long, _
       ByVal y As Long, _
       ByVal Width As Long, _
       ByVal Height As Long, _
       ByVal oPicture As StdPicture, _
       ByVal SrcWidth As Long, _
       ByVal SrcHeight As Long, _
       ByVal lOffsetY As Long _
       )
    Dim lSrcX As Long
    Dim lSrcY As Long
    Dim lSrcStartX As Long
    Dim lSrcStartY As Long
    Dim lSrcStartWidth As Long
    Dim lSrcStartHeight As Long
    Dim lDstX As Long
    Dim lDstY As Long
    Dim lDstWidth As Long
    Dim lDstHeight As Long

    lSrcStartX = (x Mod SrcWidth)
    lSrcStartY = ((y + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If

            Me.DrawPicture oPicture, lDstX, lDstY, SrcWidth, SrcHeight

            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub

'********************************************************************
'* Name: pOsVersion
'* Description: Returns OS version number.
'********************************************************************
Private Property Get pOsVersion() As Long
    Dim tOVI As OSVERSIONINFO
    tOVI.dwOSVersionInfoSize = Len(tOVI)
    If GetVersionEx(tOVI) Then pOsVersion = tOVI.dwMajorVersion * 256 + tOVI.dwMinorVersion
End Property

'********************************************************************
'* Name: IsWinNT
'* Description: Return true if windows platform is NT.
'********************************************************************
Private Function pIsWinNT() As Boolean
    Dim lpOS As OSVERSIONINFO
    lpOS.dwOSVersionInfoSize = Len(lpOS)
    GetVersionEx lpOS
    pIsWinNT = (lpOS.dwPlatformID = 2)
End Function

'********************************************************************
'* Name: IconToPicture
'* Description: Creates a StdPicture object from HICON handle.
'********************************************************************
Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    Dim oNewPic         As Picture
    Dim lpPictDesc      As PictDesc
    Dim aGuid(0 To 3)   As Long

    On Error Resume Next
    '--- check argument
    If hIcon = 0 Then Exit Function
    '--- fill struct
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeIcon
        .hBmpOrIcon = hIcon
    End With
    '--- fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '--- do convert
    OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
    '--- success
    Set IconToPicture = oNewPic
End Function

'********************************************************************
'* Name: BitmapToPicture
'* Description: Creates a StdPicture object from HBITMAP handle.
'********************************************************************
Public Function BitmapToPicture(ByVal hBmp As Long, Optional ByVal hPal As Long = 0) As IPicture
    Dim oNewPic         As Picture
    Dim lpPictDesc      As PictDesc
    Dim aGuid(0 To 3)   As Long
    '--- fill struct
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeBitmap
        .hBmpOrIcon = hBmp
        .hPal = hPal
    End With
    '--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '--- Create picture from bitmap handle
    OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
    '--- success
    Set BitmapToPicture = oNewPic
End Function
