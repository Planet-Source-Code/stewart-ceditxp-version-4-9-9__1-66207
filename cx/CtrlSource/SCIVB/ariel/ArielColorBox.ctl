VERSION 5.00
Begin VB.UserControl ArielColorBox 
   BackColor       =   &H80000005&
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   FillStyle       =   0  'Solid
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   Begin VB.PictureBox picPopup 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1770
      Left            =   0
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2955
   End
End
Attribute VB_Name = "ArielColorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0


'--------------------------------------------------------------------
'Module     : ArielColorBox
'Description: Ariel Color Box ActiveX Control
'Version    : V1.1 Sep 2000
'Release    : VB6
'Copyright  : © T De Lange, 2000
'--------------------------------------------------------------------
'V1.0 Sep 00 Original version, based on ColorCombo
'V1.1 Sep 00 With hWnd property, on request
'--------------------------------------------------------------------
'Distribution Notes
'a) Provision is made for three internal color selections: Active, Change
'   & Selected. The Active index tracks the mouse movement and is used
'   for hovering. Doesn't fire any events. The Change index tracks the
'   mouse down position and fires the Change() event when changes occur.
'   The Selected index tracks the final color selection only on mouseup
'   position and triggers the Click event.
'--------------------------------------------------------------------

DefLng A-N, P-Z
DefBool O

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


'---------------------------------------------------
'Internal constants
'---------------------------------------------------
Const mMinWidth = 35    '21 pixels for image and 13 for dropdown
Const mDropWidth = 13   'Width of dropdown button
Const mPopupEdge = 8    'Edge margin around color boxes in popup window

Public Enum ArielPalette
  ap8x2
  ap9x3
  ap8x8
  ap16x8
  ap24x8
  ap24x10
  ap8x4x4
  ap8x8x6
  ap12x6x6
  ap16x8x6
  ap24x8x6
  ap32x8x6
End Enum

Private Enum ArielHues
  apRed = 0
  apCopper = 12
  apOrange = 20
  apBronze = 25
  apGold = 30
  apTopaz = 35
  apTourmaline = 37
  apYellow = 40
  apCitrine = 45
  apLemon = 50
  apLime = 60
  apEmerald = 70
  apGreen = 80
  apBeryl = 90
  apJade = 100
  apTurquoise = 110
  apCyan = 120
  apAqua = 125
  apTeal = 130
  apAzure = 140
  apSapphire = 150
  apBlue = 160
  apIndigo = 170
  apLavender = 180
  apAmethyst = 185
  apViolet = 190
  apCobalt = 195
  apMagenta = 200
  apFuschia = 210
  apPink = 220
  apCrimson = 230
End Enum

Private Enum ArielHueLum
  apFaint = 225
  apPale = 210
  apLight = 180
  apSoft = 150
  apStd = 120
  apDense = 105
  apDeep = 90
  apMurky = 75
  apDark = 60
  apPitch = 40
End Enum

Private Enum ArielBwgLum
  apWhite = 240
  apSilver = 210
  apPlatinum = 200
  apChrome = 180
  apNickel = 160
  apTitanium = 150
  apGrey = 120
  apDarkGrey = 90
  apEbony = 80
  apCharcoal = 60
  apMidnight = 40
  apPitchBlack = 30
  apBlack = 0
End Enum


'---------------------------------------------------
'Property Variables:
Dim mPopupEnabled As Boolean
Dim mPalette As ArielPalette

'---------------------------------------------------
'Event Declarations:
'---------------------------------------------------
Public Event Click()
Public Event Change(NewActiveColor As Long)
Public Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Public Event Popup()    'Fired when popupbutton is clicked

'---------------------------------------------------
'Default Property Values:
'---------------------------------------------------
Const mdefActiveColor = 0
Const mdefActiveText = ""
Const mdefSelectedColor = 0
Const mdefPalette = 0
Const mdefPopupEnabled = True

'---------------------------------------------------
'Internal Control Variables
'---------------------------------------------------
Dim cDown As Boolean            'Dropdown button is down
Dim rDrp As RECT                'DropDown rectangle (incl border)
Dim ctlCancel As Object         'Parent Ctrl having 'cancel' property
Dim nHue As Long                'Max Hue Color Rectangle No c(0..nHue)
Dim nSat As Long                'Max Sat Color Rectangle No d(0..nSat)
Dim c() As ColorRect            'Hue color Rectangle Array
Dim d() As ColorRect            'Sat color Rectangle Array
Dim rcHue As ColorRect          'ColorRect containing the current hue for the sat/lum boxes
Dim rcCurr As ColorRect         'ColorRect containing the current color
Dim rcActive As ColorRect       'ColorRect containing the active color for Change() events
Dim rcSel As ColorRect          'ColorRect containing the selected color for Click() events
Dim mPopupSpace As Long         'Spacing between popup color boxes
Dim mPopupSize As Long          'Size of popup color boxes
Dim cMouseMove As Boolean
Dim cMouseX As Long
Dim cMouseY As Long
Dim cExtended As Boolean        'When true, uses the d() colorboxes and hovertimer
Dim cHueCols As Long            'No of Hue Color Rectangle Columns
Dim cHueRows As Long            'No of Hue Color Rectangle Rows
Dim cSatCols As Long            'No of Sat Color Rectangle Columns
Dim cSatRows As Long            'No of Sat Color Rectangle Rows
Dim WithEvents tmrHover As ArielTimer
Attribute tmrHover.VB_VarHelpID = -1

'---------------------------------------------------
'Api Type Declarations
'---------------------------------------------------

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type ColorRect
  IsHue As Boolean
  rc As RECT
  Color As OLE_COLOR
  HSL As HSLColor
  Chromatic As Boolean    'Color or Greyscale?
  ColorName As String
  Ok As Boolean
End Type

'---------------------------------------------------
'Api Function Declarations
'---------------------------------------------------
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Note: the following declaration in the API viewer is incorrect!
'Private Declare Function PtInRect Lib "user32" (lpRect As Rect, pt As PointApi) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long

'---------------------------------------------------
'Api Constants
'---------------------------------------------------
'For determining which part of the border to draw
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_SOFT = &H1000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

'For drawing borders with the DrawEdge() function
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = BDR_RAISEDOUTER Or BDR_SUNKENINNER
Private Const EDGE_ETCHED = BDR_SUNKENOUTER Or BDR_RAISEDINNER
Private Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Private Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER

'For changing the windowstyle of the popup window
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80

'MemberInfo=10,1,2,0
Public Property Get ActiveColor() As OLE_COLOR
'-------------------------------------------
'Read the active color property
'-------------------------------------------
ActiveColor = rcActive.Color

End Property

'MemberInfo=13,1,2,
Public Property Get ActiveText() As String
'-------------------------------------------
'Read the active text property
'(The colorname of the active color)
'-------------------------------------------
ActiveText = rcActive.ColorName

End Property

Private Sub ArrangeHueRect(Down As Boolean, NoItems As Long, xOff As Long, yOff As Long)
'---------------------------------------------------------------------------------
'Arrange the Hue (left) Color Rectangles
'Down     :  True  -> Vertical priority
'            False -> Horisontal priority
'NoItems  :  No of items: Rows (Down=true), Cols (Down=False)
'---------------------------------------------------------------------------------
Dim X, Y, i

For i = 0 To nHue
  If Down Then
    X = (i \ NoItems) * (mPopupSize + mPopupSpace)
    Y = (i Mod NoItems) * (mPopupSize + mPopupSpace)
  Else
    X = (i Mod NoItems) * (mPopupSize + mPopupSpace)
    Y = (i \ NoItems) * (mPopupSize + mPopupSpace)
  End If
  With c(i).rc
    .Left = xOff + X
    .Right = xOff + X + mPopupSize
    .Top = yOff + Y
    .Bottom = yOff + Y + mPopupSize
  End With
Next

End Sub

Private Sub ArrangeSatRect(Down As Boolean, NoItems As Long, xOff As Long, yOff As Long)
'---------------------------------------------------------------------------------
'Arrange the Sat (right) Color Rectangles
'Down     :  True  -> Vertical priority
'            False -> Horisontal priority
'NoItems  :  No of items: Rows (Down=true), Cols (Down=False)
'---------------------------------------------------------------------------------
Dim X, Y, i

For i = 0 To nSat
  If Down Then
    X = (i \ NoItems) * (mPopupSize + mPopupSpace)
    Y = (i Mod NoItems) * (mPopupSize + mPopupSpace)
  Else
    X = (i Mod NoItems) * (mPopupSize + mPopupSpace)
    Y = (i \ NoItems) * (mPopupSize + mPopupSpace)
  End If
  With d(i).rc
    .Left = xOff + X
    .Right = xOff + X + mPopupSize
    .Top = yOff + Y
    .Bottom = yOff + Y + mPopupSize
  End With
Next

End Sub

'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
'----------------------------------------------------
'Expose UserControl.backcolor to User
'----------------------------------------------------
BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
'------------------------------------------------------------
'Set new back color property
'------------------------------------------------------------
UserControl.BackColor() = NewBackColor
PropertyChanged "BackColor"
PaintMain

End Property

Private Sub ChangeActiveColor(rc As ColorRect)
'----------------------------------------------------------------
'Change the current active color and update references
'If changed, fire the Change() event
'----------------------------------------------------------------
If rc.Color <> rcActive.Color Then
  rcActive = rc
  PropertyChanged "ActiveColor"
  PropertyChanged "ActiveText"
  RaiseEvent Change(rcActive.Color)
  PaintMain
End If

End Sub

Private Sub ChangeSelectedColor(rc As ColorRect)
'----------------------------------------------------------------
'Update the current selected color and update references
'If changed, fire the Click() event
'----------------------------------------------------------------
Dim Update As Boolean, Changed As Boolean

Changed = rc.Color <> rcSel.Color
If Changed Then
  rcSel = rc
End If
Update = rcSel.Color <> rcActive.Color
If Update Or Changed Then
  If Changed Then
    PropertyChanged "SelectedColor"
    PropertyChanged "Text"
    RaiseEvent Click
  End If
  PaintMain
End If

End Sub

Private Sub DrawRectBorder(rc As ColorRect, cCurrent As Boolean, cDown As Boolean)
'------------------------------------------------------------------------------
'Draws a border around the specified color rectangle
'cCurrent, cDown  : Determines the shape of the border, as follows:
'True    , True     Sunken Edge
'True    , False    Raised Edge
'False   , <n/a>    No Edge (restore mint state)
'------------------------------------------------------------------------------
Dim rct As RECT, Ok, pt As POINTAPI

If rc.Ok Then
  rct = rc.rc
  With picPopup
    If cCurrent Then
      'Draw Sunken or Raised Edge, depending on cDown parameter
      If cDown Then
        .ForeColor = vb3DShadow                             'Dark shadow (not black!)
      Else
        .ForeColor = vb3DHighlight                          'Highlight (not white!)
      End If
      Ok = MoveToEx(.hdc, rct.Left - 0, rct.Bottom + 0, pt) 'Bottom Left corner
      Ok = LineTo(.hdc, rct.Left - 0, rct.Top - 0)          'Top Left
      Ok = LineTo(.hdc, rct.Right + 0, rct.Top - 0)         'Top Right
      If cDown Then
        .ForeColor = vb3DHighlight                          'Highlight (not white!)
      Else
        .ForeColor = vb3DShadow                             'Dark shadow (not black!)
      End If
      Ok = LineTo(.hdc, rct.Right + 0, rct.Bottom + 0)      'Bottom Right
      Ok = LineTo(.hdc, rct.Left - 1, rct.Bottom + 0)       'Bottom Left
    Else
      'Restore flat edge
      .ForeColor = vbButtonFace
      Ok = MoveToEx(.hdc, rct.Left - 0, rct.Bottom + 0, pt) 'Bottom Left corner
      Ok = LineTo(.hdc, rct.Left - 0, rct.Top - 0)          'Top Left
      Ok = LineTo(.hdc, rct.Right + 0, rct.Top - 0)         'Top Right
      Ok = LineTo(.hdc, rct.Right + 0, rct.Bottom + 0)      'Bottom Right
      Ok = LineTo(.hdc, rct.Left - 1, rct.Bottom + 0)       'Bottom Left
    End If
  End With
End If

End Sub

'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
'------------------------------------------------
'Read property
'------------------------------------------------
Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
'--------------------------------------------------------
'Write property
'--------------------------------------------------------
UserControl.Enabled() = NewEnabled
PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont
'----------------------------------------------------------
'Returns the current font
'----------------------------------------------------------
Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal NewFont As StdFont)
'-----------------------------------------------------
'Changes the font
'-----------------------------------------------------
If Not (NewFont Is Nothing) Then
  Set UserControl.Font = NewFont
  PropertyChanged "Font"
  ResizeCtrl
  Refresh
End If

End Property

'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
'------------------------------------------------
'Get the forecolor (text color)
'------------------------------------------------
ForeColor = UserControl.ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
'-------------------------------------------------------------
'Set the new fore (text) color
'-------------------------------------------------------------
UserControl.ForeColor() = NewForeColor
PropertyChanged "ForeColor"
PaintMain
End Property

Private Function hasFocus() As Boolean
'----------------------------------------------------
'Determine if Usercontrol currently has focus
'Getfocus() returns the hwnd of ctrl with focus
'Match with the UserCtrl.hWnd
'----------------------------------------------------
hasFocus = (GetFocus = UserControl.hWnd)

End Function

Private Sub HidePopUp()
'---------------------------------------------------------------------------
'Hides the Popup window, release the mouse and Cancel key capture
'and disable the hover timer
'---------------------------------------------------------------------------
If picPopup.visible Then
  'Release mouse capturing
  If GetCapture = picPopup.hWnd Then
    ReleaseCapture
  End If
  'Hide the popup
  picPopup.visible = False
  DoEvents
  'Repaint Ctrl
  'UserControl_Paint
  'Restore Cancel property if a default 'Cancel' control was found
  If Not ctlCancel Is Nothing Then
    ctlCancel.Cancel = True
  End If
  'Disable hover timer
  tmrHover.Enabled = False
End If

End Sub

Private Sub InitHueColorBoxes()
'-------------------------------------------------------
'Initialise Hue color rectangles (left hand side)
'-------------------------------------------------------
Dim HSL As HSLColor
Dim i, j, k
Dim ch As Variant       'Array with Color Hues
Dim chn As Variant      'Array with Color Hue Names
Dim cl As Variant       'Array with Color Luminescence values
Dim cln As Variant      'Array with Color Lum Names
Dim bl As Variant       'Array with B/W   Luminescence values
Dim bn As Variant       'Array with B/W   Names

Select Case mPalette
Case ap8x2
  '6 colors x 2 lum + 4 Achromatic = 16
  ch = Array(apRed, apYellow, apGreen, apCyan, apBlue, apMagenta)
  chn = Array("Red", "Yellow", "Green", "Cyan", "Blue", "Magenta")
  cl = Array(apStd, apDark)
  cln = Array("", "Dark ")
  bl = Array(apWhite, apGrey, apChrome, apBlack)
  bn = Array("White", "Grey", "Chrome", "Black")
  mPopupSpace = 4
  mPopupSize = 14
  cHueCols = 8
  cHueRows = 2
  cExtended = False
Case ap9x3
  '7 colors x 3 lum + 6 Achromatic = 27
  ch = Array(apRed, apOrange, apYellow, apGreen, apCyan, apBlue, apMagenta)
  chn = Array("Red", "Orange", "Yellow", "Green", "Cyan", "Blue", "Magenta")
  cl = Array(apLight, apStd, apDark)
  cln = Array("Light ", "", "Dark ")
  bl = Array(apWhite, apChrome, apCharcoal, apSilver, apGrey, apBlack)
  bn = Array("White", "Chrome", "Charcoal", "Silver", "Grey", "Black")
  mPopupSpace = 3   'Spacing between popup color boxes
  mPopupSize = 12   'Size of popup color boxes
  cHueCols = 9
  cHueRows = 3
  cExtended = False
Case ap8x8
  '7 colors x 8 lum + 8 Achromatic = 64
  ch = Array(apRed, apOrange, apYellow, apGreen, apCyan, apBlue, apMagenta)
  chn = Array("Red", "Orange", "Yellow", "Green", "Cyan", "Blue", "Magenta")
  cl = Array(apFaint, apPale, apLight, apStd, apDense, apDeep, apDark, apPitch)
  cln = Array("Faint ", "Pale ", "Light ", "", "Dense ", "Deep ", "Dark ", "Pitch ")
  bl = Array(apWhite, apSilver, apChrome, apTitanium, apGrey, apDarkGrey, apCharcoal, apBlack)
  bn = Array("White", "Silver", "Chrome", "Titanium", "Grey", "Dark Grey", "Charcoal", "Black")
  mPopupSpace = 2
  mPopupSize = 11
  cHueCols = 8
  cHueRows = 8
  cExtended = False
Case ap16x8
  '15 colors x 8 lum + 8 Achromatic = 128
  ch = Array(apRed, apCopper, apOrange, apGold, apYellow, apLime, apGreen, _
  apTurquoise, apCyan, apTeal, apBlue, apLavender, apMagenta, apFuschia, apPink)
  chn = Array("Red", "Copper", "Orange", "Gold", "Yellow", "Lime", "Green", _
  "Turquoise", "Cyan", "Teal", "Blue", "Lavender", "Magenta", "Fuschia", "Pink")
  cl = Array(apFaint, apPale, apLight, apStd, apDense, apDeep, apDark, apPitch)
  cln = Array("Faint ", "Pale ", "Light ", "", "Dense ", "Deep ", "Dark ", "Pitch ")
  bl = Array(apWhite, apSilver, apChrome, apTitanium, apGrey, apDarkGrey, apCharcoal, apBlack)
  bn = Array("White", "Silver", "Chrome", "Titanium", "Grey", "Dark Grey", "Charcoal", "Black")
  mPopupSpace = 2
  mPopupSize = 10
  cHueCols = 16
  cHueRows = 8
  cExtended = False
Case ap24x8
  '23 colors x 8 lum + 8 Achromatic = 192
  ch = Array(apRed, apCopper, apOrange, apGold, apTopaz, apYellow, _
  apLemon, apLime, apGreen, apJade, apTurquoise, apCyan, apAqua, apTeal, _
  apAzure, apSapphire, apBlue, apLavender, apViolet, apMagenta, apFuschia, _
  apPink, apCrimson)
  chn = Array("Red", "Copper", "Orange", "Gold", "Topaz", "Yellow", _
  "Lemon", "Lime", "Green", "Jade", "Turquoise", "Cyan", "Aqua", "Teal", _
  "Azure", "Sapphire", "Blue", "Lavender", "Violet", "Magenta", "Fuschia", _
  "Pink", "Crimson")
  cl = Array(apFaint, apPale, apLight, apStd, apDense, apDeep, apDark, apPitch)
  cln = Array("Faint ", "Pale ", "Light ", "", "Dense ", "Deep ", "Dark ", "Pitch ")
  bl = Array(apWhite, apSilver, apChrome, apTitanium, apGrey, apDarkGrey, apCharcoal, apBlack)
  bn = Array("White", "Silver", "Chrome", "Titanium", "Grey", "Dark Grey", "Charcoal", "Black")
  mPopupSpace = 2
  mPopupSize = 9
  cHueCols = 24
  cHueRows = 8
  cExtended = False
Case ap24x10
  '23 colors x 10 lum + 10 Achromatic = 240
  ch = Array(apRed, apCopper, apOrange, apGold, apTopaz, apYellow, _
  apLemon, apLime, apGreen, apJade, apTurquoise, apCyan, apAqua, apTeal, _
  apAzure, apSapphire, apBlue, apLavender, apViolet, apMagenta, apFuschia, _
  apPink, apCrimson)
  chn = Array("Red", "Copper", "Orange", "Gold", "Topaz", "Yellow", _
  "Lemon", "Lime", "Green", "Jade", "Turquoise", "Cyan", "Aqua", "Teal", _
  "Azure", "Sapphire", "Blue", "Lavender", "Violet", "Magenta", "Fuschia", _
  "Pink", "Crimson")
  cl = Array(apFaint, apPale, apLight, apSoft, apStd, apDense, apDeep, _
  apMurky, apDark, apPitch)
  cln = Array("Faint ", "Pale ", "Light ", "Soft ", "", "Dense ", "Deep ", _
  "Murky ", "Dark ", "Pitch ")
  bl = Array(apWhite, apSilver, apChrome, apTitanium, apGrey, apDarkGrey, _
  apEbony, apCharcoal, apPitchBlack, apBlack)
  bn = Array("White", "Silver", "Chrome", "Titanium", "Grey", "Dark Grey", _
  "Ebony", "Charcoal", "Pitch", "Black")
  mPopupSpace = 2
  mPopupSize = 9
  cHueCols = 24
  cHueRows = 10
  cExtended = False
Case ap8x4x4
  '(7 colors + 1 Achromatic) x 4 lum x 4 Sat = 128
  'Only hues configured here
  ch = Array(apRed, apOrange, apYellow, apGreen, apCyan, apBlue, apMagenta)
  chn = Array("Red", "Orange", "Yellow", "Green", "Cyan", "Blue", "Magenta")
  cl = Array(apStd)
  cln = Array("")
  bl = Array(apWhite)
  bn = Array("White")
  mPopupSpace = 3
  mPopupSize = 12
  cHueCols = 2
  cHueRows = 4
  cExtended = True
Case ap8x8x6
  '(7 colors + 1 Achromatic) x 8 lum x 6 Sat = 384
  'Only hues configured here
  ch = Array(apRed, apOrange, apYellow, apGreen, apCyan, apBlue, apMagenta)
  chn = Array("Red", "Orange", "Yellow", "Green", "Cyan", "Blue", "Magenta")
  cl = Array(apStd)
  cln = Array("")
  bl = Array(apWhite)
  bn = Array("White")
  mPopupSpace = 2
  mPopupSize = 11
  cHueCols = 1
  cHueRows = 8
  cExtended = True
Case ap16x8x6
  '(7 colors + 1 Achromatic) x 8 lum x 6 Sat = 384
  'Only hues configured here, reverse the normal order of 2nd rue to
  'blend them better
  ch = Array(apRed, apCopper, apOrange, apGold, apYellow, apLime, apGreen, apTurquoise, _
      apPink, apFuschia, apMagenta, apLavender, apBlue, apTeal, apCyan)
  chn = Array("Red", "Copper", "Orange", "Gold", "Yellow", "Lime", "Green", "Turquoise", _
      "Pink", "Fuschia", "Magenta", "Lavender", "Blue", "Teal", "Cyan")
  cl = Array(apStd)
  cln = Array("")
  bl = Array(apWhite)
  bn = Array("White")
  mPopupSpace = 2
  mPopupSize = 11
  cHueCols = 2
  cHueRows = 8
  cExtended = True
Case ap24x8x6
  '(23 colors + 1 Achromatic) x 8 lum x 6 Sat = 1104
  'Reverse the middle row to blend the colors better
  ch = Array(apRed, apCopper, apOrange, apGold, apTopaz, apYellow, apLemon, apLime, _
      apSapphire, apAzure, apTeal, apAqua, apCyan, apTurquoise, apJade, apGreen, _
      apBlue, apLavender, apViolet, apMagenta, apFuschia, apPink, apCrimson)
  chn = Array("Red", "Copper", "Orange", "Gold", "Topaz", "Yellow", "Lemon", "Lime", _
  "Sapphire", "Azure", "Teal", "Aqua", "Cyan", "Turquoise", "Jade", "Green", _
  "Blue", "Lavender", "Violet", "Magenta", "Fuschia", "Pink", "Crimson")
  cl = Array(apStd)
  cln = Array("")
  bl = Array(apWhite)
  bn = Array("White")
  mPopupSpace = 2
  mPopupSize = 10
  cHueCols = 3
  cHueRows = 8
  cExtended = True
Case ap12x6x6
  '(11 colors + 1 Achromatic) x 6 lum x 6 Sat = 432
  'Reverse the second row to blend the colors better
  ch = Array(apRed, apOrange, apYellow, apLemon, apGreen, apTurquoise, _
      apPink, apMagenta, apBlue, apTeal, apCyan)
  chn = Array("Red", "Orange", "Yellow", "Lemon", "Green", "Turquoise", _
      "Pink", "Magenta", "Blue", "Teal", "Cyan")
  cl = Array(apStd)
  cln = Array("")
  bl = Array(apWhite)
  bn = Array("White")
  mPopupSpace = 2
  mPopupSize = 11
  cHueCols = 2
  cHueRows = 6
  cExtended = True
Case ap32x8x6
  '(31 colors + 1 Achromatic) x 8 lum x 6 Sat = 1536
  'Reverse the middle row to blend the colors better
  ch = Array(apRed, apCopper, apOrange, apBronze, apGold, apTopaz, apTourmaline, apYellow, _
      apTurquoise, apJade, apBeryl, apGreen, apEmerald, apLime, apLemon, apCitrine, _
      apCyan, apAqua, apTeal, apAzure, apSapphire, apBlue, apIndigo, apLavender, _
      apCrimson, apPink, apFuschia, apMagenta, apCobalt, apViolet, apAmethyst)
  chn = Array("Red", "Copper", "Orange", "Bronze", "Gold", "Topaz", "Tourmaline", "Yellow", _
      "Turquoise", "Jade", "Beryl", "Green", "Emerald", "Lime", "Lemon", "Citrine", _
      "Cyan", "Aqua", "Teal", "Azure", "Sapphire", "Blue", "Indigo", "Lavender", _
      "Crimson", "Pink", "Fuschia", "Magenta", "Cobalt", "Violet", "Amethyst")
  cl = Array(apStd)
  cln = Array("")
  bl = Array(apWhite)
  bn = Array("White")
  mPopupSpace = 2
  mPopupSize = 10
  cHueCols = 4
  cHueRows = 8
  cExtended = True
End Select

nHue = (UBound(ch) + 1) * (UBound(cl) + 1) + UBound(bl)
nSat = 0
ReDim c(nHue) As ColorRect
ReDim d(nSat) As ColorRect
HSL.Sat = MaxHSL
For i = 0 To UBound(ch)
  HSL.Hue = ch(i)
  For j = 0 To UBound(cl)
    HSL.Lum = cl(j)
    c(k).ColorName = cln(j) & chn(i)
    c(k).Chromatic = True
    c(k).HSL = HSL
    c(k).Color = HSLtoRGB(HSL)
    c(k).IsHue = True
    c(k).Ok = True
    k = k + 1
  Next
Next
'Acrhomatics
HSL.Sat = 0
HSL.Hue = 0
For i = 0 To UBound(bl)
  HSL.Lum = bl(i)
  c(k).ColorName = bn(i)
  c(k).Chromatic = False
  c(k).HSL = HSL
  c(k).Color = HSLtoRGB(HSL)
  c(k).IsHue = True
  c(k).Ok = True
  k = k + 1
Next
'Arrange the rectangles. If priority is Down then NoItems=Rows else NoItems=Colums
ArrangeHueRect True, cHueRows, mPopupEdge, mPopupEdge

End Sub

Private Sub InitSatColorBoxes()
'-------------------------------------------------------
'Initialise Lum/Sat color rectangles (right hand side)
'-------------------------------------------------------
Dim HSL As HSLColor
Dim i, j, k, X, Y
Dim Sat As Variant      'Array containing saturation levels
Dim Lum As Variant      'Array containing luminescence levels
Dim Lnm As Variant      'Array containing Luminescence names
Dim Stp As Variant      'Luminescence step decrease for achromatic colors
Dim xOff, yOff

'Lum = Array(apFaint, apPale, apLight, apSoft, apStd, apDense, apDeep, apMurky, apDark, apPitch)
'Lnm = Array("Faint ", "Pale ", "Light ", "Soft ", "", "Dense ", "Deep ", "Murky ", "Dark ", "Pitch ")
'Lum = Array(apWhite, apSilver, apChrome, apTitanium, apGrey, apDarkGrey, apCharcoal, apBlack)
'Lnm = Array("White", "Silver", "Chrome", "Titanium", "Grey", "Dark Grey", "Charcoal", "Black")
If rcHue.Ok Then
  Select Case mPalette
  Case ap8x4x4
    Sat = Array(240, 180, 120, 60)
    If rcHue.Chromatic Then
      'Chromatic - colors
      Lum = Array(apPale, apLight, apStd, apDark)
      Lnm = Array("Pale ", "Light ", "", "Dark ")
    Else
      'Achromatic - grey scale
      Lum = Array(apWhite, apChrome, apGrey, apCharcoal)
      Lnm = Array("White", "Chrome", "Grey", "Charcoal")
      Stp = Array(15, 15, 15, 15)
    End If
    cSatCols = 4
    cSatRows = 4
  Case ap12x6x6
    Sat = Array(240, 200, 160, 120, 80, 40)
    If rcHue.Chromatic Then
      'Chromatic - colors
      Lum = Array(apPale, apLight, apStd, apDeep, apDark, apPitch)
      Lnm = Array("Pale ", "Light ", "", "Deep ", "Dark ", "Pitch ")
    Else
      'Achromatic - grey scale
      Lum = Array(apWhite, apPlatinum, apNickel, apGrey, apEbony, apMidnight)
      Lnm = Array("White", "Platinum", "Nickel", "Grey", "Ebony", "Midnight")
      Stp = Array(6, 6, 6, 6, 6, 6)
    End If
    cSatCols = 6
    cSatRows = 6
  Case ap8x8x6, ap16x8x6, ap24x8x6, ap32x8x6
    Sat = Array(240, 200, 160, 120, 80, 40)
    If rcHue.Chromatic Then
      'Chromatic - colors
      Lum = Array(apFaint, apPale, apLight, apStd, apDense, apDeep, apDark, apPitch)
      Lnm = Array("Faint ", "Pale ", "Light ", "", "Dense ", "Deep ", "Dark ", "Pitch ")
    Else
      'Achromatic - grey scale
      Lum = Array(apWhite, apSilver, apChrome, apTitanium, apGrey, apDarkGrey, apCharcoal, apPitchBlack)
      Lnm = Array("White", "Silver", "Chrome", "Titanium", "Grey", "Dark Grey", "Charcoal", "Pitch")
      Stp = Array(5, 5, 5, 5, 5, 5, 5, 5)
    End If
    cSatCols = 6
    cSatRows = 8
  End Select
  nSat = (UBound(Lum) + 1) * (UBound(Sat) + 1) - 1
  ReDim d(nSat) As ColorRect
  HSL.Hue = rcHue.HSL.Hue
  xOff = mPopupEdge + (mPopupSize + mPopupSpace) * cHueCols + 2 + mPopupSpace
  yOff = mPopupEdge
  For i = 0 To UBound(Sat)
    If rcHue.Chromatic Then
      HSL.Sat = Sat(i)
    Else
      HSL.Sat = 0
    End If
    For j = 0 To UBound(Lum)
      If rcHue.Chromatic Then
        d(k).ColorName = Lnm(j) & rcHue.ColorName
        HSL.Lum = Lum(j)
      Else
        If k < nSat Then
          d(k).ColorName = Lnm(j)
          HSL.Lum = Lum(j) - Stp(j) * i
        Else
          d(k).ColorName = "Black"
          HSL.Lum = 0
        End If
      End If
      'Save in color rectangle structure
      d(k).IsHue = False
      d(k).Ok = True
      d(k).Color = HSLtoRGB(HSL)
      d(k).HSL = HSL
      d(k).Chromatic = rcHue.Chromatic
      k = k + 1
    Next
  Next
  ArrangeSatRect True, cSatRows, xOff, yOff
End If

End Sub

Private Function MatchColor(ByVal cColor As Long) As ColorRect
'--------------------------------------------------------------
'Matches cColor with the hue() and sat() colors
'Return the ColorRect containing the closest match
'--------------------------------------------------------------
Dim i, Delta, MaxDelta, HSL As HSLColor, Hue, Sat, Lum
Dim rc As ColorRect

HSL = RGBtoHSL(cColor)
Hue = HSL.Hue
Sat = HSL.Sat
Lum = HSL.Lum
rc.Ok = False

'Match hues first - c()
MaxDelta = MaxHSL * 3  'Set the delta to its highest possible value
For i = 0 To nHue
  Delta = Abs(c(i).HSL.Hue - Hue) + Abs(c(i).HSL.Lum - Lum) + Abs(c(i).HSL.Sat - Sat)
  If Delta < MaxDelta Then
    rc = c(i)
    MaxDelta = Delta
  End If
Next

'If Delta indicates an imperfect match, search through lum/sat matrix d()
If MaxDelta > 0 And cExtended Then
  For i = 0 To nSat
    Delta = Abs(d(i).HSL.Hue - Hue) + Abs(d(i).HSL.Lum - Lum) + Abs(d(i).HSL.Sat - Sat)
    If Delta < MaxDelta Then
      rc = d(i)
      MaxDelta = Delta
    End If
  Next
End If
MatchColor = rc

End Function

Private Function MatchColorRect(ByVal X As Long, ByVal Y As Long) As ColorRect
'--------------------------------------------------------------
'Matches mouse position with the color rectangles
'When the match is found, return the appropriate rectangle
'If no match, set rc.Index=-1 and rc.Ok to false
'Uses the API PtInRect() function
'Note: Later versions of the Platform SDK API reference calls
'for the PtInRect(Rect,PointApi) parameters. VB6/Win98 still
'uses PtInRect(Rect,x,y) calling convention
'--------------------------------------------------------------
Dim rct As RECT
Dim i, Ok As Long

Ok = GetClientRect(picPopup.hWnd, rct)
If PtInRect(rct, X, Y) <> 0 Then
  'Match hues first - c()
  For i = 0 To nHue
    rct = c(i).rc
    If PtInRect(rct, X, Y) <> 0 Then
      MatchColorRect = c(i)
      Exit Function
    End If
  Next
  
  'If no point fount, search the sat/lum boxes
  If cExtended Then
    For i = 0 To nSat
      If PtInRect(d(i).rc, X, Y) <> 0 Then
        MatchColorRect = d(i)
        Exit Function
      End If
    Next
  End If
End If
MatchColorRect.Ok = False

End Function

Private Sub MatchExternalColor(ByVal cColor As Long)
'--------------------------------------------------------------
'Matches externally provided cColor with the hue() and sat() colors
'When the match is found, update the selected color rectangle
'and fire the Click event
'Notes:
'1) Always returns a match - the closest color
'--------------------------------------------------------------
Dim rc As ColorRect

If cExtended Then
  'Match Hues first in order to get the correct Saturation/Luminescence
  rcHue = MatchHue(cColor)
  InitSatColorBoxes
End If
'Match the Color
rc = MatchColor(cColor)
'Update the selected Color and raise the Click event, if changed
ChangeSelectedColor rc

End Sub

Private Function MatchHue(cColor As Long) As ColorRect
'--------------------------------------------------------------
'Matches the hue component of cColor with those
'in the primary array c() and returns the index no in c()
'A perfect match may not be possible, so therefore
'return the ColorRect having the closest match.
'--------------------------------------------------------------
Dim i, Delta, MaxDelta, HSL As HSLColor, Hue, Lum
Dim rc As ColorRect

HSL = RGBtoHSL(cColor)
Hue = HSL.Hue
Lum = HSL.Lum
rc.Ok = False
If HSL.Sat = 0 Then
  'Match achromatic (greyscale) luminescence values
  MaxDelta = MaxHSL    'Set the delta to its highest possible value
  For i = 0 To nHue
    If Not (c(i).Chromatic) Then
      Delta = Abs(c(i).HSL.Lum - Lum)
      If Delta < MaxDelta Then
        rc = c(i)
        MaxDelta = Delta
      End If
    End If
  Next
Else
  MaxDelta = MaxHSL    'Set the delta to its highest possible value
  For i = 0 To nHue
    If c(i).Chromatic Then
      Delta = Abs(c(i).HSL.Hue - Hue)
      If Delta < MaxDelta Then
        rc = c(i)
        MaxDelta = Delta
      End If
    End If
  Next
End If
MatchHue = rc

End Function

Private Function Max(t1 As Variant, ParamArray t() As Variant) As Variant
'----------------------------------------------------
'Determine the maximum of all values
'Any number can be given (minimum 2), in any datatype
'----------------------------------------------------
Dim X As Variant, i As Long

X = t1
For i = 0 To UBound(t)
  If t(i) > X Then
    X = t(i)
  End If
Next
Max = X

End Function

Private Sub PaintDropDown(Optional vDown As Variant)
'-----------------------------------------------------------------
'Paint the dropdown button
'State  : Normal (cDown=false), Down (cDown=true)
'         If cDown is omitted, use the previous state
'Width  : 13 pixels, including 2 pixel border
'Height : Normally 17 pixels (inside of edit box)
'         but changes with scaleheight
'-----------------------------------------------------------------
Static pDown As Boolean       'Previous state
Dim cDown As Boolean          'Current state
Dim pt As POINTAPI            'Not used, but required by API call
Dim c, Ok

If IsMissing(vDown) Then
  cDown = pDown
  Ok = True                   'Force a repaint
Else
  cDown = vDown
  Ok = cDown <> pDown         'Repaint only if state has changed
End If

If Ok Then
  'Get the rectangle area of the UsrCtrl and adjust to size
  'to get the rectangle of the dropdown button
  Ok = GetClientRect(UserControl.hWnd, rDrp)
  With rDrp
    .Top = .Top + 2
    .Right = .Right - 2
    .Left = .Right - mDropWidth
    .Bottom = .Bottom - 2
    c = .Bottom \ 2        'Center height
    If cDown Then
      '-----------------------------------
      'Button is in down (pressed) state
      '-----------------------------------
      'Draw the border
      Ok = DrawEdge(UserControl.hdc, rDrp, EDGE_SUNKEN, BF_RECT Or BF_FLAT)
      'Draw the face
      Line (.Left + 2, .Top + 2)-(.Right - 3, .Bottom - 3), vbButtonFace, BF
      'Draw triangle
      'Triangle is 3 lines high, 5 pixels first line, then 3, then 1
      'Remember that LineTo command does not draw last point
      UserControl.ForeColor = vbButtonText                  'Normally black
      'Ok = MoveToEx(UserControl.hDc, x, y, pt)             'Sample
      Ok = MoveToEx(UserControl.hdc, .Left + 5, c + 1, pt)  'Top line, left
      Ok = LineTo(UserControl.hdc, .Left + 10, c + 1)       'Top right
      Ok = MoveToEx(UserControl.hdc, .Left + 6, c + 2, pt)  'Mdl left
      Ok = LineTo(UserControl.hdc, .Left + 9, c + 2)        'Mdl right
      Ok = MoveToEx(UserControl.hdc, .Left + 7, c + 3, pt)  'Bot left
      Ok = LineTo(UserControl.hdc, .Left + 8, c + 3)        'Bot right
    Else
      '-----------------------------------
      'Button is in up (normal) state
      '-----------------------------------
      'Draw the border
      Ok = DrawEdge(UserControl.hdc, rDrp, EDGE_RAISED, BF_RECT)
      'Draw the face
      Line (.Left + 2, .Top + 2)-(.Right - 3, .Bottom - 3), vbButtonFace, BF
      'Draw triangle
      'Triangle is 3 lines high, 5 pixels first line, then 3, then 1
      'Remember that LineTo command does not draw last point
      'Triangle moves one pixel down & one to the right
      UserControl.ForeColor = vbButtonText                 'Normally black
      'Ok = MoveToEx(UserControl.hDc, x, y, pt)            'Sample
      Ok = MoveToEx(UserControl.hdc, .Left + 4, c, pt)     'Top line, left
      Ok = LineTo(UserControl.hdc, .Left + 9, c)           'Top right
      Ok = MoveToEx(UserControl.hdc, .Left + 5, c + 1, pt)    'Mdl left
      Ok = LineTo(UserControl.hdc, .Left + 8, c + 1)         'Mdl right
      Ok = MoveToEx(UserControl.hdc, .Left + 6, c + 2, pt)    'Bot left
      Ok = LineTo(UserControl.hdc, .Left + 7, c + 2)          'Bot right
    End If
  End With
End If
pDown = cDown

End Sub

Sub PaintMain()
'---------------------------------------------------------------------------
'Paint the main 'combobox'
'a) Paint 3D border
'b) Paint background/focus area
'c) Paint selected color image
'd) Add the current text.
'Note: The dropdown button is not painted here - see PaintDropDown()
'---------------------------------------------------------------------------
Dim cFocus As Boolean         'Current focus
Dim rct As RECT               'UsrCtrl rectangle
Dim H, w                      'UsrCtrl height & width
Dim pt As POINTAPI            'Not used, but required by API call
Dim TextColor As Long         'Saved to draw text
Dim Ok

On Error GoTo PaintMainErr
'Get environment info
H = ScaleHeight
w = ScaleWidth
cFocus = hasFocus()           'Update focus status
Cls                           'Clear Usercontrol

'----------------------------------------------
'Draw the control border
'Reduces client area with 2 pixels on all sides
'The API DrawEdge() function is used with
'EDGE_SUNKEN : the type of border (raised/sunken)
'BF_RECT     : the sides to draw  (all 4 sides)
'----------------------------------------------
'Get the rectangle area of the UsrCtrl
Ok = GetClientRect(UserControl.hWnd, rct)
'Draw the border
Ok = DrawEdge(UserControl.hdc, rct, EDGE_SUNKEN, BF_RECT)

'-----------------------------------------------------------
'Draw the background
'Focus rectangle is indicated by a filled rectangle
'similar to that of a text box with selected text
'Leave 1 pixel white (background) border
'When no focus, simply set the background color
'-----------------------------------------------------------
If cFocus Then
  'Draw the highlight rectangle, 1 pixel inside client area (border = 2 pixels)
  'Therefore, rectangle is 3 pixels from the outer edge
  Line (3, 3)-(w - mDropWidth - 4, H - 4), vbHighlight, BF
  'Set the color for text (usually white)
  TextColor = vbHighlightText
Else
  'Set the color for text, usually black
  TextColor = UserControl.ForeColor
End If

'--------------------------------------------------------------------
'Draw the selected color box
'Size   : 2 pixels white area, thus 1 pxl smaller than focus rectangle
'Border : 1 pxl wide, higlight top/left, shadow bottom/right
'Fill   : Current selected color
'Use the API/GDI functions MoveToEx(move pen) and LineTo(draw pen)
'Can't use the drawedge function here
'--------------------------------------------------------------------
'Draw 1 pixel border
UserControl.ForeColor = vb3DHighlight         'Change to highlight (not white!)
Ok = MoveToEx(UserControl.hdc, 4, H - 5, pt)  'Bottom Left corner
Ok = LineTo(UserControl.hdc, 4, 4)            'Top Left
Ok = LineTo(UserControl.hdc, H - 5, 4)        'Top Right
UserControl.ForeColor = vb3DShadow            'Change to dark shadow (not black!)
Ok = LineTo(UserControl.hdc, H - 5, H - 5)    'Bottom Right
Ok = LineTo(UserControl.hdc, 4, H - 5)        'Bottom Left

'Fill rectangle with current color
'The BF option does a flood fill in the given color
If picPopup.visible Then
  Line (5, 5)-(H - 6, H - 6), rcActive.Color, BF
Else
  Line (5, 5)-(H - 6, H - 6), rcSel.Color, BF
End If

'Draw the text in the selected TextColor, depending on the focus status
CurrentX = H - 1
CurrentY = 4
UserControl.ForeColor = TextColor
If picPopup.visible Then
  Print rcActive.ColorName
Else
  Print rcSel.ColorName
End If

'Paint the dropdown button too, as a CLS instruction was issued
'Do not adjust the button state
PaintDropDown

Exit Sub

PaintMainErr:
Debug.Print "PaintMainErr: ", Err, Error
Resume Next
End Sub

Private Sub PaintPopupHue()
'------------------------------------------------------
'Draw the hue color boxes (left hand side)
'------------------------------------------------------
Dim i

For i = 0 To nHue
  With c(i).rc
    'Draw box
    picPopup.FillColor = c(i).Color
    picPopup.Line (.Left, .Top)-(.Right, .Bottom), vbButtonFace, B
  End With
Next

End Sub

Private Sub PaintPopupSat()
'-----------------------------------------------------
'Draw the sat/lum color boxes (right hand side)
'-----------------------------------------------------
Dim i

For i = 0 To nSat
  With d(i).rc
    'Draw box
    picPopup.FillColor = d(i).Color
    picPopup.Line (.Left, .Top)-(.Right, .Bottom), vbButtonFace, B
  End With
Next

End Sub

Public Property Let Palette(NewPalette As ArielPalette)
Attribute Palette.VB_Description = "Sets/Returns the palette mode, which determines the no of hue, luminescence and saturation combinations"
Attribute Palette.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
'-------------------------------------------------------
'Change the palette
'-------------------------------------------------------
mPalette = NewPalette
InitHueColorBoxes
If cExtended Then
  'Match the hues
  rcHue = MatchHue(rcSel.Color)
  InitSatColorBoxes
End If
ResizePopupWindow
'Finally, match the selected color to one of the
'color rectangles and update the color rectangles
rcSel = MatchColor(rcSel.Color)
rcActive.Ok = False
rcCurr.Ok = False
PropertyChanged "Palette"
PaintMain

End Property

Public Property Get Palette() As ArielPalette
'-----------------------------------------------
'Expose the palette property
'-----------------------------------------------
Palette = mPalette

End Property

'MemberInfo=0,0,0,True
Public Property Get PopupEnabled() As Boolean
'--------------------------------------------------------
'Returns the current status of the PopupEnabled property
'When disabled, the popupwindo will not show, allowing
'the user to substitute his routine for color selection
'using the popup() event
'--------------------------------------------------------
PopupEnabled = mPopupEnabled

End Property

Public Property Let PopupEnabled(ByVal NewPopupEnabled As Boolean)
'--------------------------------------------------------
'Sets the current status of the PopupEnabled property
'When disabled, the popupwindo will not show, allowing
'the user to substitute his routine for color selection
'using the popup() event
'--------------------------------------------------------
mPopupEnabled = NewPopupEnabled
PropertyChanged "PopupEnabled"

End Property

Public Sub Refresh()
'----------------------------------------------
'Paint the object. Used mainly by external
'users. Not used internally
'----------------------------------------------
UserControl_Paint

End Sub

Private Sub ResizeCtrl()
'-------------------------------------------------------
'Resize the user control (excl popupwindow)
'Use the Busy flag to prevent recursive calls
'-------------------------------------------------------
Static Busy As Boolean
Dim H, w

If Not (Busy) Then
  Busy = True
  '--------------------------------------
  'Validate height of main box
  '--------------------------------------
  'Restrict to height of text + 2 pixels white space + 2 pxls border
  H = TextHeight("Dummy") + 4 * 2
  'Remember, UsrCtrl height may be in different scale modes,
  'depending on the container setting.
  'Therefore, scale from pixels to the appropriate size
  Height = ScaleY(H, vbPixels, vbContainerSize)
  '--------------------------------------
  'Validate width of main box
  '--------------------------------------
  'Same scaling applies to width
  w = ScaleX(Width, vbContainerSize, vbPixels)
  'Restrict minimum width
  If w < mMinWidth Then
    Width = ScaleX(mMinWidth, vbPixels, vbContainerSize)
  End If
  Busy = False
End If
picPopup.visible = False
 
End Sub

Private Sub ResizePopupWindow()
'-------------------------------------------------------------
'Resize the popup window
'-------------------------------------------------------------
Dim nRows

If cExtended Then
  'Both Hue & Sat color rectangles are shown, with an edge
  'in between
  nRows = Max(cHueRows, cSatRows)
  picPopup.Width = (mPopupEdge + (mPopupSize + mPopupSpace) * cHueCols + 2 + (mPopupSize + mPopupSpace) * cSatCols + mPopupEdge) * Screen.TwipsPerPixelX
  picPopup.Height = (mPopupEdge + (mPopupSize + mPopupSpace) * nRows + mPopupEdge - mPopupSpace) * Screen.TwipsPerPixelY
Else
  'Only the Hue color rectangles are shown, no vertical edge
  picPopup.Width = (mPopupEdge + (mPopupSize + mPopupSpace) * cHueCols + mPopupEdge - mPopupSpace) * Screen.TwipsPerPixelX
  picPopup.Height = (mPopupEdge + (mPopupSize + mPopupSpace) * cHueRows + mPopupEdge - mPopupSpace) * Screen.TwipsPerPixelY
End If

End Sub

Public Property Get SelectedColor() As OLE_COLOR
'------------------------------------------------------------
'Read propery : SelectedColor
'------------------------------------------------------------
SelectedColor = rcSel.Color

End Property

Public Property Let SelectedColor(ByVal NewSelectedColor As OLE_COLOR)
'------------------------------------------------------------
'Update the selected color and repaint the main box
'------------------------------------------------------------
If rcSel.Color <> NewSelectedColor Then
  MatchExternalColor NewSelectedColor
End If

End Property

Private Sub ShowPopUp()
'----------------------------------------------------------------
'Shows the Popup Window.
'a) Aligns the popup window to the edit box
'b) Paint the popup window
'c) Captures the mouse and cancel key
'd) Enables the hover timer
'----------------------------------------------------------------
Dim rCtl As RECT    'Usercontrol rectangle
Dim lLeft, lTop
Dim i

If Not (picPopup.visible) Then
  '-----------------------------------------------------------
  'Determine vertical position of popup window
  'Show the popup below the control, but if that can't be done
  'show it above
  '-----------------------------------------------------------
  'Get screen rectange of the control
  GetWindowRect UserControl.hWnd, rCtl
  'Remember the popup picture belongs to the screen
  If rCtl.Bottom + (picPopup.Height / Screen.TwipsPerPixelX) > Screen.Height / Screen.TwipsPerPixelY Then
    'Put it above
    lTop = (rCtl.Top - (picPopup.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
  Else
    'Put it below
    lTop = rCtl.Bottom * Screen.TwipsPerPixelY
  End If
  
  '-----------------------------------------------------------
  'Determine Horizontal position of popup window
  'If ctrl is wider than popup, align popup to right of ctrl.
  'Ensure that it is not off the right screen edge
  'If ctrl is narrower than popup, align to ctrl left
  '-----------------------------------------------------------
  If (rCtl.Right - rCtl.Left) > picPopup.Width / Screen.TwipsPerPixelX Then
    'Ctrl.width is wider than popup...align to ctrl.right
    If rCtl.Right > Screen.Width / Screen.TwipsPerPixelX Then
      'Ctrl is off right screen, so align popup with right screen edge
      lLeft = Screen.Width - picPopup.Width
    Else
      'Ctrl is not off right edge, so align with ctrl.right
      lLeft = rCtl.Right * Screen.TwipsPerPixelX - picPopup.Width
    End If
    'Check that position is not outside screen left edge
    If lLeft < 0 Then lLeft = 0
  Else
    'Ctrl.width is smaller than popup.width, so align to ctrl.left
    If rCtl.Left < 0 Then
      lLeft = 0
    Else
      lLeft = rCtl.Left * Screen.TwipsPerPixelX
    End If
    'Check we haven't gone outside screen right edge
    If lLeft + picPopup.Width > Screen.Width Then
      lLeft = Screen.Width - picPopup.Width
    End If
  End If
  
  '----------------------------------------
  'Set popup position, put on top and show
  '----------------------------------------
  With picPopup
    .Top = lTop
    .Left = lLeft
    .visible = True
    .ZOrder
  End With
  picPopup_Paint
  'DoEvents
  'UserControl_Paint
  
  '------------------------------------------------------------
  'Handle potential errors
  '------------------------------------------------------------
  'Capture the mouse so we get all subsequent mouse clicks
  SetCapture picPopup.hWnd
  'Store the 'Cancel' control so we stop Escape from firing the
  'default 'Cancel' button. This is restored on exit
  'Debug.Print GetCapture
  On Error Resume Next
  For i = 0 To UserControl.ParentControls.Count - 1
    If UserControl.ParentControls(i).Cancel Then
      If Err = False Then
        Set ctlCancel = UserControl.ParentControls(i)
        ctlCancel.Cancel = False
        Exit For
      End If
      Err = False
    End If
  Next
  'Enable Hover Timer - used to autoselect extended colors
  If cExtended Then
    tmrHover.Enabled = True
  End If
  On Error GoTo 0
End If

End Sub

Public Property Get Text() As String
'-------------------------------------------
'Read the text property containing the
'name of the selected color
'-------------------------------------------
Text = rcSel.ColorName

End Property

Private Sub picPopup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------------
'Handle mouse down events
'a) If Mousepointer within popwindow, select new colorbox
'b) If not, hide popupwindow
'-----------------------------------------------------------------------------
Dim rct As RECT, Ok, lx, ly
Dim rc As ColorRect

'Determine if mouse is within popup window boundaries
Ok = GetClientRect(picPopup.hWnd, rct)
lx = CLng(X)
ly = CLng(Y)
If PtInRect(rct, lx, ly) = 0 Then  'Returns 1 if true, 0 if false
  'Mouse is outside borders of popup window
  HidePopUp
Else
  'Mouse is inside borders, so check if mouse is over
  'a color rectangle.
  'Restore the border of the previous rectangle to normal state
  Call DrawRectBorder(rcCurr, False, False)
  'Find the match. Returns -1 if no rectangle found
  rc = MatchColorRect(lx, ly)
  If rc.Ok Then
    'A new color was found, update the Curr color
    rcCurr = rc
    'Update the Active color (since the mouse is down) and raise
    'the Change() event, if it has changed
    ChangeActiveColor rc
    'Repaint the sat/lum boxes if the Hue has changed
    If rc.IsHue And rc.HSL.Hue <> rcHue.HSL.Hue And cExtended Then
      rcHue = rc
      InitSatColorBoxes
      PaintPopupSat
    End If
    'Draw the new rectangle border (after the PaintPopupSat() routine!)
    Call DrawRectBorder(rcCurr, True, True)
  End If
End If

End Sub

Private Sub picPopup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------------
'Match the mouse position to a color rectangle
'If found, highlight the rectangle
'If the mouse is down, update the Active Color
'---------------------------------------------------------------------------
Dim cMouseDown As Boolean
Dim rc As ColorRect

cMouseDown = (Button And vbLeftButton) > 0
'Restore the previous current color rectangle
Call DrawRectBorder(rcCurr, False, False)

'Find the match
rc = MatchColorRect(X, Y)
If rc.Ok Then
  'Update the current color rectangle
  rcCurr = rc
  'Set the new active color, only if mousedown
  If cMouseDown Then
    ChangeActiveColor rc
  End If
  'Paint the sat/lum boxes (only if mousedown)
  'This eliminates the hover time delay
  If cMouseDown Then
    If rc.IsHue And rc.HSL.Hue <> rcHue.HSL.Hue And cExtended Then
      rcHue = rc
      InitSatColorBoxes
      PaintPopupSat
    End If
  End If
  'Draw the new rectangle border
  Call DrawRectBorder(rcCurr, True, cMouseDown)
End If
'Indicate that mouse movement has occurred
'This is tested by the hover timer
'When the hover timer fires and determines that there has been
'no mouse movement, the sat/lum boxes will be updated with the
'latest hue color
cMouseMove = True

End Sub

Private Sub picPopup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------------
'If the mouse is outside the popup window border, hide the popup
'If inside the border, match the position to a color rectangle
'If a match is found, update the current selected color and hide the popup
'If no match, do not hide the popup
'---------------------------------------------------------------------------
Dim rct As RECT, Ok, lx, ly
Dim rc As ColorRect

'Determine if mouse is within popup window boundaries
lx = CLng(X)
ly = CLng(Y)
Ok = GetClientRect(picPopup.hWnd, rct)
If PtInRect(rct, lx, ly) = 0 Then  'Returns 1 if true, 0 if false
  'Mouse is outside borders of popup window, so hide the popup
  HidePopUp
  'Restore the selected color
  ChangeSelectedColor rcSel
Else
  'Find the match
  rc = MatchColorRect(X, Y)
  If rc.Ok Then
    'Change the selected color and hide the popup window
    HidePopUp
    ChangeSelectedColor rc
  End If
End If

End Sub

Private Sub picPopup_Paint()
'----------------------------------------------------
'Paint the popup screen
'----------------------------------------------------
Dim i, rct As RECT, RctEdge As RECT, Ok
Dim pt As POINTAPI            'Not used, but required by API call

'Get the rectangle area of the popupwindow
Ok = GetClientRect(picPopup.hWnd, rct)
'Draw the popup window 3d border
Ok = DrawEdge(picPopup.hdc, rct, EDGE_RAISED, BF_RECT)

'Paint the hue & sat boxes
PaintPopupHue

If cExtended Then
  'Draw the vertical edge between hue and sat/lum boxes
  RctEdge.Top = rct.Top + mPopupEdge
  RctEdge.Bottom = rct.Bottom - mPopupEdge
  RctEdge.Left = rct.Left + mPopupEdge + (mPopupSize + mPopupSpace) * cHueCols + 1
  RctEdge.Right = RctEdge.Left   'Dummy, not used here
  Ok = DrawEdge(picPopup.hdc, RctEdge, EDGE_ETCHED, BF_LEFT)
  'Paint the extended Sat Color Rectangles
  PaintPopupSat
End If

End Sub

Private Sub tmrHover_OnTimer()
'-----------------------------------------------------------
'Select new hue and paint the sat/lum boxes
'This is only done after no mousemovement has taken place
'since the previous timer event, and then only if the
'current color index is a hue rectangle (left hand side)
'-----------------------------------------------------------
If cExtended Then
  If Not (cMouseMove) And rcCurr.IsHue Then
    rcHue = rcCurr
    InitSatColorBoxes
    If picPopup.visible Then
      PaintPopupSat
    End If
  End If
  cMouseMove = False
End If

End Sub

Private Sub UserControl_Click()
'-----------------------------------------------------
'Show popup window. A 2nd click will hide it
'-----------------------------------------------------
If mPopupEnabled Then
  If Not picPopup.visible Then
    ShowPopUp
  Else
    HidePopUp
  End If
Else
  If picPopup.visible Then
    HidePopUp
  End If
  RaiseEvent Popup
End If

End Sub

Private Sub UserControl_DblClick()
'--------------------------------------
'Expose event to user
'--------------------------------------
RaiseEvent DblClick

End Sub

Private Sub UserControl_EnterFocus()
'--------------------------------------------------
'Repaint the edit box to update the focus
'--------------------------------------------------
UserControl_Paint

End Sub

Private Sub UserControl_ExitFocus()
'--------------------------------------------------------------
'a) Repaint the edit box to update the focus
'b) Make sure that the popup window is hidden
'--------------------------------------------------------------
'Although in most circumstances the popup window will have
'already been hidden before this, we check here just in case.
If picPopup.visible Then
  HidePopUp
End If
'Hide the focus rectangle
UserControl_Paint

End Sub

Private Sub UserControl_Initialize()
'----------------------------------------------------------------------
'Initialize the control
'This occurs when
'1) a new instance is placed on the form in design mode
'2) when the developer/user runs the program
'----------------------------------------------------------------------

'Change the properties of the popupwindow
'a) Set the parent of the popup picturebox to the desktop
'b) Set style to Toolwindow so that the popup doesn't show in the Taskbar
SetWindowLong picPopup.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
SetParent picPopup.hWnd, 0

'Set the scalemode to pixels
UserControl.ScaleMode = vbPixels
picPopup.ScaleMode = vbPixels

'Initialize the Hovertimer
Set tmrHover = New ArielTimer
tmrHover.Interval = 400
'----------------------------------------------
tmrHover.Enabled = False
'----------------------------------------------
tmrHover.Name = "Hover Timer"

End Sub

Private Sub UserControl_InitProperties()
'-----------------------------------------
'Initialise properties
'-----------------------------------------
'Properties
Set UserControl.Font = Ambient.Font
mPalette = mdefPalette
mPopupEnabled = mdefPopupEnabled

'Initialize Hue color boxes - c()
'Also sets the cExtended variable
InitHueColorBoxes
If cExtended Then
  'Match the hues
  rcHue = MatchHue(mdefSelectedColor)
  'Initialise the Sat/Lum color boxes
  InitSatColorBoxes
End If
ResizePopupWindow

'Initialise the selected and active colors
rcSel = MatchColor(mdefSelectedColor)
rcActive.Ok = False
rcCurr.Ok = False

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------
'Capture [Space] to show popup, and [Esc] to hide
'--------------------------------------------------------
'Keypreview is set, so we get all of the keypresses here first.
'Since this is a button, we show the popup if the user presses
'Space and hide it if the user presses escape.
If KeyCode = vbKeySpace And (Shift = 0) Then
  UserControl_Click
ElseIf KeyCode = vbKeyEscape And picPopup.visible Then
  HidePopUp
End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------------
'No matter where the UserCtrl is clicked, emulate the Dropdown button
'being clicked
'--------------------------------------------------------------------------
PaintDropDown True

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------
'Control the state of DropDown button. To show the dropdown button in
'the 'down' state, two conditions must be fulfilled:-
'a) The left mouse button must be pressed (down)
'b) The mouse must be over the button
'In all other cases, the dropdown button is set to the 'Up' state
'Don't paint if state has not changed
'----------------------------------------------------------------------
Dim cDown As Boolean        'New state

If Button = vbLeftButton Then
  cDown = PtInRect(rDrp, X, Y) <> 0   'Returns 1 (inside) or 0 (outside)
Else
  cDown = False
End If
PaintDropDown cDown

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------------
'Set the dropdown button state to 'Up'
'--------------------------------------------------------------------------
PaintDropDown False

End Sub

Private Sub UserControl_Paint()
'-----------------------------------------------------------
'Repaint the edit box incl selected color, focus rectangle,
'text and dropdown button. The popupwindow is NOT repainted
'-----------------------------------------------------------
PaintMain

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-------------------------------------------------------------------------
'Load property values from storage
'-------------------------------------------------------------------------
Dim mSelectedColor As OLE_COLOR

Set Font = PropBag.ReadProperty("Font", Nothing)
UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
mPalette = PropBag.ReadProperty("Palette", mdefPalette)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
mPopupEnabled = PropBag.ReadProperty("PopupEnabled", mdefPopupEnabled)

'Match the hue of the selected color and initialise the
'lum/sat boxes accordingly
mSelectedColor = PropBag.ReadProperty("SelectedColor", mdefSelectedColor)
'Initialize Hue color boxes - c()
'Also sets the cExtended variable
InitHueColorBoxes
If cExtended Then
  'Match the hues
  rcHue = MatchHue(mSelectedColor)
  InitSatColorBoxes
End If
ResizePopupWindow
'Finally, match the selected color to one of the
'color rectangles and update the color rectangles
rcSel = MatchColor(mSelectedColor)
rcActive.Ok = False
rcCurr.Ok = False

End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------
'Adjust the size of constituent ctrls
'---------------------------------------------------------
On Error Resume Next
ResizeCtrl

End Sub

Private Sub UserControl_Terminate()
'--------------------------------------------------------------
'Destroy objects
'--------------------------------------------------------------
'It is important to destroy the hover timer to prevent it
'from continuously raising the timer event and to allow
'the user control to terminate
Set tmrHover = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-------------------------------------------------------
'Write property values to storage
'-------------------------------------------------------
Call PropBag.WriteProperty("Font", Font, Nothing)
Call PropBag.WriteProperty("SelectedColor", rcSel.Color, mdefSelectedColor)
Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("Palette", mPalette, mdefPalette)
Call PropBag.WriteProperty("PopupEnabled", mPopupEnabled, mdefPopupEnabled)

End Sub



'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
'------------------------------------------
'Expose the hWnd property of the Usercontrol
'------------------------------------------
hWnd = UserControl.hWnd

End Property

