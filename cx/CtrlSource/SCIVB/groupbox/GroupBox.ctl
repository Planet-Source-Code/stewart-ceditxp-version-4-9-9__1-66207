VERSION 5.00
Begin VB.UserControl GroupBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   HitBehavior     =   0  'None
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
End
Attribute VB_Name = "GroupBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As String) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As String, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Type Size
        cx As Long
        cy As Long
End Type

 Private Const BP_GROUPBOX = 4

Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2

Private Const DSS_DISABLED = &H20
Private Const DSS_NORMAL = &H0

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_GETKEYBOARDCUES = &H100A
Private Const DT_HIDEPREFIX = &H100000

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Enum GROUPBOXBORDERSTYLE
    None
    [Fixed Single]
End Enum

Public Enum GROUPBOXBACKSTYLE
    Transparent
    Solid
End Enum

Public Enum GROUPBOXAPPEARANCE
    Flat
    [3D]
End Enum

'Default Property Values:
Const m_def_Appearance = 1
Const m_def_BorderStyle = 1
'Property Variables:
Dim m_Appearance As Integer
Dim m_Caption As String
Dim m_BorderStyle As GROUPBOXBORDERSTYLE
Dim m_Enabled As Boolean
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type
Private Declare Function DllGetVersion Lib "Comctl32" (pdvi As DLLVERSIONINFO) As Long

Public Function IsThemedXP() As Boolean

    'Declare structure.
    Dim osVer As OSVERSIONINFO
    
    'Set size of structure.
    osVer.dwOSVersionInfoSize = Len(osVer)
    
    'Fill structure with data.
    GetVersionEx osVer
    
    'Evaluate return. If greater than or equal to 5.1 then running
    'WindowsXP or newer.
    If osVer.dwMajorVersion + osVer.dwMinorVersion / 10 >= 5.1 Then
        'Check for Active Visual Style(modified as per paravoid's suggestion).
        If IsAppThemed Then
            'Double Check by assessing DLL version loaded
            Dim dllVer As DLLVERSIONINFO
            dllVer.cbSize = Len(dllVer)
            DllGetVersion dllVer
            IsThemedXP = (dllVer.dwMajor >= 6)
        End If
    End If
    
End Function



Private Sub UserControl_AmbientChanged(PropertyName As String)
    DrawFrame
End Sub

Private Sub UserControl_Resize()
    DrawFrame
End Sub

Private Sub DrawFrame()

    Dim hTheme As Long
    Dim R As RECT
    Dim rText As RECT
    Dim siText As Size
    Dim bTheme As Boolean
    Dim DSFlag As Long
    Dim sCaption As String
    Dim sAccelCaption As String
    Dim b_Accel As Boolean
    
    'Clear the controls image.
    Cls
    
    'Draw Nothing if not requested
    If BorderStyle = None Then Exit Sub
    
    'Standard frame control sets top edge of box at half height of font.
    'So if there is no caption we must create a string to get this height.
    sCaption = IIf(Len(Caption) = 0, " ", Caption)
    
    'Get the width and height of the Caption
    If InStr(sCaption, "&") Then
        '"&&" is shown as "&" in the caption so we need to account for this.
        '`A` seems to be the same width as `&`, so since we only want one
        'occurance of "&" in our string it makes a good replacement.
        sAccelCaption = Replace(sCaption, "&&", "A")
        'Now we need to add the accelerator key to the control
        If InStr(sAccelCaption, "&") Then
            UserControl.AccessKeys = Mid(sAccelCaption, InStr(sAccelCaption, "&") + 1, 1)
        End If
        'Accelerator is not counted in the string length.
        sAccelCaption = Replace(sAccelCaption, "&", "", 1, InStr(sAccelCaption, "&"))
        GetTextExtentPoint32 hdc, sAccelCaption, Len(sAccelCaption), siText
    Else
        GetTextExtentPoint32 hdc, sCaption, Len(sCaption), siText
    End If
    
    'Set the frame boundry
    SetRect R, 0, siText.cy / 2, ScaleWidth, ScaleHeight
    
    'Set the width and height of the Caption Box
    SetRect rText, 9, 0, 9 + siText.cx, siText.cy
    
    'Check for applied Visual Style
    Select Case IsThemedXP
        'If visual styles are not being used
        Case False
            Select Case m_Appearance
                Case 0  'When changing appearance to Flat MS's Frame Control draws
                        'a black box on top of a wider white box. This is so that
                        'you can see the frame no matter what color you set the
                        'background. The Backcolor property is also changed to
                        'Window Background, but I'm not doing that because it bugs
                        'the hell out of me.
                    'Draw a 2 pixel wide White box
                    DrawWidth = 2
                    Line (R.Left + 1, R.Top + 1)-(R.Right - 1, R.Bottom - 1), vbWhite, B
                    'Draw a 1 pixel wide Black box
                    DrawWidth = 1
                    Line (R.Left, R.Top)-(R.Right - 1, R.Bottom - 1), vbBlack, B
                Case Else
                    'Draw an Etched box
                    DrawEdge hdc, R, EDGE_ETCHED, BF_RECT
            End Select
            'Erase/Draw the box where the text is to be displayed
            If Len(Caption) > 0 Then Line (7, 0)-(9 + siText.cx, siText.cy), UserControl.BackColor, BF
            'Draw the Caption either Enabled or Disabled.
            DSFlag = IIf(Enabled, DSS_NORMAL, DSS_DISABLED) Or DST_PREFIXTEXT
            DrawStateText hdc, 0&, 0&, Caption, Len(Caption), 9, 0, siText.cx, siText.cy, DSFlag
        'If using visual styles
        Case True
            'Open the Theme Data for the Button Class.
            hTheme = OpenThemeData(hWnd, StrConv("BUTTON", vbUnicode))
            'Draw a Groupbox Frame
            DrawThemeBackground hTheme, hdc, BP_GROUPBOX, 0&, R, ByVal 0&
            'Erase/Draw the box where the text is to be displayed
            If Len(Caption) > 0 Then Line (7, 0)-(9 + siText.cx, siText.cy), UserControl.BackColor, BF
            'Find out if Accelerator's are hidden.
            SystemParametersInfo 0&, SPI_GETKEYBOARDCUES, b_Accel, 0&
            'Draw the text in current Visual style
            DSFlag = IIf(Enabled, DSS_NORMAL, DSS_DISABLED)
            If b_Accel Then
                DrawThemeText hTheme, hdc, BP_GROUPBOX, 1, StrConv(Caption, vbUnicode), Len(Caption), 0&, DSFlag, rText
            Else
                DrawThemeText hTheme, hdc, BP_GROUPBOX, 1, StrConv(Caption, vbUnicode), Len(Caption), DT_HIDEPREFIX, DSFlag, rText
            End If
            'Close Theme Data
            CloseThemeData hTheme
    End Select
    
    'Set the Control's Mask
    Set MaskPicture = Image
    MaskColor = BackColor
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."" \r\n"
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    If Ambient.UserMode Then UserControl.Enabled = m_Enabled
    DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As GROUPBOXBORDERSTYLE
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As GROUPBOXBORDERSTYLE)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
    DrawFrame
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_BorderStyle = m_def_BorderStyle
    m_Caption = Ambient.DisplayName
    m_Appearance = m_def_Appearance
    Enabled = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    If Ambient.UserMode Then UserControl.Enabled = m_Enabled
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    
End Sub

Private Sub UserControl_Show()
    DrawFrame
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As GROUPBOXBACKSTYLE
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As GROUPBOXBACKSTYLE)
    If New_BackStyle = Transparent Then
        MsgBox "This feature has not yet been implemented!" & vbCrLf & _
                "I'm not quite sure how I am going to implement this feature" & _
                ", as some controls are shaped.", _
                vbOKOnly
    Else
        UserControl.BackStyle() = New_BackStyle
        PropertyChanged "BackStyle"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Appearance() As GROUPBOXAPPEARANCE
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As GROUPBOXAPPEARANCE)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    DrawFrame
End Property

