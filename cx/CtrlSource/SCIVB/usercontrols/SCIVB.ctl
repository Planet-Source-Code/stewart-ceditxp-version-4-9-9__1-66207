VERSION 5.00
Begin VB.UserControl SCIVB 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   KeyPreview      =   -1  'True
   Picture         =   "SCIVB.ctx":0000
   ScaleHeight     =   555
   ScaleWidth      =   525
   ToolboxBitmap   =   "SCIVB.ctx":0731
End
Attribute VB_Name = "SCIVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Private SCI As Long

Private fWindowProc As Long   ' Proc Address of Scintilla.

Private iSCISet As Integer    ' Generic way to see if Scintilla's set or not

Private SC As cSubclass     ' Subclass for Scintilla Messages
Private hWndParent As Long



Private lSCI As Long
Private m_hMod As Long

Private chStore As Long

Public DirectSCI As New cDirectSCI


Private APIStringLoaded As Boolean
Private APIStrings() As String
Private ActiveCallTip As Integer

' EOL Style Enum  (Scintilla supports Windows, Linux and Mac Line Endings)
Public Enum EOLStyle
  SC_EOL_CRLF = 0                     ' CR + LF
  SC_EOL_CR = 1                       ' CR
  sc_eol_lf = 2                       ' LF
End Enum

' Edge Style Enum (This is for a column edge)
Public Enum edge
  EdgeNone = 0
  EdgeLine = 1
  EdgeBackground = 2
End Enum

' Word wrap style Enum (Word wrap can be based on none, character or word)
Public Enum WrapStyle
  noWrap = 0
  wrap = 1
  WrapChar = 2
End Enum

' Macro Type.  This the array of information recorded while
' macro recording is on.
Public Type MacroType
  lMsg As Long
  strChar As String
End Type
Private Macro() As MacroType

' Folding Style Enum (Folding can draw a box, arrow, circle, or Plus/Minus)
Public Enum FoldingStyle
  FoldMarkerArrow = 0
  foldMarkerBox = 1
  FoldMarkerCircle = 2
  FoldMarkerPlusMinus = 3
End Enum

' Gutter Type Enum (Using a symbol or linenumber gutter style.)
Public Enum GutterType
  GutSymbol = 0
  GutLineNumber = 1
End Enum

Event OnError(Number As String, Description As String)
Event KeyDown(KeyCode As Long, Shift As Long)
Event KeyUp(KeyCode As Long, Shift As Long)
Event KeyPress(Char As Long)
Event FindFailed(FindText As String)        'Find failed
Event StyleNeeded(Position As Long)                         'Style Needed Event
Event CharAdded(Char As Long)                            'A Character was added
Event SavePointReached()                    'No longer Modified
Public Event MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Event SavePointLeft()                       'File is now modified
Event ModifyAttemptRO()
' # GTK+ Specific to work around focus and accelerator problems:
Event Key(ch As Long, modifiers As Long)    'Key was pressed
Event DoubleClick()                         'Double clicked Scintilla
Event UpdateUI()                            'Something has been updated
Event OnModified(Position As Long, modificationType As Long)  'Modified
Event MacroRecord(message As Long, wParam As Long)          'Record Macro
Event MarginClick(modifiers As Long, Position As Long)      'Clicked Margin
Event NeedShown(Position As Long, length As Long)
Event Painted()                             'Painted the display
Event PosChanged(Position As Long)                          'Changed Position (Update Status)
Event UserListSelection(listType As Long, Text As String)   'Selected AutoComplete
Event URIDropped(Text As String)
Event DwellStart(Position As Long)
Event DwellEnd(Position As Long)
Event Zoom()                                'Zoom level has changed
Event HotSpotClick(modifiers As Long, Position As Long)     'Clicked a hotspot
Event HotSpotDoubleClick(modifiers As Long, Position As Long)   'Doubleclicked a hotspot
Event CallTipClick(Position As Long)                         'Clicked a calltip
Event AutoCSelection(Text As String)                      'Auto Completed selected

'Default Property Values:
Const m_def_BraceMatchBold = 1
Const m_def_BraceMatchItalic = 0
Const m_def_BraceMatchUnderline = 0
Const m_def_BraceMatchBack = vbWhite
Const m_def_BraceBadBack = vbWhite
Const m_def_BraceMatch = vbBlue
Const m_def_BraceBad = vbRed
Const m_def_SelStart = 0
Const m_def_SelEnd = 0
Const m_def_IndentationGuide = 1
Const m_def_FoldAtElse = 0
Const m_def_AutoCompleteStart = "."
Const m_def_AutoCompleteOnCTRLSpace = True
Const m_def_AutoCompleteString = "if then else"
Const m_def_AutoShowAutoComplete = 0
Const m_def_ContextMenu = 1
Const m_def_IgnoreAutoCompleteCase = 1
Const m_def_LineNumbers = 0
Const m_def_ReadOnly = 0
Const m_def_ScrollWidth = 2000
Const m_def_ShowFlags = 1
Const m_def_Text = "0"
Const m_def_SelText = "0"
Const m_def_MarginFore = vbBlack
Const m_def_MarginBack = &HE0E0E0
Const m_def_FoldMarker = 2
Const m_def_AutoCloseBraces = 0
Const m_def_AutoCloseQuotes = 0
Const m_def_BraceHighlight = 1
Const m_def_CaretForeColor = 0
Const m_def_LineBackColor = vbYellow
Const m_def_LineVisible = 0
Const m_def_CaretWidth = 1
Const m_def_FoldComment = True
Const m_def_FoldCompact = False
Const m_def_FoldHTML = True
Const m_def_ClearUndoAfterSave = 1
Const m_def_BookmarkBack = vbBlack
Const m_def_BookMarkFore = vbWhite
Const m_def_FoldHi = 0
Const m_def_FoldLo = 0
Const m_def_MarkerBack = vbBlack
Const m_def_MarkerFore = vbWhite
Const m_def_SelBack = &HFFC0C0
Const m_def_SelFore = vbBlack
Const m_def_EndAtLastLine = 0
Const m_def_OverType = 0
Const m_def_ScrollBarH = 1
Const m_def_ScrollBarV = 1
Const m_def_ViewEOL = 0
Const m_def_ViewWhiteSpace = 0
Const m_def_ShowCallTips = 1
Const m_def_EdgeColor = &HE0E0E0
Const m_def_EdgeColumn = 0
Const m_def_EdgeMode = 0
Const m_def_EOL = 0
Const m_def_Folding = 1
Const m_def_Gutter0Type = 1
Const m_def_Gutter0Width = 32
Const m_def_Gutter1Type = 0
Const m_def_Gutter1Width = 16
Const m_def_Gutter2Type = 0
Const m_def_Gutter2Width = 20
Const m_def_MaintainIndentation = 1
Const m_def_TabIndents = 1
Const m_def_BackSpaceUnIndents = 0
Const m_def_IndentWidth = 4
Const m_def_UseTabs = 0
Const m_def_WordWrap = 0
'Property Variables:
Dim m_BraceMatchBold As Boolean
Dim m_BraceMatchItalic As Boolean
Dim m_BraceMatchUnderline As Boolean
Dim m_BraceMatchBack As OLE_COLOR
Dim m_BraceBadBack As OLE_COLOR
Dim m_BraceMatch As OLE_COLOR
Dim m_BraceBad As OLE_COLOR
Dim m_SelStart As Long
Dim m_SelEnd As Long
Dim m_IndentationGuide As Boolean
Dim m_FoldAtElse As Boolean
Dim m_FoldComment As Boolean
Dim m_FoldCompact As Boolean
Dim m_FoldHTML As Boolean
Dim m_AutoCompleteStart As String
Dim m_AutoCompleteOnCTRLSpace As Boolean
Dim m_AutoCompleteString As String
Dim m_AutoShowAutoComplete As Boolean
Dim m_ContextMenu As Boolean
Dim m_IgnoreAutoCompleteCase As Boolean
Dim m_LineNumbers As Boolean
Dim m_ReadOnly As Boolean
Dim m_ScrollWidth As Long
Dim m_ShowFlags As Boolean
Dim m_Text As String
Dim m_SelText As String
Dim m_MarginFore As OLE_COLOR
Dim m_MarginBack As OLE_COLOR
Dim m_FoldMarker As FoldingStyle
Dim m_AutoCloseBraces As Boolean
Dim m_AutoCloseQuotes As Boolean
Dim m_BraceHighlight As Boolean
Dim m_CaretForeColor As OLE_COLOR
Dim m_LineBackColor As OLE_COLOR
Dim m_LineVisible As Boolean
Dim m_CaretWidth As Long
Dim m_ClearUndoAfterSave As Boolean
Dim m_BookmarkBack As OLE_COLOR
Dim m_BookMarkFore As OLE_COLOR
Dim m_FoldHi As OLE_COLOR
Dim m_FoldLo As OLE_COLOR
Dim m_MarkerBack As OLE_COLOR
Dim m_MarkerFore As OLE_COLOR
Dim m_SelBack As OLE_COLOR
Dim m_SelFore As OLE_COLOR
Dim m_EndAtLastLine As Boolean
Dim m_OverType As Boolean
Dim m_ScrollBarH As Boolean
Dim m_ScrollBarV As Boolean
Dim m_ViewEOL As Boolean
Dim m_ViewWhiteSpace As Boolean
Dim m_ShowCallTips As Boolean
Dim m_EdgeColor As OLE_COLOR
Dim m_EdgeColumn As Long
Dim m_EdgeMode As edge
Dim m_EOL As EOLStyle
Dim m_Folding As Boolean
Dim m_Gutter0Type As GutterType
Dim m_Gutter0Width As Long
Dim m_Gutter1Type As GutterType
Dim m_Gutter1Width As Long
Dim m_Gutter2Type As GutterType
Dim m_Gutter2Width As Long
Dim m_MaintainIndentation As Boolean
Dim m_TabIndents As Boolean
Dim m_BackSpaceUnIndents As Boolean
Dim m_IndentWidth As Long
Dim m_UseTabs As Boolean
Dim m_WordWrap As WrapStyle

Private bRegEx As Boolean
Private bWholeWord As Boolean
Private m_matchBraces
Private m_curHigh As String
Private bWrap As Boolean
Private bWordStart As Boolean
Private bCase As Boolean
Private strFind As String
Private bFindEvent As Boolean
Private bFindInRange As Boolean
Private bFindReverse As Boolean
Private bShowCallTips As Boolean
Private bShowFlags As Boolean
Private strAutoComplete As String
Private strAutoCompleteStart As String
Private bShowAutoComplete As Boolean

Private bRepLng As Boolean
Private bRepAll As Boolean


Implements iSubclass 'WinSubHook2.iSubclass

'impliments         ' iSuperclass provides the subclassing
                              ' Credit to: Paul Caton


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoCloseBraces() As Boolean    'When this is set to true braces <B>{, [, (</b> will be closed automatically.
    AutoCloseBraces = m_AutoCloseBraces
End Property

'Purpose: Auto Closes Braces
'Remarks: When this property is set to true (, [, and < will automaticly.
Public Property Let AutoCloseBraces(ByVal New_AutoCloseBraces As Boolean)
    m_AutoCloseBraces = New_AutoCloseBraces
    PropertyChanged "AutoCloseBraces"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoCloseQuotes() As Boolean    'When set to true quotes will automatically be closed.
    AutoCloseQuotes = m_AutoCloseQuotes
End Property

Public Property Let AutoCloseQuotes(ByVal New_AutoCloseQuotes As Boolean)
    m_AutoCloseQuotes = New_AutoCloseQuotes
    PropertyChanged "AutoCloseQuotes"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get HighlightBraces() As Boolean    'When set to true any braces the cursor is next to will be highlighted.
    HighlightBraces = m_BraceHighlight
End Property

Public Property Let HighlightBraces(ByVal New_BraceHighlight As Boolean)
    m_BraceHighlight = New_BraceHighlight
    PropertyChanged "BraceHighlight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaretForeColor() As OLE_COLOR   'Set's the color of the caret.
    CaretForeColor = m_CaretForeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Text() As String    'Allows you to get and set the text of the scintilla window.
    Text = DirectSCI.GetText
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    DirectSCI.SetText New_Text
    DirectSCI.SetFocus
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get SelText() As String 'Allows you to get and set the seltext of the scintilla window.
    SelText = DirectSCI.GetSelText 'm_SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    m_SelText = New_SelText
    PropertyChanged "SelText"
    DirectSCI.SetSelText m_SelText
    DirectSCI.SetFocus
End Property


Public Property Let CaretForeColor(ByVal New_CaretForeColor As OLE_COLOR)
    m_CaretForeColor = New_CaretForeColor
    PropertyChanged "CaretForeColor"
    DirectSCI.SetCaretFore New_CaretForeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblue
Public Property Get LineBackColor() As OLE_COLOR    'Allows you to control the backcolor of the active line.
    LineBackColor = m_LineBackColor
End Property

Public Property Let LineBackColor(ByVal New_LineBackColor As OLE_COLOR)
    m_LineBackColor = New_LineBackColor
    PropertyChanged "LineBackColor"
    DirectSCI.SetCaretLineBack New_LineBackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get LineVisible() As Boolean    'When set to true the active line will be highlighted using the color selected from LineBackColor.
    LineVisible = m_LineVisible
End Property

Public Property Let LineVisible(ByVal New_LineVisible As Boolean)
    m_LineVisible = New_LineVisible
    PropertyChanged "LineVisible"
    DirectSCI.SetCaretLineVisible m_LineVisible
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get CaretWidth() As Long    'Allow's you to control the width of the caret line.  The maximum value is 3.
    CaretWidth = m_CaretWidth
End Property

Public Property Let CaretWidth(ByVal New_CaretWidth As Long)
    If New_CaretWidth > 3 Then New_CaretWidth = 3
    m_CaretWidth = New_CaretWidth
    PropertyChanged "CaretWidth"
    DirectSCI.SetCaretWidth m_CaretWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ClearUndoAfterSave() As Boolean 'If set to true then the undo buffer will be cleared when calling SaveToFile.
    ClearUndoAfterSave = m_ClearUndoAfterSave
End Property

Public Property Let ClearUndoAfterSave(ByVal New_ClearUndoAfterSave As Boolean)
    m_ClearUndoAfterSave = New_ClearUndoAfterSave
    PropertyChanged "ClearUndoAfterSave"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get BookmarkBack() As OLE_COLOR 'Allows you to configure the backcolor of the bookmark display.
    BookmarkBack = m_BookmarkBack
End Property

Public Property Let BookmarkBack(ByVal New_BookmarkBack As OLE_COLOR)
    m_BookmarkBack = New_BookmarkBack
    PropertyChanged "BookMarkBack"
    DirectSCI.MarkerSetBack 1, m_BookmarkBack
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get BookMarkFore() As OLE_COLOR 'Allows you to configure the forecolor of the bookmark display.
    BookMarkFore = m_BookMarkFore
End Property

Public Property Let BookMarkFore(ByVal New_BookMarkFore As OLE_COLOR)
    m_BookMarkFore = New_BookMarkFore
    PropertyChanged "BookMarkFore"
    DirectSCI.MarkerSetFore 1, m_BookMarkFore
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FoldHi() As OLE_COLOR   'This property is used for the folding gutter's back color.  The Hi color is the primary color, the Lo Color is the secondary color.
    FoldHi = m_FoldHi
End Property

Public Property Let FoldHi(ByVal New_FoldHi As OLE_COLOR)
    m_FoldHi = New_FoldHi
    PropertyChanged "FoldHi"
    If New_FoldHi <> m_def_FoldHi Then
      DirectSCI.SetFoldMarginHiColour True, m_FoldHi
    Else
      DirectSCI.SetFoldMarginHiColour False, m_FoldHi
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FoldLo() As OLE_COLOR   'This property is used for the folding gutter's back color.  The Hi color is the primary color, the Lo Color is the secondary color.
    FoldLo = m_FoldLo
End Property

Public Property Let FoldLo(ByVal New_FoldLo As OLE_COLOR)
    m_FoldLo = New_FoldLo
    PropertyChanged "FoldLo"
    If New_FoldLo <> m_def_FoldLo Then
      DirectSCI.SetFoldMarginColour True, m_FoldLo
    Else
      DirectSCI.SetFoldMarginColour False, m_FoldLo
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get MarkerBack() As OLE_COLOR   'Allows you to configure the backcolor of the folding markers.
    MarkerBack = m_MarkerBack
End Property

Public Property Let MarkerBack(ByVal New_MarkerBack As OLE_COLOR)
    m_MarkerBack = New_MarkerBack
    PropertyChanged "MarkerBack"
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEROPEN, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDER, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDERMIDTAIL, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDERSUB, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDERTAIL, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEROPEN, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEROPENMID, m_MarkerBack
    DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEREND, m_MarkerBack
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlack
Public Property Get MarkerFore() As OLE_COLOR   'Allows you to configure the forecolor of the folding marker.
    MarkerFore = m_MarkerFore
End Property

Public Property Let MarkerFore(ByVal New_MarkerFore As OLE_COLOR)
    m_MarkerFore = New_MarkerFore
    PropertyChanged "MarkerFore"
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEROPEN, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDER, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDERMIDTAIL, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDERSUB, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDERTAIL, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEROPEN, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEROPENMID, m_MarkerFore
    DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEREND, m_MarkerFore
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbhighlight
Public Property Get SelBack() As OLE_COLOR  'This allow's you to set the backcolor for selected text.
    SelBack = m_SelBack
End Property

Public Property Let SelBack(ByVal New_SelBack As OLE_COLOR)
    m_SelBack = New_SelBack
    PropertyChanged "SelBack"
    DirectSCI.SetSelBack True, m_SelBack
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get IndentationGuide() As Boolean   'If true indention guide's will be displayed.
    IndentationGuide = m_IndentationGuide
End Property

Public Property Let IndentationGuide(ByVal New_IndentationGuide As Boolean)
    m_IndentationGuide = New_IndentationGuide
    PropertyChanged "IndentationGuide"
    DirectSCI.SetIndentationGuides m_IndentationGuide
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,vbwhite
Public Property Get SelFore() As OLE_COLOR  'The allows you to control the fore color of the selected color.
    SelFore = m_SelFore
End Property

Public Property Let SelFore(ByVal New_SelFore As OLE_COLOR)
    m_SelFore = New_SelFore
    PropertyChanged "SelFore"
    DirectSCI.SetSelFore True, m_SelFore
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get EndAtLastLine() As Boolean  'If set to true then the document won't scroll past the last line.  If false it will allow you to scroll a bit past the end of the file.
    EndAtLastLine = m_EndAtLastLine
End Property

Public Property Let EndAtLastLine(ByVal New_EndAtLastLine As Boolean)
    m_EndAtLastLine = New_EndAtLastLine
    PropertyChanged "EndAtLastLine"
    DirectSCI.SetEndAtLastLine m_EndAtLastLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get OverType() As Boolean   'If true then entered text will overtype any text beyond it.
    OverType = m_OverType
End Property

Public Property Let OverType(ByVal New_OverType As Boolean)
    m_OverType = New_OverType
    PropertyChanged "OverType"
    DirectSCI.SetOvertype m_OverType
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ScrollBarH() As Boolean  'If true then the horizontal scrollbar will be visible.  If false it will be hidden.
    ScrollBarH = m_ScrollBarH
End Property

Public Property Let ScrollBarH(ByVal New_ScrollBarH As Boolean)
    m_ScrollBarH = New_ScrollBarH
    PropertyChanged "ScrollBarH"
    DirectSCI.SetHScrollBar m_ScrollBarH
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ScrollBarV() As Boolean 'If true then the vertical scrollbar will be visible.  If alse it will be hidden.
    ScrollBarV = m_ScrollBarV
End Property

Public Property Let ScrollBarV(ByVal New_ScrollBarV As Boolean)
    m_ScrollBarV = New_ScrollBarV
    PropertyChanged "ScrollBarV"
    DirectSCI.SetVScrollBar New_ScrollBarV
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ViewEOL() As Boolean    'If this is set to true EOL markers will be displayed.
    ViewEOL = m_ViewEOL
End Property

Public Property Let ViewEOL(ByVal New_ViewEOL As Boolean)
    m_ViewEOL = New_ViewEOL
    PropertyChanged "ViewEOL"
    DirectSCI.SetViewEOL New_ViewEOL
End Property

Public Property Get ShowCallTips() As Boolean   'If this is set to true then calltips will be displayed.  To use this you must also use <B>LoadAPIFile</b> to load an external API file which contains simple instructions to the editor on what calltips to display.
    ShowCallTips = m_ShowCallTips
End Property

Public Property Let ShowCallTips(ByVal New_ShowCallTips As Boolean)
    m_ShowCallTips = New_ShowCallTips
    PropertyChanged "ShowCallTips"
    bShowCallTips = m_ShowCallTips
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ViewWhiteSpace() As Boolean 'When this is set to true whitespace markers will be visible.
    ViewWhiteSpace = m_ViewWhiteSpace
End Property

Public Property Let ViewWhiteSpace(ByVal New_ViewWhiteSpace As Boolean)
    m_ViewWhiteSpace = New_ViewWhiteSpace
    PropertyChanged "ViewWhiteSpace"
    DirectSCI.SetViewWS CLng(m_ViewWhiteSpace)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get EdgeColor() As OLE_COLOR 'This allows you to control the color of the Edge line.
    EdgeColor = m_EdgeColor
End Property

Public Property Let EdgeColor(ByVal New_EdgeColor As OLE_COLOR)
    m_EdgeColor = New_EdgeColor
    PropertyChanged "EdgeColor"
    DirectSCI.SetEdgeColour m_EdgeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get EdgeColumn() As Long    'This allows you to control which column the edge line is located at.
    EdgeColumn = m_EdgeColumn
End Property

Public Property Let EdgeColumn(ByVal New_EdgeColumn As Long)
    m_EdgeColumn = New_EdgeColumn
    PropertyChanged "EdgeColumn"
    DirectSCI.SetEdgeColumn m_EdgeColumn
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get EdgeMode() As edge  'This allow's you to control which edge mode to utilize.
    EdgeMode = m_EdgeMode
End Property

Public Property Let EdgeMode(ByVal New_EdgeMode As edge)
    m_EdgeMode = New_EdgeMode
    PropertyChanged "EdgeMode"
    DirectSCI.SetEdgeMode m_EdgeMode
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get EOL() As EOLStyle   'This allows you to control which EOL style to utilize.  Scintilla supports CR+LF, CR, and LF.
    EOL = m_EOL
End Property

Public Property Let EOL(ByVal New_EOL As EOLStyle)
    m_EOL = New_EOL
    PropertyChanged "EOL"
    DirectSCI.SetEOLMode m_EOL
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Folding() As Boolean    'If true folding will be automatically handled.
    Folding = m_Folding
End Property

Public Property Let Folding(ByVal New_Folding As Boolean)
    m_Folding = New_Folding
    PropertyChanged "Folding"
    If m_Folding Then
      DirectSCI.SetMarginWidthN 2, Gutter2Width
    Else
      DirectSCI.SetMarginWidthN 2, 0
    End If
    InitFolding New_Folding
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Gutter0Type() As GutterType
    Gutter0Type = m_Gutter0Type
End Property

Public Property Let Gutter0Type(ByVal New_Gutter0Type As GutterType)
    m_Gutter0Type = New_Gutter0Type
    PropertyChanged "Gutter0Type"
    DirectSCI.SetMarginTypeN 0, m_Gutter0Type
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Gutter0Width() As Long
    Gutter0Width = m_Gutter0Width
End Property

Public Property Let Gutter0Width(ByVal New_Gutter0Width As Long)
    m_Gutter0Width = New_Gutter0Width
    PropertyChanged "Gutter0Width"
    DirectSCI.SetMarginWidthN 0, New_Gutter0Width
    If LineNumbers = True Then
      DirectSCI.SetMarginWidthN 0, m_Gutter0Width
    Else
      DirectSCI.SetMarginWidthN 0, 0
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Gutter1Type() As GutterType
    Gutter1Type = m_Gutter1Type
End Property

Public Property Let Gutter1Type(ByVal New_Gutter1Type As GutterType)
    m_Gutter1Type = New_Gutter1Type
    PropertyChanged "Gutter1Type"
    DirectSCI.SetMarginTypeN 1, m_Gutter1Type
    If ShowFlags = True Then
      DirectSCI.SetMarginWidthN 1, New_Gutter1Type
    Else
      DirectSCI.SetMarginWidthN 1, 0
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Gutter1Width() As Long
    Gutter1Width = m_Gutter1Width
End Property

Public Property Let Gutter1Width(ByVal New_Gutter1Width As Long)
    m_Gutter1Width = New_Gutter1Width
    PropertyChanged "Gutter1Width"
    If Folding = True Then
      DirectSCI.SetMarginWidthN 2, m_Gutter1Width
    Else
      DirectSCI.SetMarginWidthN 2, 0
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Gutter2Type() As GutterType
    Gutter2Type = m_Gutter2Type
End Property

Public Property Let Gutter2Type(ByVal New_Gutter2Type As GutterType)
    m_Gutter2Type = New_Gutter2Type
    PropertyChanged "Gutter2Type"
    DirectSCI.SetMarginTypeN 2, m_Gutter2Type
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Gutter2Width() As Long
    Gutter2Width = m_Gutter2Width
End Property

Public Property Let Gutter2Width(ByVal New_Gutter2Width As Long)
    m_Gutter2Width = New_Gutter2Width
    PropertyChanged "Gutter2Width"
    DirectSCI.SetMarginWidthN 2, New_Gutter2Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get MaintainIndentation() As Boolean 'If this is set to true the editor will automatically keep the previous line's indentation.
    MaintainIndentation = m_MaintainIndentation
End Property

Public Property Let MaintainIndentation(ByVal New_MaintainIndentation As Boolean)
    m_MaintainIndentation = New_MaintainIndentation
    PropertyChanged "MaintainIndentation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get TabIndents() As Boolean 'If this is true tab inserts indent characters.  If it is set to false tab will insert spaces.
    TabIndents = m_TabIndents
End Property

Public Property Let TabIndents(ByVal New_TabIndents As Boolean)
    m_TabIndents = New_TabIndents
    PropertyChanged "TabIndents"
    DirectSCI.SetTabIndents m_TabIndents
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BackSpaceUnIndents() As Boolean 'If tabindents is set to false, and BackSpaceUnIndents is set to true then the backspaceunindents will remove the same number of spaces as tab inserts.  If it's set to false then it will work normally.
    BackSpaceUnIndents = m_BackSpaceUnIndents
End Property

Public Property Let BackSpaceUnIndents(ByVal New_BackSpaceUnIndents As Boolean)
    m_BackSpaceUnIndents = New_BackSpaceUnIndents
    PropertyChanged "BackSpaceUnIndents"
    DirectSCI.SetBackSpaceUnIndents m_BackSpaceUnIndents
End Property

Public Property Get AutoCompleteOnCTRLSpace() As Boolean    'If this is set to true then an autocomplete list will be displayed when a user hits Ctrl+Space.
  AutoCompleteOnCTRLSpace = m_AutoCompleteOnCTRLSpace
End Property

Public Property Let AutoCompleteOnCTRLSpace(ByVal New_AutoCompleteOnCTRLSpace As Boolean)
  m_AutoCompleteOnCTRLSpace = New_AutoCompleteOnCTRLSpace
  PropertyChanged "AutoCompleteOnCTRLSpace"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,.
Public Property Get AutoCompleteStart() As String   'This property allows you to assign a specific single character to display autocomplete.  By default the character is <B>"."</b>.
    AutoCompleteStart = m_AutoCompleteStart
End Property

Public Property Let AutoCompleteStart(ByVal New_AutoCompleteStart As String)
    If Len(New_AutoCompleteStart) > 1 Then
      MsgBox "AutoCompleteStart property can only be set to a single character.", vbOKOnly, "Error"
      New_AutoCompleteStart = m_def_AutoCompleteStart
    End If
    m_AutoCompleteStart = New_AutoCompleteStart
    PropertyChanged "AutoCompleteStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get AutoCompleteString() As String  'This store's the list which autocomplete will use.  Each word needs to be seperated by a space.
    AutoCompleteString = m_AutoCompleteString
End Property

Public Property Let AutoCompleteString(ByVal New_AutoCompleteString As String)
    m_AutoCompleteString = New_AutoCompleteString
    PropertyChanged "AutoCompleteString"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoShowAutoComplete() As Boolean   'If set to true then an autocomplete box will be displayed if a user enters the single character in the <B>AutoCompleteStart</b> property.
    AutoShowAutoComplete = m_AutoShowAutoComplete
End Property

Public Property Let AutoShowAutoComplete(ByVal New_AutoShowAutoComplete As Boolean)
    m_AutoShowAutoComplete = New_AutoShowAutoComplete
    PropertyChanged "AutoShowAutoComplete"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ContextMenu() As Boolean    'If set to true then the default Scintilla context menu will be displayed when a user right clicks on the window.  If this is set to false then no context menu will be displayed.  If you are utilizing a customer context menu then this should be set to false.
    ContextMenu = m_ContextMenu
End Property

Public Property Let ContextMenu(ByVal New_ContextMenu As Boolean)
    m_ContextMenu = New_ContextMenu
    PropertyChanged "ContextMenu"
    DirectSCI.UsePopUp m_ContextMenu
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!    'If this is set to true then the AutoComplete list will ignore the case.  If it is set to false then proper character case will be required.
'MemberInfo=0,0,0,1
Public Property Get IgnoreAutoCompleteCase() As Boolean
    IgnoreAutoCompleteCase = m_IgnoreAutoCompleteCase
End Property

Public Property Let IgnoreAutoCompleteCase(ByVal New_IgnoreAutoCompleteCase As Boolean)
    m_IgnoreAutoCompleteCase = New_IgnoreAutoCompleteCase
    PropertyChanged "IgnoreAutoCompleteCase"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get LineNumbers() As Boolean    'If this is set to true then the first gutter will be visible and display line numbers.  If this is false then the first gutter will remain hidden.
    LineNumbers = m_LineNumbers
End Property

Public Property Let LineNumbers(ByVal New_LineNumbers As Boolean)
    m_LineNumbers = New_LineNumbers
    PropertyChanged "LineNumbers"
    If m_LineNumbers Then
      DirectSCI.SetMarginWidthN 0, Gutter0Width
    Else
      DirectSCI.SetMarginWidthN 0, 0
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ReadOnly() As Boolean  'This property allows you to set the readonly status of Scintilla.  When in readonly you can scroll the document, but no editing can be done.
    ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(ByVal New_ReadOnly As Boolean)
    m_ReadOnly = New_ReadOnly
    PropertyChanged "ReadOnly"
    DirectSCI.SetReadOnly m_ReadOnly
End Property

Public Property Get Modified() As Boolean   'This is a read only property.  It allows you to get the modified status of the Scintilla window.
    Modified = DirectSCI.GetModify
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,2000
Public Property Get ScrollWidth() As Long   'Scintilla's design does not automatically size the horizontal scrollbar to the size of the longest line.  It gives it a set size.  By default it allows 2000 characters per line.  This allows you to control how far the Horizontal scrollbar can be scrolled.
    ScrollWidth = m_ScrollWidth
End Property

Public Property Let ScrollWidth(ByVal New_ScrollWidth As Long)
    m_ScrollWidth = New_ScrollWidth
    PropertyChanged "ScrollWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ShowFlags() As Boolean  'If this is true the second gutter will be displayed and Flags/Bookmarks will be displayed.
    ShowFlags = m_ShowFlags
End Property

Public Property Let ShowFlags(ByVal New_ShowFlags As Boolean)
    m_ShowFlags = New_ShowFlags
    PropertyChanged "ShowFlags"
    If m_ShowFlags Then
      DirectSCI.SetMarginWidthN 1, Gutter1Width
    Else
      DirectSCI.SetMarginWidthN 1, 0
    End If
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,4
Public Property Get IndentWidth() As Long   'This controls the number of spaces Tab will indent.  IndentWidth only applies if <B>TabIndents</b> is set to false.
    IndentWidth = m_IndentWidth
End Property

Public Property Let IndentWidth(ByVal New_IndentWidth As Long)
    m_IndentWidth = New_IndentWidth
    PropertyChanged "IndentWidth"
    DirectSCI.SetTabWidth IndentWidth
    'SetIndent m_IndentWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get useTabs() As Boolean
    useTabs = m_UseTabs
End Property

Public Property Let useTabs(ByVal New_UseTabs As Boolean)
    m_UseTabs = New_UseTabs
    PropertyChanged "UseTabs"
    DirectSCI.SetUseTabs m_UseTabs
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get WordWrap() As WrapStyle 'If set to true the document will wrap lines which are longer than itself.  If false then it will dsiplay normally.
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As WrapStyle)
    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"
    DirectSCI.SetWrapMode New_WordWrap
End Property

Public Property Get CurHigh() As String
  CurHigh = m_curHigh
End Property

Public Property Let CurHigh(New_CurHigh As String)
  m_curHigh = New_CurHigh
End Property

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    Dim scMsg As SCNotification
    Dim iMsg As Long
    iMsg = uMsg
    Dim tHdr As NMHDR
    Dim strTmp As String
    Dim Shift As Long
    Dim tmpStr As String
    Dim lP As POINTAPI
    Dim zPos As Long
    Dim chl As String, strMatch As String
    Dim lPos As Long
    Select Case iMsg
      Case WM_NOTIFY
        CopyMemory scMsg, ByVal lParam, Len(scMsg)
        tHdr = scMsg.NotifyHeader
        If (tHdr.hwndFrom = SCI) Then
          'Scintilla has given some information.  Let's see what it is
          'and route it to the proper place.
          ' Any commented with TODO have not been implimented yet.
          Select Case tHdr.Code
            Case SCN_MODIFIED
              RaiseEvent OnModified(scMsg.Position, scMsg.modificationType)
            Case 2012
              RaiseEvent PosChanged(scMsg.Position)
            Case SCN_KEY
              RaiseEvent Key(scMsg.ch, scMsg.modifiers)
            Case SCN_STYLENEEDED
              RaiseEvent StyleNeeded(scMsg.Position)
            Case SCN_CHARADDED
              RaiseEvent CharAdded(scMsg.ch)
              chStore = scMsg.ch
              If AutoCloseBraces Then
                chl = Chr(scMsg.ch)
                If chl = "(" Or chl = "[" Or chl = "{" Then
                  strMatch = MatchBrace(chl)
                  lPos = DirectSCI.GetCurPos
                  DirectSCI.AddText 1, strMatch
                  DirectSCI.SetSel lPos, lPos
                End If
              End If
              If AutoCloseQuotes Then
                chl = Chr(scMsg.ch)
                If chl = """" Or chl = "'" Then
                  If chl = """" Then
                    strMatch = """"
                  Else
                    strMatch = "'"
                  End If
                  lPos = DirectSCI.GetCurPos
                  DirectSCI.AddText 1, strMatch
                  DirectSCI.SetSel lPos, lPos
                End If
              End If
              'chl = scMsg.ch
              If MaintainIndentation = True Then
                If scMsg.ch = 13 Or scMsg.ch = 10 Then
                  MaintainIndent
                End If
              End If
              If AutoShowAutoComplete Then
                StartAutoComplete scMsg.ch
              End If
              If bShowCallTips Then
                StartCallTip scMsg.ch
              End If

            Case SCN_SAVEPOINTREACHED
              RaiseEvent SavePointReached
            Case SCN_SAVEPOINTLEFT
              RaiseEvent SavePointLeft
            Case SCN_MODIFYATTEMPTRO
              'TODO
            Case SCN_DOUBLECLICK
              RaiseEvent DoubleClick
            Case SCN_UPDATEUI
              If HighlightBraces = False Then
                DirectSCI.BraceBadLight -1
                DirectSCI.BraceHighlight -1, -1
              End If
              If HighlightBraces = True Then
                  Dim pos As Long, pos2 As Long
                  pos2 = INVALID_POSITION
                  If IsBrace(DirectSCI.CharAtPos(DirectSCI.GetCurPos)) Then
                      pos2 = DirectSCI.GetCurPos
                  ElseIf IsBrace(DirectSCI.CharAtPos(DirectSCI.GetCurPos - 1)) Then
                      pos2 = DirectSCI.GetCurPos - 1
                  End If
                  If pos2 <> INVALID_POSITION Then
                      pos = SendMessage(SCI, SCI_BRACEMATCH, pos2, CLng(0))
                      If pos = INVALID_POSITION Then
                          Call SendEditor(SCI_BRACEBADLIGHT, pos2)
                      Else
                          Call SendEditor(SCI_BRACEHIGHLIGHT, pos, pos2)
                          'If m_IndGuides Then
                              Call SendEditor(SCI_SETHIGHLIGHTGUIDE, DirectSCI.GetColumn)
                          'End If
                      End If
                  Else
                      Call SendEditor(SCI_BRACEHIGHLIGHT, INVALID_POSITION, INVALID_POSITION)
                  End If
              End If
              RaiseEvent UpdateUI
            Case SCN_MACRORECORD
              HandleMacroCall scMsg.message, Chr(chStore)
              RaiseEvent MacroRecord(scMsg.message, wParam)
            Case SCN_MARGINCLICK
              Dim lLine As Long, lMargin As Long, lPosition As Long
              lPosition = scMsg.Position
              lLine = SendEditor(SCI_LINEFROMPOSITION, lPosition)
              lMargin = scMsg.margin
              If lMargin = MARGIN_SCRIPT_FOLD_INDEX Then
                
                Call SendEditor(SCI_TOGGLEFOLD, lLine, 0)
              End If
              RaiseEvent MarginClick(scMsg.modifiers, scMsg.Position)
            Case SCN_NEEDSHOWN
              'TODO
            Case SCN_PAINTED
              RaiseEvent Painted
            Case SCN_AUTOCSELECTION
              strTmp = String(255, " ")
              ConvCStringToVBString strTmp, scMsg.Text
              zPos = InStr(strTmp, vbNullChar)
              strTmp = Left(strTmp, zPos - 1)
              RaiseEvent AutoCSelection(strTmp)
            Case SCN_USERLISTSELECTION
              strTmp = String(255, " ")
              ConvCStringToVBString strTmp, scMsg.Text
              zPos = InStr(strTmp, vbNullChar)
              strTmp = Left(strTmp, zPos - 1)
              RaiseEvent UserListSelection(scMsg.listType, strTmp)
            Case SCN_DWELLSTART
              'TODO
            Case SCN_DWELLEND
              'TODO
              
          End Select
        End If
      Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(1, 0, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
      Case WM_CLOSE
        ' Just to be safe detach it.
        iSCISet = 0
        'Detach
      Case WM_CHAR
        If AutoCompleteOnCTRLSpace Then
          If wParam = 32 And piGetShiftState = 4 Then
            bHandled = True
            lReturn = 0
            RaiseEvent KeyPress(wParam)
            ShowAutoComplete AutoCompleteString
          Else
            bHandled = False
            lReturn = 0
            RaiseEvent KeyPress(wParam)
          End If
        Else
          RaiseEvent KeyPress(wParam)
        End If
      Case WM_RBUTTONDOWN
        lP = GetWindowCursorPos(SCI)
        RaiseEvent MouseDown(2, 0, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
      Case WM_LBUTTONUP
        lP = GetWindowCursorPos(SCI)
        RaiseEvent MouseUp(1, 0, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
      Case WM_RBUTTONUP
        lP = GetWindowCursorPos(SCI)
        RaiseEvent MouseUp(2, 0, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
      Case WM_KEYDOWN
        If wParam = 32 Then
          If piGetShiftState = 5 Then
            StartCallTip Asc("(")
          End If
        End If
        If bShowCallTips Then
          StartCallTip scMsg.ch
        End If
        
        RaiseEvent KeyDown(wParam, piGetShiftState)
      Case WM_KEYUP
        If bShowCallTips Then
          StartCallTip scMsg.ch
        End If
        RaiseEvent KeyUp(wParam, piGetShiftState)
      Case WM_SETFOCUS
        DirectSCI.SetFocus
    End Select


End Sub

Private Sub UserControl_GotFocus()
  DirectSCI.SetFocus
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    iccex.lngSize = LenB(iccex)
    iccex.lngICC = ICC_USEREX_CLASSES
    InitCommonControlsEx iccex
    'this is to prevent crash
    m_hMod = LoadLibrary("shell32.dll")

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_AutoCloseBraces = m_def_AutoCloseBraces
    
    m_AutoCloseQuotes = m_def_AutoCloseQuotes
    m_BraceHighlight = m_def_BraceHighlight
    m_CaretForeColor = m_def_CaretForeColor
    m_LineBackColor = m_def_LineBackColor
    m_LineVisible = m_def_LineVisible
    m_CaretWidth = m_def_CaretWidth
    m_ClearUndoAfterSave = m_def_ClearUndoAfterSave
    m_BookmarkBack = m_def_BookmarkBack
    m_BookMarkFore = m_def_BookMarkFore
    m_FoldHi = m_def_FoldHi
    m_FoldLo = m_def_FoldLo
    m_MarkerBack = m_def_MarkerBack
    m_MarkerFore = m_def_MarkerFore
    m_SelBack = m_def_SelBack
    m_SelFore = m_def_SelFore
    m_EndAtLastLine = m_def_EndAtLastLine
    m_OverType = m_def_OverType
    m_ScrollBarH = m_def_ScrollBarH
    m_ScrollBarV = m_def_ScrollBarV
    m_ViewEOL = m_def_ViewEOL
    m_ViewWhiteSpace = m_def_ViewWhiteSpace
    m_ShowCallTips = m_def_ShowCallTips
    bShowCallTips = m_def_ShowCallTips
    m_EdgeColor = m_def_EdgeColor
    m_EdgeColumn = m_def_EdgeColumn
    m_EdgeMode = m_def_EdgeMode
    m_EOL = m_def_EOL
    m_Folding = m_def_Folding
    m_Gutter0Type = m_def_Gutter0Type
    m_Gutter0Width = m_def_Gutter0Width
    m_Gutter1Type = m_def_Gutter1Type
    m_Gutter1Width = m_def_Gutter1Width
    m_Gutter2Type = m_def_Gutter2Type
    m_Gutter2Width = m_def_Gutter2Width
    m_MaintainIndentation = m_def_MaintainIndentation
    m_TabIndents = m_def_TabIndents
    m_BackSpaceUnIndents = m_def_BackSpaceUnIndents
    m_IndentWidth = m_def_IndentWidth
    m_UseTabs = m_def_UseTabs
    m_WordWrap = m_def_WordWrap
    m_FoldMarker = m_def_FoldMarker
    m_MarginFore = m_def_MarginFore
    m_MarginBack = m_def_MarginBack
    m_Text = m_def_Text
    m_SelText = m_def_SelText
    m_AutoCompleteStart = m_def_AutoCompleteStart
    m_AutoCompleteOnCTRLSpace = m_def_AutoCompleteOnCTRLSpace
    m_AutoCompleteString = m_def_AutoCompleteString
    m_AutoShowAutoComplete = m_def_AutoShowAutoComplete
    m_ContextMenu = m_def_ContextMenu
    m_IgnoreAutoCompleteCase = m_def_IgnoreAutoCompleteCase
    m_LineNumbers = m_def_LineNumbers
    m_ReadOnly = m_def_ReadOnly
    m_ScrollWidth = m_def_ScrollWidth
    m_ShowFlags = m_def_ShowFlags
    m_FoldAtElse = m_def_FoldAtElse
    m_FoldComment = m_def_FoldComment
    m_FoldCompact = m_def_FoldCompact
    m_FoldHTML = m_def_FoldHTML
    m_IndentationGuide = m_def_IndentationGuide
    m_SelStart = m_def_SelStart
    m_SelEnd = m_def_SelEnd
    m_BraceMatch = m_def_BraceMatch
    m_BraceBad = m_def_BraceBad
    m_BraceMatchBold = m_def_BraceMatchBold
    m_BraceMatchItalic = m_def_BraceMatchItalic
    m_BraceMatchUnderline = m_def_BraceMatchUnderline
    m_BraceMatchBack = m_def_BraceMatchBack
    m_BraceBadBack = m_def_BraceBadBack
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoCloseBraces = PropBag.ReadProperty("AutoCloseBraces", m_def_AutoCloseBraces)
    m_AutoCloseQuotes = PropBag.ReadProperty("AutoCloseQuotes", m_def_AutoCloseQuotes)
    m_BraceHighlight = PropBag.ReadProperty("BraceHighlight", m_def_BraceHighlight)
    m_CaretForeColor = PropBag.ReadProperty("CaretForeColor", m_def_CaretForeColor)
    m_LineBackColor = PropBag.ReadProperty("LineBackColor", m_def_LineBackColor)
    m_LineVisible = PropBag.ReadProperty("LineVisible", m_def_LineVisible)
    m_CaretWidth = PropBag.ReadProperty("CaretWidth", m_def_CaretWidth)
    m_ClearUndoAfterSave = PropBag.ReadProperty("ClearUndoAfterSave", m_def_ClearUndoAfterSave)
    m_BookmarkBack = PropBag.ReadProperty("BookMarkBack", m_def_BookmarkBack)
    m_BookMarkFore = PropBag.ReadProperty("BookMarkFore", m_def_BookMarkFore)
    m_FoldHi = PropBag.ReadProperty("FoldHi", m_def_FoldHi)
    m_FoldLo = PropBag.ReadProperty("FoldLo", m_def_FoldLo)
    m_MarkerBack = PropBag.ReadProperty("MarkerBack", m_def_MarkerBack)
    m_MarkerFore = PropBag.ReadProperty("MarkerFore", m_def_MarkerFore)
    m_SelBack = PropBag.ReadProperty("SelBack", m_def_SelBack)
    m_SelFore = PropBag.ReadProperty("SelFore", m_def_SelFore)
    m_EndAtLastLine = PropBag.ReadProperty("EndAtLastLine", m_def_EndAtLastLine)
    m_OverType = PropBag.ReadProperty("OverType", m_def_OverType)
    m_ScrollBarH = PropBag.ReadProperty("ScrollBarH", m_def_ScrollBarH)
    m_ScrollBarV = PropBag.ReadProperty("ScrollBarV", m_def_ScrollBarV)
    m_ViewEOL = PropBag.ReadProperty("ViewEOL", m_def_ViewEOL)
    m_ViewWhiteSpace = PropBag.ReadProperty("ViewWhiteSpace", m_def_ViewWhiteSpace)
    m_ShowCallTips = PropBag.ReadProperty("ShowCallTips", m_def_ShowCallTips)
    bShowCallTips = m_ShowCallTips
    m_EdgeColor = PropBag.ReadProperty("EdgeColor", m_def_EdgeColor)
    m_EdgeColumn = PropBag.ReadProperty("EdgeColumn", m_def_EdgeColumn)
    m_EdgeMode = PropBag.ReadProperty("EdgeMode", m_def_EdgeMode)
    m_EOL = PropBag.ReadProperty("EOL", m_def_EOL)
    m_Folding = PropBag.ReadProperty("Folding", m_def_Folding)
    m_Gutter0Type = PropBag.ReadProperty("Gutter0Type", m_def_Gutter0Type)
    m_Gutter0Width = PropBag.ReadProperty("Gutter0Width", m_def_Gutter0Width)
    m_Gutter1Type = PropBag.ReadProperty("Gutter1Type", m_def_Gutter1Type)
    m_Gutter1Width = PropBag.ReadProperty("Gutter1Width", m_def_Gutter1Width)
    m_Gutter2Type = PropBag.ReadProperty("Gutter2Type", m_def_Gutter2Type)
    m_Gutter2Width = PropBag.ReadProperty("Gutter2Width", m_def_Gutter2Width)
    m_MaintainIndentation = PropBag.ReadProperty("MaintainIndentation", m_def_MaintainIndentation)
    m_TabIndents = PropBag.ReadProperty("TabIndents", m_def_TabIndents)
    m_BackSpaceUnIndents = PropBag.ReadProperty("BackSpaceUnIndents", m_def_BackSpaceUnIndents)
    m_IndentWidth = PropBag.ReadProperty("IndentWidth", m_def_IndentWidth)
    m_UseTabs = PropBag.ReadProperty("UseTabs", m_def_UseTabs)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
    m_FoldMarker = PropBag.ReadProperty("FoldMarker", m_def_FoldMarker)
    m_MarginFore = PropBag.ReadProperty("MarginFore", m_def_MarginFore)
    m_MarginBack = PropBag.ReadProperty("MarginBack", m_def_MarginBack)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_SelText = PropBag.ReadProperty("SelText", m_def_SelText)
    m_AutoCompleteStart = PropBag.ReadProperty("AutoCompleteStart", m_def_AutoCompleteStart)
    m_AutoCompleteOnCTRLSpace = PropBag.ReadProperty("AutoCompleteOnCTRLSpace", m_def_AutoCompleteOnCTRLSpace)
    m_AutoCompleteString = PropBag.ReadProperty("AutoCompleteString", m_def_AutoCompleteString)
    m_AutoShowAutoComplete = PropBag.ReadProperty("AutoShowAutoComplete", m_def_AutoShowAutoComplete)
    m_ContextMenu = PropBag.ReadProperty("ContextMenu", m_def_ContextMenu)
    m_IgnoreAutoCompleteCase = PropBag.ReadProperty("IgnoreAutoCompleteCase", m_def_IgnoreAutoCompleteCase)
    m_LineNumbers = PropBag.ReadProperty("LineNumbers", m_def_LineNumbers)
    m_ReadOnly = PropBag.ReadProperty("ReadOnly", m_def_ReadOnly)
    m_ScrollWidth = PropBag.ReadProperty("ScrollWidth", m_def_ScrollWidth)
    m_ShowFlags = PropBag.ReadProperty("ShowFlags", m_def_ShowFlags)
    m_FoldAtElse = PropBag.ReadProperty("FoldAtElse", m_def_FoldAtElse)
    m_FoldComment = PropBag.ReadProperty("FoldComment", m_def_FoldComment)
    m_FoldCompact = PropBag.ReadProperty("FoldCompact", m_def_FoldCompact)
    m_FoldHTML = PropBag.ReadProperty("FoldHTML", m_def_FoldHTML)
    m_IndentationGuide = PropBag.ReadProperty("IndentationGuide", m_def_IndentationGuide)
    m_SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
    m_SelEnd = PropBag.ReadProperty("SelEnd", m_def_SelEnd)
    m_BraceMatch = PropBag.ReadProperty("BraceMatch", m_def_BraceMatch)
    m_BraceBad = PropBag.ReadProperty("BraceBad", m_def_BraceBad)
    m_BraceMatchBold = PropBag.ReadProperty("BraceMatchBold", m_def_BraceMatchBold)
    m_BraceMatchItalic = PropBag.ReadProperty("BraceMatchItalic", m_def_BraceMatchItalic)
    m_BraceMatchUnderline = PropBag.ReadProperty("BraceMatchUnderline", m_def_BraceMatchUnderline)
    m_BraceMatchBack = PropBag.ReadProperty("BraceMatchBack", m_def_BraceMatchBack)
    m_BraceBadBack = PropBag.ReadProperty("BraceBadBack", m_def_BraceBadBack)
End Sub

Public Function InitScintilla(hWndA As Long) As Boolean
    'On Error GoTo errHandler
    InitScintilla = True
    lSCI = LoadLibrary("SciLexer.DLL")   'Load SciLexer.dll from windows directory
    'sci = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "SciMain", WS_CHILD Or WS_MAXIMIZE Or WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or WS_TABSTOP Or WS_CLIPCHILDREN, 0, 0, 0, 0, hWnd, 0, App.hInstance, 0)
    Set DirectSCI = New cDirectSCI  ' Setup the directsci class
    SCI = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "Scint.ocx", WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, hWndA, 0, App.hInstance, 0)
    DirectSCI.SCI = SCI
    If SCI = 0 Then
      RaiseEvent OnError("SCIVB #001", "Failed to initialize the Scintilla interface." & vbCrLf & vbCrLf & _
                       "Please verify that SciLexer.dll is in the program" & vbCrLf & "directory or the windows system32 directory")
      InitScintilla = False
      Exit Function
    End If
    
             
    fWindowProc = GetWindowLong(SCI, GWL_WNDPROC)
    Attach hWndA
    DirectSCI.SetBackSpaceUnIndents BackSpaceUnIndents
    SetOptions
    RemoveHotKeys
    DirectSCI.SetPasteConvertEndings True
    hWndParent = hWndA
    iSCISet = 1
    DirectSCI.SetFocus
    Exit Function
errHandler:
    RaiseEvent OnError(Err.Number, Err.Description)
End Function

Private Sub UserControl_Resize()
  Width = 32 * Screen.TwipsPerPixelX
  Height = 32 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Terminate()
  On Error GoTo Catch
  'Stop all subclassing
  Detach
  FreeLibrary m_hMod
  FreeLibrary lSCI
Catch:
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoCloseBraces", m_AutoCloseBraces, m_def_AutoCloseBraces)
    Call PropBag.WriteProperty("AutoCloseQuotes", m_AutoCloseQuotes, m_def_AutoCloseQuotes)
    Call PropBag.WriteProperty("BraceHighlight", m_BraceHighlight, m_def_BraceHighlight)
    Call PropBag.WriteProperty("CaretForeColor", m_CaretForeColor, m_def_CaretForeColor)
    Call PropBag.WriteProperty("LineBackColor", m_LineBackColor, m_def_LineBackColor)
    Call PropBag.WriteProperty("LineVisible", m_LineVisible, m_def_LineVisible)
    Call PropBag.WriteProperty("CaretWidth", m_CaretWidth, m_def_CaretWidth)
    Call PropBag.WriteProperty("ClearUndoAfterSave", m_ClearUndoAfterSave, m_def_ClearUndoAfterSave)
    Call PropBag.WriteProperty("BookMarkBack", m_BookmarkBack, m_def_BookmarkBack)
    Call PropBag.WriteProperty("BookMarkFore", m_BookMarkFore, m_def_BookMarkFore)
    Call PropBag.WriteProperty("FoldHi", m_FoldHi, m_def_FoldHi)
    Call PropBag.WriteProperty("FoldLo", m_FoldLo, m_def_FoldLo)
    Call PropBag.WriteProperty("MarkerBack", m_MarkerBack, m_def_MarkerBack)
    Call PropBag.WriteProperty("MarkerFore", m_MarkerFore, m_def_MarkerFore)
    Call PropBag.WriteProperty("SelBack", m_SelBack, m_def_SelBack)
    Call PropBag.WriteProperty("SelFore", m_SelFore, m_def_SelFore)
    Call PropBag.WriteProperty("EndAtLastLine", m_EndAtLastLine, m_def_EndAtLastLine)
    Call PropBag.WriteProperty("OverType", m_OverType, m_def_OverType)
    Call PropBag.WriteProperty("ScrollBarH", m_ScrollBarH, m_def_ScrollBarH)
    Call PropBag.WriteProperty("ScrollBarV", m_ScrollBarV, m_def_ScrollBarV)
    Call PropBag.WriteProperty("ViewEOL", m_ViewEOL, m_def_ViewEOL)
    Call PropBag.WriteProperty("ViewWhiteSpace", m_ViewWhiteSpace, m_def_ViewWhiteSpace)
    Call PropBag.WriteProperty("ShowCallTips", m_ShowCallTips, m_def_ShowCallTips)
    Call PropBag.WriteProperty("EdgeColor", m_EdgeColor, m_def_EdgeColor)
    Call PropBag.WriteProperty("EdgeColumn", m_EdgeColumn, m_def_EdgeColumn)
    Call PropBag.WriteProperty("EdgeMode", m_EdgeMode, m_def_EdgeMode)
    Call PropBag.WriteProperty("EOL", m_EOL, m_def_EOL)
    Call PropBag.WriteProperty("Folding", m_Folding, m_def_Folding)
    Call PropBag.WriteProperty("Gutter0Type", m_Gutter0Type, m_def_Gutter0Type)
    Call PropBag.WriteProperty("Gutter0Width", m_Gutter0Width, m_def_Gutter0Width)
    Call PropBag.WriteProperty("Gutter1Type", m_Gutter1Type, m_def_Gutter1Type)
    Call PropBag.WriteProperty("Gutter1Width", m_Gutter1Width, m_def_Gutter1Width)
    Call PropBag.WriteProperty("Gutter2Type", m_Gutter2Type, m_def_Gutter2Type)
    Call PropBag.WriteProperty("Gutter2Width", m_Gutter2Width, m_def_Gutter2Width)
    Call PropBag.WriteProperty("MaintainIndentation", m_MaintainIndentation, m_def_MaintainIndentation)
    Call PropBag.WriteProperty("TabIndents", m_TabIndents, m_def_TabIndents)
    Call PropBag.WriteProperty("BackSpaceUnIndents", m_BackSpaceUnIndents, m_def_BackSpaceUnIndents)
    Call PropBag.WriteProperty("IndentWidth", m_IndentWidth, m_def_IndentWidth)
    Call PropBag.WriteProperty("UseTabs", m_UseTabs, m_def_UseTabs)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
    Call PropBag.WriteProperty("FoldMarker", m_FoldMarker, m_def_FoldMarker)
    Call PropBag.WriteProperty("MarginFore", m_MarginFore, m_def_MarginFore)
    Call PropBag.WriteProperty("MarginBack", m_MarginBack, m_def_MarginBack)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("SelText", m_SelText, m_def_SelText)
    Call PropBag.WriteProperty("AutoCompleteStart", m_AutoCompleteStart, m_def_AutoCompleteStart)
    Call PropBag.WriteProperty("AutoCompleteOnCTRLSpace", m_AutoCompleteOnCTRLSpace, m_def_AutoCompleteOnCTRLSpace)
    Call PropBag.WriteProperty("AutoCompleteString", m_AutoCompleteString, m_def_AutoCompleteString)
    Call PropBag.WriteProperty("AutoShowAutoComplete", m_AutoShowAutoComplete, m_def_AutoShowAutoComplete)
    Call PropBag.WriteProperty("ContextMenu", m_ContextMenu, m_def_ContextMenu)
    Call PropBag.WriteProperty("IgnoreAutoCompleteCase", m_IgnoreAutoCompleteCase, m_def_IgnoreAutoCompleteCase)
    Call PropBag.WriteProperty("LineNumbers", m_LineNumbers, m_def_LineNumbers)
    Call PropBag.WriteProperty("ReadOnly", m_ReadOnly, m_def_ReadOnly)
    Call PropBag.WriteProperty("ScrollWidth", m_ScrollWidth, m_def_ScrollWidth)
    Call PropBag.WriteProperty("ShowFlags", m_ShowFlags, m_def_ShowFlags)
    Call PropBag.WriteProperty("FoldAtElse", m_FoldAtElse, m_def_FoldAtElse)
    
    Call PropBag.WriteProperty("FoldComment", m_FoldComment, m_def_FoldComment)
    Call PropBag.WriteProperty("FoldCompact", m_FoldCompact, m_def_FoldCompact)
    Call PropBag.WriteProperty("FoldHTML", m_FoldHTML, m_def_FoldHTML)
    
    Call PropBag.WriteProperty("IndentationGuide", m_IndentationGuide, m_def_IndentationGuide)
    Call PropBag.WriteProperty("SelStart", m_SelStart, m_def_SelStart)
    Call PropBag.WriteProperty("SelEnd", m_SelEnd, m_def_SelEnd)
    Call PropBag.WriteProperty("BraceMatch", m_BraceMatch, m_def_BraceMatch)
    Call PropBag.WriteProperty("BraceBad", m_BraceBad, m_def_BraceBad)
    Call PropBag.WriteProperty("BraceMatchBold", m_BraceMatchBold, m_def_BraceMatchBold)
    Call PropBag.WriteProperty("BraceMatchItalic", m_BraceMatchItalic, m_def_BraceMatchItalic)
    Call PropBag.WriteProperty("BraceMatchUnderline", m_BraceMatchUnderline, m_def_BraceMatchUnderline)
    Call PropBag.WriteProperty("BraceMatchBack", m_BraceMatchBack, m_def_BraceMatchBack)
    Call PropBag.WriteProperty("BraceBadBack", m_BraceBadBack, m_def_BraceBadBack)
End Sub

Public Sub MoveSCI(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
  SetWindowPos SCI, 0, lLeft, lTop, lWidth / Screen.TwipsPerPixelX, lHeight / Screen.TwipsPerPixelY, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Sub RemoveHotKeys()
  ' This just removes some of the common hot keys that
  ' could cause scintilla to interfere with the application
  DirectSCI.ClearCmdKey Asc("A") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("B") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("C") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("D") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("E") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("F") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("G") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("H") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("I") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("J") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("K") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("L") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("M") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("N") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("O") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("P") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("Q") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("R") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("S") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("T") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("U") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("V") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("W") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("X") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("Y") + LShift(SCMOD_CTRL, 16)
  DirectSCI.ClearCmdKey Asc("Z") + LShift(SCMOD_CTRL, 16)
  'AssignCmdKey 32 + LShift(SCMOD_CTRL, 16), SCI_AUTOCSHOW
End Sub


Public Sub SetOptions()
  DirectSCI.SetCaretFore CaretForeColor
  DirectSCI.SetCaretWidth CaretWidth
  DirectSCI.SetEdgeColour EdgeColor
  DirectSCI.SetEdgeColumn EdgeColumn
  DirectSCI.SetEdgeMode EdgeMode
  DirectSCI.SetIndentationGuides IndentationGuide
  DirectSCI.UsePopUp ContextMenu
  DirectSCI.SetReadOnly ReadOnly
  DirectSCI.SetEndAtLastLine EndAtLastLine
  DirectSCI.SetEOLMode EOL
  FoldLo = FoldLo
  FoldHi = FoldHi
  SetFoldMarker FoldMarker
  DirectSCI.SetMarginTypeN 0, Gutter0Type
  DirectSCI.SetMarginTypeN 1, Gutter1Type
  DirectSCI.SetMarginTypeN 2, Gutter2Type
  BraceBadFore = BraceBadFore
  BraceMatchFore = BraceMatchFore
  'SetMarginWidthN 0, Gutter0Width
  'SetMarginWidthN 1, Gutter1Width
  'SetMarginWidthN 2, Gutter2Width
  If Folding = True Then
    DirectSCI.SetMarginWidthN 2, Gutter2Width
  End If
  If LineNumbers = True Then
    DirectSCI.SetMarginWidthN 0, Gutter0Width
  End If
  If ShowFlags = True Then
    DirectSCI.SetMarginWidthN 1, Gutter1Width
  End If
  DirectSCI.SetCaretLineVisible LineVisible
  DirectSCI.SetCaretLineBack LineBackColor
  MarkerBack = MarkerBack
  MarkerFore = MarkerFore
  BraceMatchBack = BraceMatchBack
  BraceBadBack = BraceBadBack
  BraceMatchBold = BraceMatchBold
  BraceMatchItalic = BraceMatchItalic
  BraceMatchUnderline = BraceMatchUnderline
  BookmarkBack = BookmarkBack
  BookMarkFore = BookMarkFore
  DirectSCI.SetOvertype OverType
  DirectSCI.SetHScrollBar ScrollBarH
  DirectSCI.SetVScrollBar ScrollBarV
  DirectSCI.SetSelBack True, SelBack
  DirectSCI.SetSelFore True, SelFore
  DirectSCI.SetTabIndents TabIndents
  DirectSCI.SetUseTabs useTabs
  DirectSCI.SetTabWidth m_IndentWidth
  DirectSCI.SetViewEOL ViewEOL
  DirectSCI.SetViewWS ViewWhiteSpace
  DirectSCI.SetWrapMode WordWrap
  Folding = Folding
  ShowFlags = ShowFlags
  LineNumbers = LineNumbers
  InitFolding Folding
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FoldMarker() As FoldingStyle
    FoldMarker = m_FoldMarker
End Property

Public Property Let FoldMarker(ByVal New_FoldMarker As FoldingStyle)
    m_FoldMarker = New_FoldMarker
    PropertyChanged "FoldMarker"
    SetFoldMarker New_FoldMarker
End Property

Private Sub SetFoldMarker(Value As FoldingStyle)
    Select Case Value
    Case 1
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER)
    Case 2
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_CIRCLEMINUS)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_CIRCLEPLUS)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNERCURVE)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_CIRCLEPLUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_CIRCLEMINUSCONNECTED)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNERCURVE)
    Case 3
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_MINUS)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_PLUS)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_EMPTY)
    Case 0
      Call DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_ARROWDOWN)
      Call DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_ARROW)
      Call DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)
      Call DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_EMPTY)
  End Select
End Sub

Private Sub DefineMarker(marknum As Long, Marker As Long)
  Call DirectSCI.MarkerDefine(marknum, Marker)
End Sub

Private Sub InitFolding(EnableIt As Boolean)
  If EnableIt = True Then
    DirectSCI.SetProperty "fold", "1"
    DirectSCI.SetProperty "fold.compact", IIf(FoldCompact, "1", "0")
    DirectSCI.SetProperty "fold.comment", IIf(FoldComment, "1", "0")
    DirectSCI.SetProperty "fold.html", IIf(FoldHTML, "1", "0")
    If FoldAtElse = True Then
      DirectSCI.SetProperty "fold.at.else", "1"
    Else
      DirectSCI.SetProperty "fold.at.else", "0"
    End If
    'SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
    Call SendEditor(SCI_SETMARGINTYPEN, MARGIN_SCRIPT_FOLD_INDEX, SC_MARGIN_SYMBOL)
    Call SendEditor(SCI_SETMARGINMASKN, MARGIN_SCRIPT_FOLD_INDEX, SC_MASK_FOLDERS)
    'SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 20
    Call SendEditor(SCI_SETMARGINSENSITIVEN, MARGIN_SCRIPT_FOLD_INDEX, 1)
    Call SendEditor(SCI_SETFOLDFLAGS, 16, 0)
  Else
    DirectSCI.SetProperty "fold", "0"
    DirectSCI.SetProperty "fold.compact", 0
    DirectSCI.SetProperty "fold.html", "0"
    DirectSCI.SetProperty "fold.comment", "0"
    SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
    Call SendEditor(SCI_SETMARGINSENSITIVEN, MARGIN_SCRIPT_FOLD_INDEX, 0)
  End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FoldAtElse() As Boolean
    FoldAtElse = m_FoldAtElse
End Property

Public Property Let FoldAtElse(ByVal New_FoldAtElse As Boolean)
    m_FoldAtElse = New_FoldAtElse
    PropertyChanged "FoldAtElse"
    If FoldAtElse = True Then
      DirectSCI.SetProperty "fold.at.else", "1"
    Else
      DirectSCI.SetProperty "fold.at.else", "0"
    End If
End Property

Public Property Get FoldComment() As Boolean
    FoldComment = m_FoldComment
End Property

Public Property Let FoldComment(ByVal New_FoldComment As Boolean)
    m_FoldComment = New_FoldComment
    PropertyChanged "FoldComment"
    If FoldComment = True Then
      DirectSCI.SetProperty "fold.comment", "1"
    Else
      DirectSCI.SetProperty "fold.comment", "0"
    End If
End Property

Public Property Get FoldCompact() As Boolean
    FoldCompact = m_FoldCompact
End Property

Public Property Let FoldCompact(ByVal New_Compact As Boolean)
    m_FoldCompact = New_Compact
    PropertyChanged "FoldComment"
    If FoldCompact = True Then
      DirectSCI.SetProperty "fold.compact", "1"
    Else
      DirectSCI.SetProperty "fold.compact", "0"
    End If
End Property

Public Property Get FoldHTML() As Boolean
    FoldHTML = m_FoldHTML
End Property

Public Property Let FoldHTML(ByVal New_FoldHTML As Boolean)
    m_FoldHTML = New_FoldHTML
    PropertyChanged "FoldHTML"
    If FoldHTML = True Then
      DirectSCI.SetProperty "fold.HTML", "1"
    Else
      DirectSCI.SetProperty "fold.HTML", "0"
    End If
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MarginFore() As OLE_COLOR
    MarginFore = m_MarginFore
End Property

Public Property Let MarginFore(ByVal New_MarginFore As OLE_COLOR)
    m_MarginFore = New_MarginFore
    PropertyChanged "MarginFore"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MarginBack() As OLE_COLOR
    MarginBack = m_MarginBack
End Property

Public Property Let MarginBack(ByVal New_MarginBack As OLE_COLOR)
    m_MarginBack = New_MarginBack
    PropertyChanged "MarginBack"
End Property

Private Sub Attach(hWndA As Long)
  Set SC = New cSubclass
  With SC
    .Subclass hWndA, Me
    .AddMsg hWndA, WM_NOTIFY, MSG_AFTER
    .AddMsg hWndA, WM_SETFOCUS, MSG_AFTER
    .AddMsg hWndA, WM_CLOSE, MSG_BEFORE
    .AddMsg hWndA, WM_KEYDOWN, MSG_BEFORE_AFTER
    .Subclass SCI, Me
    .AddMsg SCI, WM_RBUTTONDOWN, MSG_AFTER
    .AddMsg SCI, WM_LBUTTONDOWN, MSG_AFTER
    .AddMsg SCI, WM_KEYDOWN, MSG_BEFORE_AFTER
    .AddMsg SCI, WM_KEYUP, MSG_AFTER
    .AddMsg SCI, WM_LBUTTONUP, MSG_AFTER
    .AddMsg SCI, WM_RBUTTONUP, MSG_AFTER
    .AddMsg SCI, WM_CHAR, MSG_BEFORE
    .AddMsg SCI, WM_COMMAND, MSG_BEFORE
  End With
End Sub

Private Sub Detach()
  SC.UnSubAll
  Set SC = Nothing
End Sub

Public Function GetCurrentLine() As Long
  GetCurrentLine = DirectSCI.LineFromPosition(DirectSCI.GetCurPos)
End Function

Private Function ToLastSpaceCount() As Long
  ' This function will figure out how many characters there are in the currently
  ' selected word.  It gets the line text, finds the position of the caret in
  ' the line text, then converts the line to a byte array to do a faster compare
  ' till it reaches something not interpreted as a letter IE a space or a
  ' line break.  This is kind of overly complex but seems to be faster overall
  
  Dim L As Long, i As Long, current As Long, pos As Long, startWord As Long, iHold As Long
  Dim str As String, bByte() As Byte, strTmp As String
  Dim line As String
  line = GetLineText(GetCurrentLine)
  current = GetCaretInLine
  startWord = current
   
  Str2Byte line, bByte()
  
  iHold = 0
  While (startWord > 0) And InStr(1, CallTipWordCharacters, strTmp) > 0
    startWord = startWord - 1
    iHold = iHold + 1
    If startWord >= 0 Then
      strTmp = Chr(bByte(startWord))
    End If
  Wend
  If strTmp = " " Or strTmp = "." Then iHold = iHold - 1
  ToLastSpaceCount = iHold
End Function


Public Function GetCaretInLine() As Long
  Dim caret As Long, lineStart As Long, line As Long
  caret = DirectSCI.GetCurPos
  line = GetCurrentLine
  lineStart = PositionFromLine(line)
  GetCaretInLine = caret - lineStart
End Function

Private Function SortString(str As String) As String
  Dim ua() As String, X As Long
  ua = Split(str, " ")
  If GetUpper(ua) <> 0 Then
    Call ArraySortString(ua, UBound(ua) + 1)
    SortString = ""
    For X = 0 To UBound(ua)
      SortString = SortString & ua(X) & " "
    Next X
    SortString = Left(SortString, Len(SortString) - 1)
  End If
End Function

Private Sub ArraySortString(ByRef xArray() As String, ByVal xArrayCount As Long)

    Dim xLong1 As Long
    Dim xLong2 As Long
    Dim xLong3 As Long
    Dim xChar1 As String
    Dim xChar2 As String
    xArrayCount = xArrayCount - 1&


    Do
        xLong1 = 3 * xLong1 + 1&
    Loop Until xLong1 > xArrayCount


    Do
        xLong1 = xLong1 \ 3&


        For xLong2 = xLong1 To xArrayCount
            xChar1 = xArray(xLong2)
            xChar2 = UCase(xChar1)


            For xLong3 = xLong2 - xLong1 To 0& Step -xLong1
                If Not UCase(xArray(xLong3)) > xChar2 Then Exit For
                xArray(xLong3 + xLong1) = xArray(xLong3)
            Next

            xArray(xLong3 + xLong1) = xChar1
        Next

    Loop Until xLong1 = 0&

End Sub


Public Sub ShowAutoComplete(strVal As String)
  Dim i As Long
  i = ToLastSpaceCount
  SendMessageString SCI, SCI_AUTOCSHOW, i, SortString(strVal)
End Sub

'+------------------------------------------------------+
'| This function is used to maintain the level of       |
'| indentation.  No values are required.                |
'+------------------------------------------------------+
Private Sub MaintainIndent()
  On Error Resume Next
  Dim g As Long
  Dim indentAmount As Long
  Dim lastLine As Long
  Dim curLine As Long
  g = DirectSCI.GetCurPos
  ' Get the current line
  curLine = GetCurrentLine + 1
  ' Get the previous line
  lastLine = curLine - 1
  
  If GetLineText(lastLine - 1) = "" Then
    'We can move on because in this case there is no text on the
    'previous line.
    Exit Sub
  End If
  indentAmount = 0
  While (lastLine >= 0) And (DirectSCI.GetLineEndPosition(lastLine) - PositionFromLine(lastLine) = 0)
    ' Loop threw the line counting spaces
    lastLine = lastLine - 1
    If lastLine >= 0 Then
      indentAmount = DirectSCI.GetLineIndentation(lastLine)
    End If
    If indentAmount > 0 Then
      Call DirectSCI.SetLineIndentation(curLine - 1, indentAmount)
      Call SetCurrentPosition(GetLineIndentPosition(curLine - 1))
      
      Call DirectSCI.SetSel(DirectSCI.GetCurPos, DirectSCI.GetCurPos)
    End If
  Wend
End Sub

Public Function PositionFromLine(lLine As Long) As Long
  PositionFromLine = SendEditor(SCI_POSITIONFROMLINE, lLine)
End Function

Public Sub SetCurrentPosition(lval As Long)
  SendEditor SCI_SETCURRENTPOS, lval
End Sub

Public Function LoadAPIFile(strFile As String)
  ' This function will load an api file for calltips.
  Dim iFile As Integer, str As String, i As Integer
  iFile = FreeFile
  If FileExists(strFile) = False Then Exit Function
  Erase APIStrings  'Clear the old array
  i = 0
  APIStrings = Split(GetFile(strFile), vbCr)
  For i = 0 To UBound(APIStrings) - 1
    APIStrings(i) = Replace(APIStrings(i), Chr(13), "")
    APIStrings(i) = Replace(APIStrings(i), Chr(10), "")
  Next i
  APIStringLoaded = True
End Function

Public Function AddToAPIFile(ApiFunction As String)
   If Not APIStringLoaded Then
        ReDim Preserve APIStrings(1)
   Else
       ReDim Preserve APIStrings(UBound(APIStrings) + 1)
   End If
    APIStrings(UBound(APIStrings)) = ApiFunction
    APIStringLoaded = True
End Function

Private Function CountOccurancesOfChar(SearchText As String, SearchChar As String) As Integer

Dim lCtr As Integer

CountOccurancesOfChar = 0

  For lCtr = 1 To Len(SearchText)
        If StrComp(Mid(SearchText, lCtr, 1), SearchChar) = 0 Then
            CountOccurancesOfChar = CountOccurancesOfChar + 1
        End If
    Next

End Function

Private Function ReturnPositionOfOcurrance(SearchText As String, SearchChar As String, ByVal pPos As Integer) As Integer
Dim lCtr As Integer
  ReturnPositionOfOcurrance = InStr(1, SearchText, "(") + 1
    
    If pPos <> 0 Then
        For lCtr = InStr(1, SearchText, "(") To Len(SearchText)
        If StrComp(Mid(SearchText, lCtr, 1), SearchChar) = 0 Then
                ReturnPositionOfOcurrance = lCtr
                pPos = pPos - 1
                If pPos = 0 Then
                    Exit Function
                End If
            End If
        Next
        
        ReturnPositionOfOcurrance = InStr(1, SearchText, ")") - 1
    
    End If
  


End Function

Public Sub SetCallTipHighlight(lStart As Long, lEnd As Long)
  SendEditor SCI_CALLTIPSETHLT, lStart, lEnd
End Sub

Public Sub StopCallTip()
  SendEditor SCI_CALLTIPCANCEL
End Sub

Public Sub ShowCallTip(strVal As String)
  Dim bByte() As Byte
  Str2Byte strVal, bByte
  Call SendEditor(SCI_CALLTIPSHOW, DirectSCI.GetCurPos, VarPtr(bByte(0)))
End Sub



Public Function CurrentFunction()

Dim line As String
Dim i As Integer, i2 As Integer, X As Integer
line = GetLineText(GetCurrentLine())
  
  CurrentFunction = ""
  X = GetCaretInLine
  
  For i = X To 1 Step -1
    If Mid(line, i, 1) = "(" Then
        For i2 = i - 1 To 1 Step -1
            If Mid(line, i2, 1) < 33 And CurrentFunction <> "" Then    ' ignore whitespace before (
                Exit For
            Else
                If InStr(1, CallTipWordCharacters, Mid(line, i2, 1)) > 0 Then
                    CurrentFunction = Mid(line, i2, 1) & CurrentFunction
                Else
                    If Asc(Mid(line, i2, 1)) > 33 Then   ' not valid character (and not whitespace)
                        Exit For
                    End If
                End If
            End If
        Next i2
    End If
    
    If CurrentFunction <> "" Then
        Exit For
    End If
  Next i
  
  ' Cant find a function going backwards - check forwards instead ?
  If CurrentFunction = "" Then
    For i = X To Len(line)
        If Mid(line, i, 1) = "(" Then
            For i2 = i - 1 To 1 Step -1
                If Mid(line, i2, 1) < 33 And CurrentFunction <> "" Then    ' ignore whitespace before (
                    Exit For
                Else
                    If InStr(1, CallTipWordCharacters, Mid(line, i2, 1)) > 0 Then
                        CurrentFunction = Mid(line, i2, 1) & CurrentFunction
                    Else
                        If Asc(Mid(line, i2, 1)) > 33 Then   ' not valid character (and not whitespace)
                            Exit For
                        End If
                    End If
                End If
            Next i2
        End If
        If CurrentFunction <> "" Then
            Exit For
    End If
    Next i
  
  End If
  
  End Function
  



Private Sub StartCallTip(ch As Long)
' This entire function is a bit of a hack.  It seems to work but it's very
' messy.  If anyone cleans it up please send me a new version so I can add
' it to this release.  Thanks :)
Dim line As String, PartLine As String, i As Integer, X As Integer
Dim newstr As String, iPos As Integer, iStart As Long, iEnd As Long
Dim a, i2 As Integer

If APIStringLoaded = False Then Exit Sub
If UBound(APIStrings) = 0 Then Exit Sub
  
If ch = Asc("(") Then
  line = GetLineText(GetCurrentLine())
  
  X = GetCaretInLine
  newstr = ""
  
  '
  ' For those compilers that allow whitespace between function and parenthesis
  ' ignore whitespace
  '
  '
  '
    
        For i2 = X - 1 To 1 Step -1
            If Mid(line, i2, 1) < 33 And newstr <> "" Then    ' ignore whitespace before (
                Exit For
            Else
                If InStr(1, CallTipWordCharacters, Mid(line, i2, 1)) > 0 Then
                    newstr = Mid(line, i2, 1) & newstr
                Else
                    If Asc(Mid(line, i2, 1)) > 33 Then   ' not valid character (and not whitespace)
                        Exit For
                    End If
                End If
            End If
        Next i2
  
    If Len(newstr) = 0 Then   ' blank line ?
      StopCallTip
      Exit Sub
    End If
    
    newstr = newstr & "("    ' make it into a function name so no partial searches of other API functions
  
  ' Lookup the Function name in the API list
    If GetUpper(APIStrings) > 0 Then
      For i = 0 To UBound(APIStrings)
        If InStr(1, LCase$(APIStrings(i)), LCase$(newstr)) <> 0 Then ' case insensitive string
                    
            ActiveCallTip = i
            
            iPos = InStr(1, APIStrings(i), ")")
            ShowCallTip Left$(APIStrings(i), iPos) ' to end of function
            
            iPos = InStr(1, APIStrings(i), ",")
            If iPos > 0 Then
                iStart = Len(newstr)
                iEnd = iPos - 1
                SetCallTipHighlight iStart, iEnd
                Exit For
            Else
                ' single parameter ?
                If Len(newstr) + 1 <> Len(APIStrings(i)) Then
                    iStart = Len(newstr)
                    iEnd = Len(APIStrings(i)) - 1
                    SetCallTipHighlight iStart, iEnd
                    Exit For
                End If
            End If
        End If
      Next
    End If
    Exit Sub
End If
  
' Do we have a tip already active ?
If DirectSCI.CallTipActive Then
    If ch = Asc(")") Then
        StopCallTip
    Else
        ' are we still in the current tooltip ?
        line = GetLineText(GetCurrentLine())
        X = GetCaretInLine
        iPos = InStrRev(line, "(", X)
        PartLine = Mid(line, iPos + 1, X - iPos) 'Get the chunk of the string were in
        
        If InStr(1, APIStrings(ActiveCallTip), ",") = 0 Then   ' only one param
            iStart = InStr(1, APIStrings(ActiveCallTip), "(") - 1
            iEnd = InStr(1, APIStrings(ActiveCallTip), ")") - 1
        Else
           
            'Count which param
            iPos = CountOccurancesOfChar(PartLine, ",")
            'Highlight Param in calltip
            iStart = ReturnPositionOfOcurrance(APIStrings(ActiveCallTip), ",", iPos) - 1
            iEnd = ReturnPositionOfOcurrance(APIStrings(ActiveCallTip), ",", iPos + 1)
        End If
        SetCallTipHighlight iStart, iEnd
  End If
End If
End Sub





Public Function GetLineIndentPosition(lLine As Long) As Long
  GetLineIndentPosition = SendEditor(SCI_GETLINEINDENTPOSITION, lLine)
End Function



Private Function IsBrace(ch As Long) As Boolean
    IsBrace = (ch = 40 Or ch = 41 Or ch = 60 Or ch = 62 Or ch = 91 Or ch = 93 Or ch = 123 Or ch = 125)
End Function

Private Function MatchBrace(ch As String) As String
  If ch = "<" Then MatchBrace = ">"
  If ch = "(" Then MatchBrace = ")"
  If ch = "[" Then MatchBrace = "]"
  If ch = "{" Then MatchBrace = "}"
End Function

Public Sub LoadFile(strFile As String)
  Dim str As String
  'isRead = readOnly
  'If isRead = True Then readOnly = False
  If Dir(strFile) = "" Then Exit Sub  'We don't want to have an error if the file doesn't exist.
  str = GetFile(strFile)
  'GetFile2 strFile
  'SetText ""
  DirectSCI.SetText str
  'AddText Len(str), str
  ClearUndoBuffer
  DirectSCI.ConvertEOLs DirectSCI.GetEOLMode
  DirectSCI.SetFocus
  DirectSCI.GotoPos 0
  DirectSCI.SetSavePoint
End Sub


Private Function GetFile(strFilePath As String, Optional bolAsString = True) As String
  On Error Resume Next
  Dim arrFileMain() As Byte
  Dim arrFileBuffer() As Byte
  Dim lngAllBytes As Long
  Dim lngSize As Long, lngRet As Long
  Dim i As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT
  Const lngMaxSizeForOneStep = 1000000
    'Prepare Arrays ==========================================================
    ReDim arrFileMain(0)
    ReDim arrFileBuffer(lngMaxSizeForOneStep)

    'Open the two files
    lngFileHandle = OpenFile(strFilePath, ofData, OF_READ)

    'Get the file size
    lngSize = GetFileSize(lngFileHandle, 0)
    Do While Not UBound(arrFileMain) = lngSize - 1
        If lngSize = 0 Then Exit Function

        'Redim Array to fit a smaller file
        lngAllBytes = UBound(arrFileMain)
        If lngSize - lngAllBytes < lngMaxSizeForOneStep Then ReDim arrFileBuffer(lngSize - lngAllBytes - 2)

        'Read from the file
        ReadFile lngFileHandle, arrFileBuffer(0), UBound(arrFileBuffer) + 1, lngRet, ByVal 0&

        'Calculate Buffer's position in Main Array
        If lngAllBytes > 0 Then lngAllBytes = lngAllBytes + 1

        'Make place for the Buffer in the Main Array
        ReDim Preserve arrFileMain(lngAllBytes + UBound(arrFileBuffer))

        'Put Buffer at end of Main Array
        MemCopy arrFileMain(lngAllBytes), arrFileBuffer(0), UBound(arrFileBuffer) + 1

        DoEvents

    Loop

    'Close the file
    CloseHandle lngFileHandle
    ReDim arrFileBuffer(0)

    'Convert Main Array to String
    GetFile = StrConv(arrFileMain(), vbUnicode)
End Function

Public Sub ClearUndoBuffer()
  SendEditor SCI_EMPTYUNDOBUFFER
End Sub

Public Sub ToggleMarker()
  On Error Resume Next
  If GetMarker(GetCurrentLine) = 4 Then
    DeleteMarker GetCurrentLine, 2
  Else
    MarkerSet GetCurrentLine, 2
  End If
End Sub

Private Function GetMarker(iLine As Long) As Long
  GetMarker = SendEditor(SCI_MARKERGET, iLine)
End Function

Private Sub DeleteMarker(iLine As Long, marknum As Long)
  SendEditor SCN_MARKERDELETE, iLine, marknum
End Sub

Private Sub NextMarker(lLine As Long, marknum As Long)
  Dim X As Long
  X = SendEditor(SCN_MARKERNEXT, lLine, marknum)
  If X = -1 Then
    X = SendEditor(SCN_MARKERNEXT, 0, marknum)
  End If
  DirectSCI.GotoLine X
End Sub

Private Sub PrevMarker(lLine As Long, marknum As Long)
  Dim X As Long
  X = SendEditor(SCN_MARKERPREVIOUS, lLine, marknum)
  If X = -1 Then
    X = SendEditor(SCN_MARKERPREVIOUS, DirectSCI.GetLineCount, marknum)
  End If
  DirectSCI.GotoLine X
End Sub

Private Sub DeleteAllMarker(marknum As Long)
  SendEditor SCN_MARKERDELETEALL, marknum
End Sub


Public Sub NextBookmark()
  NextMarker GetCurrentLine + 1, 4
End Sub

Public Sub PrevBookmark()
  PrevMarker GetCurrentLine - 1, 4
End Sub

Public Sub ClearBookmarks()
  DeleteAllMarker 2
End Sub

Public Sub MarkerSet(iLine As Long, iMarkerNum As Long)
  SendEditor SCI_MARKERADD, iLine, iMarkerNum
End Sub

Public Sub SaveToFile(strFile As String)
  Dim str As String
  str = DirectSCI.GetText
  WriteToFile strFile, str
  ' Remove the modified flag from scintilla
  DirectSCI.SetSavePoint
  If ClearUndoAfterSave Then ClearUndoBuffer
End Sub

Private Sub StartAutoComplete(ch As Long)
  If Len(AutoCompleteStart) > 1 Then Exit Sub
  If ch = Asc(AutoCompleteStart) Then
    ShowAutoComplete AutoCompleteString
  End If
End Sub

Private Sub WriteToFile(strFile As String, strdata As String)
  On Error GoTo eHandle
  Dim i As Long
  Dim L As Long
  Dim hFile As Long
  Dim bByte() As Byte
  ConvertEOLMode
  Str2Byte strdata, bByte()
  L = UBound(bByte()) - 1
  hFile = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0&)
  WriteFile hFile, bByte(0), L, 0, ByVal 0&
  CloseHandle hFile
  Exit Sub
eHandle:
  ' Just in case anything happens let's close the handle
  CloseHandle hFile
End Sub

Public Function ConvertEOLMode()
  SendEditor SCI_CONVERTEOLS, DirectSCI.GetEOLMode
End Function
Public Sub PrintDoc()
  PrintSCI SCI, DirectSCI.GetTextLength, 1000, 1000, 1000, 1000
End Sub

Private Sub PrintSCI(sciHwnd As Long, txtLen As Long, LeftMarginWidth As Long, _
   TopMarginHeight, RightMarginWidth, BottomMarginHeight)
   Dim LeftOffset As Long, TopOffset As Long
   Dim LeftMargin As Long, TopMargin As Long
   Dim RightMargin As Long, BottomMargin As Long
   Dim fr As FormatRange
   Dim rcDrawTo As RECT
   Dim rcPage As RECT
   Dim TextLength As Long
   Dim NextCharPosition As Long
   Dim R As Long
   Dim PhysWidth As Long, PhysHeight As Long
   Dim PrintWidth As Long, PrintHeight As Long
   Dim ptDPI As POINTAPI, ptPage As POINTAPI
   Dim rectPhysMargins As RECT, rectMargins As RECT, rectSetup As RECT
   Printer.Print Space(1)
   Printer.ScaleMode = vbPixels


   ' Get the offsett to the printable area on the page in twips
   ptDPI.X = GetDeviceCaps(Printer.hDC, LOGPIXELSX)
   ptDPI.Y = GetDeviceCaps(Printer.hDC, LOGPIXELSY)
   ptPage.X = GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
   ptPage.Y = GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
   
   rectPhysMargins.Left = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX)
   rectPhysMargins.Top = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY)
   rectPhysMargins.Right = ptPage.X - GetDeviceCaps(Printer.hDC, HORZRES) - rectPhysMargins.Left
   rectPhysMargins.Bottom = ptPage.Y - GetDeviceCaps(Printer.hDC, VERTRES) - rectPhysMargins.Top
   
   rectSetup.Left = MulDiv(LeftMarginWidth, ptDPI.X, 1000)
   rectSetup.Top = MulDiv(TopMarginHeight, ptDPI.Y, 1000)
   rectSetup.Right = MulDiv(RightMarginWidth, ptDPI.X, 1000)
   rectSetup.Bottom = MulDiv(BottomMarginHeight, ptDPI.Y, 1000)
    
   rectMargins.Left = Max(rectPhysMargins.Left, rectSetup.Left)
   rectMargins.Top = Max(rectPhysMargins.Top, rectSetup.Top)
   rectMargins.Right = Max(rectPhysMargins.Right, rectSetup.Right)
   rectMargins.Bottom = Max(rectPhysMargins.Bottom, rectSetup.Bottom)
    
   ' Calculate the Left, Top, Right, and Bottom margins
   'LeftMargin = (LeftMarginWidth - LeftOffset) \ Printer.TwipsPerPixelX
   'TopMargin = (TopMarginHeight - TopOffset) \ Printer.TwipsPerPixelY
   'RightMargin = (((Printer.Width - RightMarginWidth) - LeftOffset) \ Printer.TwipsPerPixelX) + (LeftMargin + LeftOffset)
   'BottomMargin = (((Printer.Height - BottomMarginHeight) - TopOffset) \ Printer.TwipsPerPixelY) + (TopMargin + TopOffset)

   ' Set printable area rect
   'rcPage.Left = 0
   'rcPage.Top = 0
   'rcPage.Right = Printer.ScaleWidth
   'rcPage.Bottom = Printer.ScaleHeight

   ' Set rect in which to print (relative to printable area)
   'rcDrawTo.Left = LeftMargin
   'rcDrawTo.Top = TopMargin
   'rcDrawTo.Right = RightMargin
   'rcDrawTo.Bottom = BottomMargin
   
   'rcPage = rcDrawTo
   ' Set up the print instructions
   
  fr.rc.Left = rectMargins.Left - rectPhysMargins.Left
  fr.rc.Top = rectMargins.Top - rectPhysMargins.Top
  fr.rc.Right = ptPage.X - rectMargins.Right - rectPhysMargins.Left
  fr.rc.Bottom = ptPage.Y - rectMargins.Bottom - rectPhysMargins.Top
  fr.rcPage.Left = 0
  fr.rcPage.Top = 0
  fr.rcPage.Right = ptPage.X - rectPhysMargins.Left - rectPhysMargins.Right - 1
  fr.rcPage.Bottom = ptPage.Y - rectPhysMargins.Top - rectPhysMargins.Bottom - 1

   fr.chrg.cpMin = 0           ' Indicate start of text through
'
   fr.chrg.cpMax = txtLen          ' end of the text

   ' Get length of text in RTF
   TextLength = txtLen
   'NextCharPosition = SendMessage2(SCIHwnd, SCI_FORMATRANGE, True, fr)
   NextCharPosition = 0
   Do
      'Printer.NewPage                  ' Move on to next page
      Printer.Print Space(1) ' Re-initialize hDC
      fr.hDC = Printer.hDC
      fr.hdcTarget = Printer.hDC
      fr.chrg.cpMin = NextCharPosition ' Starting position for next page
      fr.chrg.cpMax = txtLen
      NextCharPosition = SendMessage2(sciHwnd, SCI_FORMATRANGE, True, fr)
      If NextCharPosition >= txtLen Then Exit Do
      Printer.NewPage
      
   Loop
'
'   ' Commit the print job
   Printer.EndDoc
'
'   ' Allow the RTF to free up memory
   R = SendMessage2(sciHwnd, SCI_FORMATRANGE, False, ByVal CLng(0))
End Sub


Private Sub Str2Byte(sInput As String, bOutput() As Byte)
  ' This function is used to convert strings to bytes
  ' This comes in handy for saving the file.  It's also
  ' useful when dealing with certain things related to
  ' sending info to Scintilla
  
  Dim i As Long
  ReDim bOutput(Len(sInput))

  For i = 0 To Len(sInput) - 1
    bOutput(i) = Asc(Mid(sInput, i + 1, 1))
  Next i
  bOutput(UBound(bOutput)) = 0  ' Null terminated :)
End Sub

Public Sub GotoLineColumn(iLine As Long, iCol As Long)
  Dim i As Long
  i = SendEditor(SCI_FINDCOLUMN, iLine, iCol)
  DirectSCI.SetSel i, i
End Sub

Public Function ReplaceText(strSearchFor As String, strReplaceWith As String, Optional ReplaceAll As Boolean = False, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Boolean
  bRepLng = True
  If FindText(strSearchFor, False, False, True, CaseSensative, WordStart, WholeWord) = True Then
    DirectSCI.ReplaceSel strReplaceWith
    If ReplaceAll Then
      bRepAll = True
      Do Until FindText(strSearchFor, False, False, True, CaseSensative, WordStart, WholeWord) = False
        DirectSCI.ReplaceSel strReplaceWith
      Loop
      bRepAll = False
    End If
  End If
  bRepLng = False
End Function

Public Function ReplaceAll(strSearchFor As String, strReplaceWith As String, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Long
  ReplaceAll = 0
  Dim lval As Long
  Dim lenSearch As Long, lenReplace As Long
  Dim Find As Long
  If strSearchFor = "" Then Exit Function
  lval = 0
  If CaseSensative Then
    lval = lval Or SCFIND_MATCHCASE
  End If
  If WordStart Then
    lval = lval Or SCFIND_WORDSTART
  End If
  If WholeWord Then
    lval = lval Or SCFIND_WHOLEWORD
  End If
  If RegExp Then
    lval = lval Or SCFIND_REGEXP
  End If
  Dim targetstart As Long, targetend As Long, pos As Long, docLen As Long
  targetstart = 0
  docLen = DirectSCI.GetTextLength
  lenSearch = Len(strSearchFor)
  lenReplace = Len(strReplaceWith)
  
  targetend = docLen
  Call SendEditor(SCI_SETSEARCHFLAGS, lval)
  Call SendEditor(SCI_SETTARGETSTART, targetstart)
  Call SendEditor(SCI_SETTARGETEND, targetend)
  Find = SendMessageString(SCI, SCI_SEARCHINTARGET, lenSearch, strSearchFor)
  Do Until Find = -1
    targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
    targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
    
    DirectSCI.ReplaceTarget lenReplace, strReplaceWith
    targetstart = targetstart + lenReplace
    targetend = docLen
    ReplaceAll = ReplaceAll + 1
    Call SendEditor(SCI_SETTARGETSTART, targetstart)
    Call SendEditor(SCI_SETTARGETEND, targetend)
    Find = SendMessageString(SCI, SCI_SEARCHINTARGET, lenSearch, strSearchFor)
  Loop
End Function


Public Function FindText(txttofind As String, Optional FindReverse As Boolean = False, Optional ByVal findinrng As Boolean, Optional WrapDocument As Boolean = True, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Boolean
  Dim lval As Long, Find As Long
  ' Sending a null string to scintilla for the find text willc ause errors!
  If txttofind = "" Then Exit Function
  lval = 0
  If CaseSensative Then
    lval = lval Or SCFIND_MATCHCASE
  End If
  If WordStart Then
    lval = lval Or SCFIND_WORDSTART
  End If
  If WholeWord Then
    lval = lval Or SCFIND_WHOLEWORD
  End If
  If RegExp Then
    lval = lval Or SCFIND_REGEXP
  End If
  Dim targetstart As Long, targetend As Long, pos As Long
    Call SendEditor(SCI_SETSEARCHFLAGS, lval)
    If findinrng Then
        targetstart = SendMessage(SCI, SCI_GETSELECTIONSTART, CLng(0), CLng(0))
        targetend = SendMessage(SCI, SCI_GETSELECTIONEND, CLng(0), CLng(0))
    Else
      If FindReverse = False Then
        targetstart = SendMessage(SCI, SCI_GETSELECTIONEND, 0, 0)
        targetend = Len(Text)
      Else
        targetstart = SendMessage(SCI, SCI_GETSELECTIONSTART, 0, 0)
        targetend = 0
      End If
    End If
    ' Creamos una región de búsqueda (que puede ser el texto completo)
    Call SendEditor(SCI_SETTARGETSTART, targetstart)
    Call SendEditor(SCI_SETTARGETEND, targetend)
    Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
    ' Seleccionamos lo que se ha encontrado
    If Find > -1 Then

        targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
        targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
        DirectSCI.SetSel targetstart, targetend
    Else
      If WrapDocument Then
        If FindReverse = False Then
          targetstart = 0
          targetend = Len(Text)
        Else
          targetstart = Len(Text)
          targetend = 0
        End If
        Call SendEditor(SCI_SETTARGETSTART, targetstart)
        Call SendEditor(SCI_SETTARGETEND, targetend)
        Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
        If Find > -1 Then
          targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
          targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
          DirectSCI.SetSel targetstart, targetend
        End If
      End If
    End If
    
  ' A find has been performed so now FindNext will work.
  bFindEvent = True
  If Find > -1 Then
    FindText = True
  Else
    FindText = False
  End If
    
  ' Set the info that we've used so we findnext can send the same thing
  ' out if called.
  
    bWrap = WrapDocument
    bCase = CaseSensative
    bWholeWord = WholeWord
    bRegEx = RegExp
    bWordStart = WordStart
    bFindInRange = findinrng
    bFindReverse = FindReverse
    strFind = txttofind
  
End Function

Public Sub ShowAbout()
    frmAbout.show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub

Public Function FindNext() As Boolean
  'If no find events have occurred exit this sub or it may cause errors.
  If bFindEvent = False Then Exit Function
  FindNext = FindText(strFind, False, bFindInRange, bWrap, bCase, bWordStart, bWholeWord, bRegEx)
End Function

Public Function FindPrev() As Boolean
  If bFindEvent = False Then Exit Function
  FindPrev = FindText(strFind, True, bFindInRange, bWrap, bCase, bWordStart, bWholeWord, bRegEx)
End Function


Public Function GetLineText(lLine As Long) As String
  'On Error Resume Next
  Dim txt As String
  Dim lLength As Long
  Dim i As Long
  Dim bByte() As Byte
  lLength = SendMessage(SCI, SCI_LINELENGTH, lLine, 0)
  lLength = lLength - 1 'By default this will tag on Chr(10) + chr(13)
  If lLength > 1 Then
    ReDim bByte(0 To lLength)
    SendMessage SCI, SCI_GETLINE, lLine, VarPtr(bByte(0))
    
    txt = Byte2Str(bByte())
  Else
    txt = ""  'This line is 0 length
  End If
  GetLineText = txt
End Function

Public Sub DoReplace()
  Load frmReplace
  With frmReplace
    Set .cScintilla = Me
    If SelText <> "" Then .cmbFind.Text = SelText
    .show
  End With
End Sub

Public Sub DoGoto()
  Load frmGoto
  Dim iLine As Long, iCol As Long
  With frmGoto
    .lblCurLine.Caption = "Current Line: " & GetCurrentLine + 1
    .lblLineCount.Caption = "Last Line: " & DirectSCI.GetLineCount
    .lblColumn.Caption = "Column: " & DirectSCI.GetColumn
    .show vbModal
    If .iWhatToDo = 1 Then
      If .txtLine.Text = "" Then .txtLine.Text = 1
      If .txtColumn.Text = "" Then .txtColumn.Text = 1
      iLine = .txtLine.Text
      iCol = .txtColumn.Text
      GotoLineColumn iLine - 1, iCol - 1
    End If
  End With
  Unload frmGoto
  SetFocus
End Sub


Public Sub DoFind()
  Dim bFind As Boolean
  Dim fFind As frmFind
  Set fFind = New frmFind
  Load fFind
  With fFind
    If SelText <> "" Then
      .cmbFind.Text = SelText
      .txtFind.Text = SelText
    End If
      
    .show vbModal
    If .DoWhat = 0 Then
      SetFocus
      Exit Sub
    ElseIf .DoWhat = 1 Then
      If .bMulti = False Then
        bFind = FindText(.cmbFind.Text, .optUp.Value, False, .chkWrap.Value, .chkCase.Value, False, .chkWhole.Value, .chkRegExp.Value)
      Else
        ' First we must conver the line ends in the search field
        ' to match scintilla's current line endings.
        Select Case DirectSCI.GetEOLMode
          Case SC_EOL_CRLF
            ' Do nothing.  VB's textboxes are CRLF
          Case SC_EOL_CR
            ' Replace LF's in document with nothing:
            .txtFind.Text = Replace(.txtFind.Text, vbLf, "")
          Case sc_eol_lf
            ' Replace CR's in document with nothing:
            .txtFind.Text = Replace(.txtFind.Text, vbCr, "")
        End Select
        
        ' Now the text box text will have the line endings which
        ' this scintilla document is currently using.  Failure
        ' to do this will cause it to not detect a line break
        ' because there won't be say an LF if it's only using
        ' CR as it's line break mode.
        bFind = FindText(.txtFind.Text, .optUp.Value, False, .chkWrap.Value, .chkCase.Value, False, .chkWhole.Value, .chkRegExp.Value)
      End If
      If bFind = False Then RaiseEvent FindFailed(.cmbFind.Text)
    ElseIf .DoWhat = 2 Then
      If .bMulti = False Then
        MarkAll .cmbFind.Text
      Else
        ' First we must conver the line ends in the search field
        ' to match scintilla's current line endings.
        Select Case DirectSCI.GetEOLMode
          Case SC_EOL_CRLF
            ' Do nothing.  VB's textboxes are CRLF
          Case SC_EOL_CR
            ' Replace LF's in document with nothing:
            .txtFind.Text = Replace(.txtFind.Text, vbLf, "")
          Case sc_eol_lf
            ' Replace CR's in document with nothing:
            .txtFind.Text = Replace(.txtFind.Text, vbCr, "")
        End Select
        
        ' Now the text box text will have the line endings which
        ' this scintilla document is currently using.  Failure
        ' to do this will cause it to not detect a line break
        ' because there won't be say an LF if it's only using
        ' CR as it's line break mode.
        MarkAll .txtFind.Text
      End If
      ' This will be in a future release
    End If
    Unload fFind
  End With
  SetFocus
  
End Sub


Public Sub MarkAll(strFind As String)
  Dim X As Long
  Dim g As Boolean
  Dim bFind As Long
  X = DirectSCI.GetCurPos
  DirectSCI.SetSel 0, 0
  Call SendEditor(SCI_SETTARGETSTART, 0)
  Call SendEditor(SCI_SETTARGETEND, DirectSCI.GetTextLength)
  bFind = DirectSCI.SearchInTarget(Len(strFind), strFind)
  'bFind = FindText(strFind, False, False, False, False, False, False, False)
  g = True
  Do While bFind > 0
    
    ' Save some time here.  Since were marking all instances if the same
    ' string is found twice in the same line we don't need to know that.
    ' So once we find it in a line and mark it automaticly jump to the next
    ' line
    
    DirectSCI.GotoPos bFind
    MarkerSet GetCurrentLine, 2
    DirectSCI.GotoLine GetCurrentLine + 1
    Call SendEditor(SCI_SETTARGETSTART, DirectSCI.GetCurPos)
    Call SendEditor(SCI_SETTARGETEND, DirectSCI.GetTextLength)
    bFind = DirectSCI.SearchInTarget(Len(strFind), strFind)
  Loop
  DirectSCI.SetSel X, X
End Sub


'+++++++++ The following functions are for loading, saving, recording
'+++++++++ and playing macro files.
Public Sub StartMacroRecord()
  Erase Macro
  SendEditor SCI_STARTRECORD
End Sub

Public Sub StopMacroRecord()
  SendEditor SCI_STOPRECORD
End Sub

Private Sub HandleMacroCall(iMsg As Long, ch As String)
  If iMsg = SCI_CUT Or iMsg = SCI_COPY Or iMsg = SCI_PASTE Or iMsg = SCI_CLEAR Or iMsg = SCI_ADDTEXT Or iMsg = SCI_REPLACESEL Or iMsg = SCI_DELETEBACK Or iMsg = SCI_CHARLEFT Or iMsg = SCI_CHARRIGHT Then
    AddMacroMsg iMsg, ch
  End If
End Sub

Private Sub AddMacroMsg(iMsg, ch As String)
  On Error Resume Next
  Dim L As Long
  If GetUpper(Macro) <> 0 Then
    L = UBound(Macro)
  Else
    L = 0
  End If
  ReDim Preserve Macro(L + 1)
  Macro(L).lMsg = iMsg
  Macro(L).strChar = ch
End Sub

Public Sub SaveMacro(strFile As String)
  On Error GoTo errHandle
  
  Dim lFile As Integer
  lFile = FreeFile
  Dim i As Long
  If GetUpper(Macro) <> 0 Then
    If UBound(Macro) > 0 Then
      Open strFile For Output As #lFile
        For i = 0 To UBound(Macro) - 1
          Write #lFile, Macro(i).lMsg & "æ" & Macro(i).strChar
        Next i
      Close #lFile
    End If
  End If
errHandle:
  'Just exit the sub.  The only reason this should ever fail is if
  'the macro array is null.
End Sub

Public Sub PlayMacro()
  On Error Resume Next
  If UBound(Macro) = 0 Then Exit Sub
  Dim i As Long
  If GetUpper(Macro) <> 0 Then
    For i = 0 To UBound(Macro) - 1
      Select Case Macro(i).lMsg
        Case SCI_REPLACESEL
          DirectSCI.ReplaceSel Macro(i).strChar
        Case SCI_DELETEBACK
          SendEditor SCI_DELETEBACK
        Case SCI_PASTE
          DirectSCI.Paste
        Case SCI_CHARLEFT
          SendEditor SCI_CHARLEFT
        Case SCI_CHARRIGHT
          SendEditor SCI_CHARRIGHT
      End Select
    Next i
  End If
  SetFocus
End Sub

Public Sub LoadMacro(strFile As String)
  'On Error Resume Next
  SetFocus
  Erase Macro   ' This way if it attempts loading a non existent macro
                ' and then playing it we don't end up playing the wrong
                ' macro
  If FileExists(strFile) = False Then Exit Sub
  Dim lFile As Integer
  Dim p As Long, ch As String
  Dim str As String
  Dim d() As String
  lFile = FreeFile
  
  
  Open strFile For Input As #lFile
    Do While Not EOF(lFile)
      Input #lFile, str
      d = Split(str, "æ")
       p = d(0)
      ch = d(1)
      HandleMacroCall p, ch
    Loop
  Close #lFile
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SelStart() As Long
    SelStart = DirectSCI.GetSelectionStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    m_SelStart = New_SelStart
    PropertyChanged "SelStart"
    DirectSCI.SetSelectionStart New_SelStart
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SelEnd() As Long
    SelEnd = DirectSCI.GetSelectionEnd
End Property

Public Property Let SelEnd(ByVal New_SelEnd As Long)
    m_SelEnd = New_SelEnd
    PropertyChanged "SelEnd"
    DirectSCI.SetSelectionEnd New_SelEnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function GotoLine(line As Long) As Long
  DirectSCI.GotoLine line
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function GotoCol(Column As Long) As Long
  GotoLineColumn GetCurrentLine, Column
End Function

Public Function SetFocus() As Long
  DirectSCI.SetFocus
End Function

Public Function Redo() As Long
  DirectSCI.Redo
End Function

Public Function Undo() As Long
  DirectSCI.Undo
End Function

Public Function Cut() As Long
  DirectSCI.Cut
End Function

Public Function Copy() As Long
  DirectSCI.Copy
End Function

Public Function Paste() As Long
  DirectSCI.Paste
End Function

Public Function SelectAll() As Long
  DirectSCI.SelectAll
End Function

Public Function SelectLine() As Long
  DirectSCI.SetSel PositionFromLine(GetCurrentLine), DirectSCI.GetLineEndPosition(GetCurrentLine)
End Function

Public Function SetSavePoint() As Long
  DirectSCI.SetSavePoint
End Function

Public Function GetColumn() As Long
  GetColumn = DirectSCI.GetColumn
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblue
Public Property Get BraceMatchFore() As OLE_COLOR
    BraceMatchFore = m_BraceMatch
End Property

Public Property Let BraceMatchFore(ByVal New_BraceMatch As OLE_COLOR)
    m_BraceMatch = New_BraceMatch
    PropertyChanged "BraceMatch"
    DirectSCI.StyleSetFore 34, New_BraceMatch
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbred
Public Property Get BraceBadFore() As OLE_COLOR
    BraceBadFore = m_BraceBad
End Property

Public Property Let BraceBadFore(ByVal New_BraceBad As OLE_COLOR)
    m_BraceBad = New_BraceBad
    PropertyChanged "BraceBad"
    DirectSCI.StyleSetFore 35, New_BraceBad
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get BraceMatchBold() As Boolean
    BraceMatchBold = m_BraceMatchBold
End Property

Public Property Let BraceMatchBold(ByVal New_BraceMatchBold As Boolean)
    m_BraceMatchBold = New_BraceMatchBold
    PropertyChanged "BraceMatchBold"
    DirectSCI.StyleSetBold 35, New_BraceMatchBold
    DirectSCI.StyleSetBold 34, New_BraceMatchBold
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BraceMatchItalic() As Boolean
    BraceMatchItalic = m_BraceMatchItalic
End Property

Public Property Let BraceMatchItalic(ByVal New_BraceMatchItalic As Boolean)
    m_BraceMatchItalic = New_BraceMatchItalic
    PropertyChanged "BraceMatchItalic"
    DirectSCI.StyleSetItalic 35, New_BraceMatchItalic
    DirectSCI.StyleSetItalic 34, New_BraceMatchItalic
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BraceMatchUnderline() As Boolean
    BraceMatchUnderline = m_BraceMatchUnderline
End Property

Public Property Let BraceMatchUnderline(ByVal New_BraceMatchUnderline As Boolean)
    m_BraceMatchUnderline = New_BraceMatchUnderline
    PropertyChanged "BraceMatchUnderline"
    DirectSCI.StyleSetUnderline 35, New_BraceMatchUnderline
    DirectSCI.StyleSetUnderline 34, New_BraceMatchUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get BraceMatchBack() As OLE_COLOR
    BraceMatchBack = m_BraceMatchBack
End Property

Public Property Let BraceMatchBack(ByVal New_BraceMatchBack As OLE_COLOR)
    m_BraceMatchBack = New_BraceMatchBack
    PropertyChanged "BraceMatchBack"
    DirectSCI.StyleSetBack 34, New_BraceMatchBack
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get BraceBadBack() As OLE_COLOR
    BraceBadBack = m_BraceBadBack
End Property

Public Property Let BraceBadBack(ByVal New_BraceBadBack As OLE_COLOR)
    m_BraceBadBack = New_BraceBadBack
    PropertyChanged "BraceBadBack"
    DirectSCI.StyleSetBack 35, New_BraceBadBack
End Property

Public Function SendEditor(ByVal Msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal lParam = 0) As Long
    If VarType(lParam) = vbString Then
        SendEditor = SendMessageString(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), CStr(lParam))
    Else
        SendEditor = SendMessage(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), IIf(lParam = 0, CLng(lParam), lParam))
    End If
End Function

Public Sub FoldAll()
  Dim MaxLine As Long, LineSeek As Long
  MaxLine = DirectSCI.GetLineCount
  DirectSCI.Colourise 0, -1
  For LineSeek = 0 To MaxLine - 1
    If DirectSCI.GetFoldLevel(LineSeek) And SC_FOLDLEVELHEADERFLAG Then
      DirectSCI.ToggleFold LineSeek
    End If
  Next
  DirectSCI.ShowLines 0, 0
End Sub

Public Sub TabRight()
  SendEditor SCI_TAB
End Sub

Public Sub TabLeft()
  SendEditor SCI_BACKTAB
End Sub

