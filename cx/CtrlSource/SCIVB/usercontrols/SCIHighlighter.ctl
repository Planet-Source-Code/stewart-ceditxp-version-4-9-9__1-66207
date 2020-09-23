VERSION 5.00
Begin VB.UserControl SCIHighlighter 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   InvisibleAtRuntime=   -1  'True
   Picture         =   "SCIHighlighter.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "SCIHighlighter.ctx":0C42
End
Attribute VB_Name = "SCIHighlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_AutoCloseBraces = 0
Const m_def_AutoCloseQuotes = 0
Const m_def_TabIndents = 0
Const m_def_BackSpaceUnIndents = 0
Const m_def_BookmarkBack = vbBlack
Const m_def_BookMarkFore = vbWhite
Const m_def_MarkerBack = vbBlack
Const m_def_MarkerFore = vbWhite
Const m_def_TabWidth = 4
Const m_def_CaretForeColor = vbBlack
Const m_def_CaretWidth = 1
Const m_def_EdgeColor = &HE0E0E0
Const m_def_EOLMode = 0
Const m_def_HighlightBraces = 1
Const m_def_ClearUndoAfterSave = 1
Const m_def_EndAtLastLine = 0
Const m_def_MaintainIndentation = 1
Const m_def_OverType = 0
'Property Variables:
Dim m_AutoCloseBraces As Boolean
Dim m_AutoCloseQuotes As Boolean
Dim m_TabIndents As Boolean
Dim m_BackSpaceUnIndents As Boolean
Dim m_BookmarkBack As OLE_COLOR
Dim m_BookMarkFore As OLE_COLOR
Dim m_MarkerBack As OLE_COLOR
Dim m_MarkerFore As OLE_COLOR
Dim m_TabWidth As Long
Dim m_CaretForeColor As OLE_COLOR
Dim m_CaretWidth As Long
Dim m_EdgeColor As OLE_COLOR
Dim m_EOLMode As Long
Dim m_HighlightBraces As Boolean
Dim m_ClearUndoAfterSave As Boolean
Dim m_EndAtLastLine As Boolean
Dim m_MaintainIndentation As Boolean
Dim m_OverType As Boolean

Event ClearHighlighters()   'This event will be called then the control is clearing it's highlighter array
Event HighlighterSet(HighlighterName As String)
Event AddHighlighter(HighlighterName As String, Filter As String)

Private Sub UserControl_Initialize()
  InitLexerInterface
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 32 * Screen.TwipsPerPixelX
  UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub

Public Sub SetHighlighter(Scintilla As SCIVB, HighlighterName As String)
  If hlCount = 0 Then Exit Sub
  SetHighlighters Scintilla, HighlighterName, Scintilla.MarginBack, Scintilla.MarginFore
End Sub

Public Sub LoadHighlighters(Path As String)
  On Error Resume Next ' Just in case there were no highlighters loaded
  RaiseEvent ClearHighlighters
  Dim i As Long
  LoadDirectory Path
  If hlCount = 0 Then Exit Sub
  For i = 0 To hlCount - 1 'UBound(Highlighters) - 1
    RaiseEvent AddHighlighter(Highlighters(i).strName, Highlighters(i).strFilter)
  Next i
End Sub

Public Function DoOptions(HighlightersPath As String) As Boolean
  'Dim fOptions As New frmOptions
  Load frmOptions
  With frmOptions
    .strHoldDir = HighlightersPath
    ' Set the values to the options dialog :)
    .chkTabIndents.Value = -(CLng(TabIndents))
    .chkAutoCloseQuotes.Value = -(CLng(AutoCloseQuotes))
    .chkAutoCloseBraces.Value = -(CLng(AutoCloseBraces))
    .chkBackSpaceUnIndents.Value = -(CLng(BackSpaceUnIndents))
    .clBookBack.SelectedColor = BookmarkBack
    .clBookFore.SelectedColor = BookMarkFore
    .clMarkerBack.SelectedColor = MarkerBack
    .clMarkerFore.SelectedColor = MarkerFore
    .txtTabWidth.Text = TabWidth
    .clrCaretFore.SelectedColor = CaretForeColor
    .txtCaretWidth.Text = CaretWidth
    .clrEdgeColor.SelectedColor = EdgeColor
    .cmbEOLMode.ListIndex = EOLMode
    .chkHighlight.Value = -(CLng(HighlightBraces))
    .chkClearUndoAfterSave.Value = -(CLng(ClearUndoAfterSave))
    .chkEndLastLine.Value = -(CLng(EndAtLastLine))
    .chkMaintainIndentation.Value = -(CLng(MaintainIndentation))
    .chkOverType.Value = -(CLng(OverType))
  End With
  DoOptions = DoSyntaxOptions(HighlightersPath, Me)
  
  If DoOptions = True Then
    With frmOptions
      ' Set the values from the options dialog to the control
      TabIndents = .chkTabIndents.Value
      AutoCloseQuotes = .chkAutoCloseQuotes.Value
      AutoCloseBraces = .chkAutoCloseBraces.Value
      BackSpaceUnIndents = .chkBackSpaceUnIndents.Value
      BookmarkBack = .clBookBack.SelectedColor
      BookMarkFore = .clBookFore.SelectedColor
      MarkerBack = .clMarkerBack.SelectedColor
      MarkerFore = .clMarkerFore.SelectedColor
      TabWidth = .txtTabWidth.Text
      CaretForeColor = .clrCaretFore.SelectedColor
      CaretWidth = .txtCaretWidth.Text
      EdgeColor = .clrEdgeColor.SelectedColor
      EOLMode = .cmbEOLMode.ListIndex
      HighlightBraces = .chkHighlight.Value
      ClearUndoAfterSave = .chkClearUndoAfterSave.Value
      EndAtLastLine = .chkEndLastLine.Value
      MaintainIndentation = .chkMaintainIndentation.Value
      OverType = .chkOverType.Value
      LoadHighlighters HighlightersPath
    End With
    
  End If
  Unload frmOptions
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoCloseBraces() As Boolean
    AutoCloseBraces = m_AutoCloseBraces
End Property

Public Property Let AutoCloseBraces(ByVal New_AutoCloseBraces As Boolean)
    m_AutoCloseBraces = New_AutoCloseBraces
    PropertyChanged "AutoCloseBraces"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoCloseQuotes() As Boolean
    AutoCloseQuotes = m_AutoCloseQuotes
End Property

Public Property Let AutoCloseQuotes(ByVal New_AutoCloseQuotes As Boolean)
    m_AutoCloseQuotes = New_AutoCloseQuotes
    PropertyChanged "AutoCloseQuotes"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get TabIndents() As Boolean
    TabIndents = m_TabIndents
End Property

Public Property Let TabIndents(ByVal New_TabIndents As Boolean)
    m_TabIndents = New_TabIndents
    PropertyChanged "TabIndents"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BackSpaceUnIndents() As Boolean
    BackSpaceUnIndents = m_BackSpaceUnIndents
End Property

Public Property Let BackSpaceUnIndents(ByVal New_BackSpaceUnIndents As Boolean)
    m_BackSpaceUnIndents = New_BackSpaceUnIndents
    PropertyChanged "BackSpaceUnIndents"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblack
Public Property Get BookmarkBack() As OLE_COLOR
    BookmarkBack = m_BookmarkBack
End Property

Public Property Let BookmarkBack(ByVal New_BookmarkBack As OLE_COLOR)
    m_BookmarkBack = New_BookmarkBack
    PropertyChanged "BookmarkBack"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get BookMarkFore() As OLE_COLOR
    BookMarkFore = m_BookMarkFore
End Property

Public Property Let BookMarkFore(ByVal New_BookMarkFore As OLE_COLOR)
    m_BookMarkFore = New_BookMarkFore
    PropertyChanged "BookMarkFore"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblack
Public Property Get MarkerBack() As OLE_COLOR
    MarkerBack = m_MarkerBack
End Property

Public Property Let MarkerBack(ByVal New_MarkerBack As OLE_COLOR)
    m_MarkerBack = New_MarkerBack
    PropertyChanged "MarkerBack"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get MarkerFore() As OLE_COLOR
    MarkerFore = m_MarkerFore
End Property

Public Property Let MarkerFore(ByVal New_MarkerFore As OLE_COLOR)
    m_MarkerFore = New_MarkerFore
    PropertyChanged "MarkerFore"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,4
Public Property Get TabWidth() As Long
    TabWidth = m_TabWidth
End Property

Public Property Let TabWidth(ByVal New_TabWidth As Long)
    m_TabWidth = New_TabWidth
    PropertyChanged "TabWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblack
Public Property Get CaretForeColor() As OLE_COLOR
    CaretForeColor = m_CaretForeColor
End Property

Public Property Let CaretForeColor(ByVal New_CaretForeColor As OLE_COLOR)
    m_CaretForeColor = New_CaretForeColor
    PropertyChanged "CaretForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get CaretWidth() As Long
    CaretWidth = m_CaretWidth
End Property

Public Property Let CaretWidth(ByVal New_CaretWidth As Long)
    m_CaretWidth = New_CaretWidth
    PropertyChanged "CaretWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get EdgeColor() As OLE_COLOR
    EdgeColor = m_EdgeColor
End Property

Public Property Let EdgeColor(ByVal New_EdgeColor As OLE_COLOR)
    m_EdgeColor = New_EdgeColor
    PropertyChanged "EdgeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get EOLMode() As EOLStyle
    EOLMode = m_EOLMode
End Property

Public Property Let EOLMode(ByVal New_EOLMode As EOLStyle)
    m_EOLMode = New_EOLMode
    PropertyChanged "EOLMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get HighlightBraces() As Boolean
    HighlightBraces = m_HighlightBraces
End Property

Public Property Let HighlightBraces(ByVal New_HighlightBraces As Boolean)
    m_HighlightBraces = New_HighlightBraces
    PropertyChanged "HighlightBraces"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ClearUndoAfterSave() As Boolean
    ClearUndoAfterSave = m_ClearUndoAfterSave
End Property

Public Property Let ClearUndoAfterSave(ByVal New_ClearUndoAfterSave As Boolean)
    m_ClearUndoAfterSave = New_ClearUndoAfterSave
    PropertyChanged "ClearUndoAfterSave"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get EndAtLastLine() As Boolean
    EndAtLastLine = m_EndAtLastLine
End Property

Public Property Let EndAtLastLine(ByVal New_EndAtLastLine As Boolean)
    m_EndAtLastLine = New_EndAtLastLine
    PropertyChanged "EndAtLastLine"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get MaintainIndentation() As Boolean
    MaintainIndentation = m_MaintainIndentation
End Property

Public Property Let MaintainIndentation(ByVal New_MaintainIndentation As Boolean)
    m_MaintainIndentation = New_MaintainIndentation
    PropertyChanged "MaintainIndentation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get OverType() As Boolean
    OverType = m_OverType
End Property

Public Property Let OverType(ByVal New_OverType As Boolean)
    m_OverType = New_OverType
    PropertyChanged "OverType"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AutoCloseBraces = m_def_AutoCloseBraces
    m_AutoCloseQuotes = m_def_AutoCloseQuotes
    m_TabIndents = m_def_TabIndents
    m_BackSpaceUnIndents = m_def_BackSpaceUnIndents
    m_BookmarkBack = m_def_BookmarkBack
    m_BookMarkFore = m_def_BookMarkFore
    m_MarkerBack = m_def_MarkerBack
    m_MarkerFore = m_def_MarkerFore
    m_TabWidth = m_def_TabWidth
    m_CaretForeColor = m_def_CaretForeColor
    m_CaretWidth = m_def_CaretWidth
    m_EdgeColor = m_def_EdgeColor
    m_EOLMode = m_def_EOLMode
    m_HighlightBraces = m_def_HighlightBraces
    m_ClearUndoAfterSave = m_def_ClearUndoAfterSave
    m_EndAtLastLine = m_def_EndAtLastLine
    m_MaintainIndentation = m_def_MaintainIndentation
    m_OverType = m_def_OverType
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AutoCloseBraces = PropBag.ReadProperty("AutoCloseBraces", m_def_AutoCloseBraces)
    m_AutoCloseQuotes = PropBag.ReadProperty("AutoCloseQuotes", m_def_AutoCloseQuotes)
    m_TabIndents = PropBag.ReadProperty("TabIndents", m_def_TabIndents)
    m_BackSpaceUnIndents = PropBag.ReadProperty("BackSpaceUnIndents", m_def_BackSpaceUnIndents)
    m_BookmarkBack = PropBag.ReadProperty("BookmarkBack", m_def_BookmarkBack)
    m_BookMarkFore = PropBag.ReadProperty("BookMarkFore", m_def_BookMarkFore)
    m_MarkerBack = PropBag.ReadProperty("MarkerBack", m_def_MarkerBack)
    m_MarkerFore = PropBag.ReadProperty("MarkerFore", m_def_MarkerFore)
    m_TabWidth = PropBag.ReadProperty("TabWidth", m_def_TabWidth)
    m_CaretForeColor = PropBag.ReadProperty("CaretForeColor", m_def_CaretForeColor)
    m_CaretWidth = PropBag.ReadProperty("CaretWidth", m_def_CaretWidth)
    m_EdgeColor = PropBag.ReadProperty("EdgeColor", m_def_EdgeColor)
    m_EOLMode = PropBag.ReadProperty("EOLMode", m_def_EOLMode)
    m_HighlightBraces = PropBag.ReadProperty("HighlightBraces", m_def_HighlightBraces)
    m_ClearUndoAfterSave = PropBag.ReadProperty("ClearUndoAfterSave", m_def_ClearUndoAfterSave)
    m_EndAtLastLine = PropBag.ReadProperty("EndAtLastLine", m_def_EndAtLastLine)
    m_MaintainIndentation = PropBag.ReadProperty("MaintainIndentation", m_def_MaintainIndentation)
    m_OverType = PropBag.ReadProperty("OverType", m_def_OverType)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoCloseBraces", m_AutoCloseBraces, m_def_AutoCloseBraces)
    Call PropBag.WriteProperty("AutoCloseQuotes", m_AutoCloseQuotes, m_def_AutoCloseQuotes)
    Call PropBag.WriteProperty("TabIndents", m_TabIndents, m_def_TabIndents)
    Call PropBag.WriteProperty("BackSpaceUnIndents", m_BackSpaceUnIndents, m_def_BackSpaceUnIndents)
    Call PropBag.WriteProperty("BookmarkBack", m_BookmarkBack, m_def_BookmarkBack)
    Call PropBag.WriteProperty("BookMarkFore", m_BookMarkFore, m_def_BookMarkFore)
    Call PropBag.WriteProperty("MarkerBack", m_MarkerBack, m_def_MarkerBack)
    Call PropBag.WriteProperty("MarkerFore", m_MarkerFore, m_def_MarkerFore)
    Call PropBag.WriteProperty("TabWidth", m_TabWidth, m_def_TabWidth)
    Call PropBag.WriteProperty("CaretForeColor", m_CaretForeColor, m_def_CaretForeColor)
    Call PropBag.WriteProperty("CaretWidth", m_CaretWidth, m_def_CaretWidth)
    Call PropBag.WriteProperty("EdgeColor", m_EdgeColor, m_def_EdgeColor)
    Call PropBag.WriteProperty("EOLMode", m_EOLMode, m_def_EOLMode)
    Call PropBag.WriteProperty("HighlightBraces", m_HighlightBraces, m_def_HighlightBraces)
    Call PropBag.WriteProperty("ClearUndoAfterSave", m_ClearUndoAfterSave, m_def_ClearUndoAfterSave)
    Call PropBag.WriteProperty("EndAtLastLine", m_EndAtLastLine, m_def_EndAtLastLine)
    Call PropBag.WriteProperty("MaintainIndentation", m_MaintainIndentation, m_def_MaintainIndentation)
    Call PropBag.WriteProperty("OverType", m_OverType, m_def_OverType)
End Sub

Public Function SetStylesAndOptions(Scintilla As SCIVB, HighlighterName As String) As Boolean
  On Error GoTo errHandler
  SetStylesAndOptions = True
  With Scintilla
    ' Set the options and styles to the editor
    .TabIndents = Me.TabIndents
    .AutoCloseQuotes = Me.AutoCloseQuotes
    .AutoCloseBraces = Me.AutoCloseBraces
    .BackSpaceUnIndents = Me.BackSpaceUnIndents
    .BookmarkBack = Me.BookmarkBack
    .BookMarkFore = Me.BookMarkFore
    .MarkerBack = Me.MarkerBack
    .MarkerFore = Me.MarkerFore
    .IndentWidth = Me.TabWidth
    .CaretForeColor = Me.CaretForeColor
    .CaretWidth = Me.CaretWidth
    .EdgeColor = Me.EdgeColor
    .EOL = Me.EOLMode
    .HighlightBraces = Me.HighlightBraces
    .ClearUndoAfterSave = Me.ClearUndoAfterSave
    .EndAtLastLine = Me.EndAtLastLine
    .MaintainIndentation = Me.MaintainIndentation
    .OverType = Me.OverType
  End With
  SetHighlighter Scintilla, HighlighterName
  
  'Scintilla.SetOptions
  Exit Function
errHandler:
  SetStylesAndOptions = False
End Function

Public Function WriteSettings(RegPath As String) As Boolean
  On Error GoTo errHandler
  WriteSettings = True
  ' Writes the options to the registry
  SaveString HKEY_CLASSES_ROOT, RegPath, "TabIndents", Me.TabIndents
  SaveString HKEY_CLASSES_ROOT, RegPath, "AutoCloseQuotes", Me.AutoCloseQuotes
  SaveString HKEY_CLASSES_ROOT, RegPath, "AutoCloseBraces", Me.AutoCloseBraces
  SaveString HKEY_CLASSES_ROOT, RegPath, "BackSpaceUnIndents", Me.BackSpaceUnIndents
  SaveString HKEY_CLASSES_ROOT, RegPath, "BookMarkBack", Me.BookmarkBack
  SaveString HKEY_CLASSES_ROOT, RegPath, "BookMarkFore", Me.BookMarkFore
  SaveString HKEY_CLASSES_ROOT, RegPath, "MarkerBack", Me.MarkerBack
  SaveString HKEY_CLASSES_ROOT, RegPath, "MarkerFore", Me.MarkerFore
  SaveString HKEY_CLASSES_ROOT, RegPath, "IndentWidth", Me.TabWidth
  SaveString HKEY_CLASSES_ROOT, RegPath, "CaretForeColor", Me.CaretForeColor
  SaveString HKEY_CLASSES_ROOT, RegPath, "CaretWidth", Me.CaretWidth
  SaveString HKEY_CLASSES_ROOT, RegPath, "EdgeColor", Me.EdgeColor
  SaveString HKEY_CLASSES_ROOT, RegPath, "EOL", Me.EOLMode
  SaveString HKEY_CLASSES_ROOT, RegPath, "HighlightBraces", Me.HighlightBraces
  SaveString HKEY_CLASSES_ROOT, RegPath, "ClearUndoAfterSave", Me.ClearUndoAfterSave
  SaveString HKEY_CLASSES_ROOT, RegPath, "EndAtLastLine", Me.EndAtLastLine
  SaveString HKEY_CLASSES_ROOT, RegPath, "MaintainIndentation", Me.MaintainIndentation
  SaveString HKEY_CLASSES_ROOT, RegPath, "OverType", Me.OverType

  
  Exit Function
errHandler:
  WriteSettings = False
End Function

Public Function ReadSettings(RegPath As String) As Boolean
  On Error GoTo errHandler
  ReadSettings = True
  TabIndents = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "TabIndents", "1")
  AutoCloseQuotes = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "AutoCloseQuotes", "0")
  AutoCloseBraces = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "AutoCloseBraces", "0")
  BackSpaceUnIndents = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "BackSpaceUnIndents", "0")
  BookmarkBack = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "BookMarkBack", vbBlack)
  BookMarkFore = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "BookMarkFore", vbWhite)
  MarkerBack = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "MarkerBack", vbBlack)
  MarkerFore = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "MarkerFore", vbWhite)
  TabWidth = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "IndentWidth", "4")
  CaretForeColor = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "CaretForeColor", vbBlack)
  CaretWidth = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "CaretWidth", "1")
  EdgeColor = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "EdgeColor", &HE0E0E0)
  EOLMode = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "EOL", 0)
  HighlightBraces = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "HighlightBraces", "1")
  ClearUndoAfterSave = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "ClearUndoAfterSave", "1")
  EndAtLastLine = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "EndAtLastLine", "1")
  MaintainIndentation = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "MaintainIndentation", "1")
  OverType = ReadSetting(HKEY_CLASSES_ROOT, RegPath, "OverType", "0")
  Exit Function
errHandler:
  ReadSettings = False
End Function

Public Function ExportToHTML(FilePath As String, Scintilla As SCIVB) As Boolean
  On Error GoTo errHandler
  Call ExportToHTML2(FilePath, Scintilla)
  Exit Function
errHandler:
  ExportToHTML = False
End Function

Public Function SetHighlighterExt(Scintilla As SCIVB, strFile As String) As Boolean
  Dim str As String
  If hlCount = 0 Then Exit Function
  str = SetHighlighterBasedOnExtension(GetExtension(strFile))
  SetStylesAndOptions Scintilla, str
  RaiseEvent HighlighterSet(str)
  Scintilla.CurHigh = str
End Function

Public Sub CommentBlock(sci As SCIVB)
  CommentBlock2 sci
End Sub

Public Sub UncommentBlock(sci As SCIVB)
  UncommentBlock2 sci
End Sub

