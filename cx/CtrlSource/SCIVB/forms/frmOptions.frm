VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Syntax Settings"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SCIVBX.ctlTabs tbMain 
      Height          =   5775
      Left            =   2760
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10186
      TABWIDE         =   70
      TABHIGH         =   20
      TABCOUNT        =   2
      TABSELECTED     =   1
      TABSTYLE        =   0
      CAPTIONSTYLE    =   4
      FOCUSRECT       =   0   'False
      TABCAPTION1     =   "Styles"
      TABCAPTION2     =   "Keywords"
      BeginProperty TABFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TABFONTACTIVE {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TABCOLOR        =   -2147483632
      TABCOLORACTIVE  =   -2147483633
      TEXTCOLOR       =   -2147483628
      TEXTCOLORACTIVE =   -2147483630
      Begin VB.PictureBox picStyles 
         BorderStyle     =   0  'None
         Height          =   5355
         Left            =   120
         ScaleHeight     =   5355
         ScaleWidth      =   5175
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   360
         Width           =   5175
         Begin SCIVBX.GroupBox gbGeneral 
            Height          =   975
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   1720
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "General"
            Begin VB.TextBox txtFilter 
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox txtComment 
               Height          =   315
               Left            =   2640
               TabIndex        =   47
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label1 
               Caption         =   "Filter:"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Single Line Comment:"
               Height          =   255
               Left            =   2640
               TabIndex        =   49
               Top             =   240
               Width           =   1575
            End
         End
         Begin SCIVBX.GroupBox gbSettings 
            Height          =   1875
            Left            =   0
            TabIndex        =   51
            Top             =   3360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3307
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Settings"
            Begin VB.ComboBox txtSize 
               Height          =   315
               ItemData        =   "frmOptions.frx":000C
               Left            =   2520
               List            =   "frmOptions.frx":0051
               TabIndex        =   58
               Top             =   480
               Width           =   2295
            End
            Begin VB.ComboBox cmbFont 
               Height          =   315
               ItemData        =   "frmOptions.frx":0096
               Left            =   120
               List            =   "frmOptions.frx":0098
               TabIndex        =   57
               Text            =   "cmbFont"
               Top             =   480
               Width           =   2295
            End
            Begin VB.CheckBox chkBold 
               Caption         =   "&Bold"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   1485
               Width           =   735
            End
            Begin VB.CheckBox chkItalic 
               Caption         =   "&Italic"
               Height          =   255
               Left            =   960
               TabIndex        =   55
               Top             =   1485
               Width           =   735
            End
            Begin VB.CheckBox chkEOL 
               Caption         =   "&EOL"
               Height          =   255
               Left            =   1800
               TabIndex        =   54
               Top             =   1485
               Width           =   735
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "&Underline"
               Height          =   195
               Left            =   3600
               TabIndex        =   53
               Top             =   1485
               Width           =   975
            End
            Begin VB.CheckBox chkVisible 
               Caption         =   "&Visible"
               Height          =   195
               Left            =   2640
               TabIndex        =   52
               Top             =   1485
               Width           =   855
            End
            Begin SCIVBX.ArielColorBox clrBack 
               Height          =   315
               Left            =   2520
               TabIndex        =   59
               Top             =   1080
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   2
            End
            Begin SCIVBX.ArielColorBox clrFore 
               Height          =   315
               Left            =   120
               TabIndex        =   60
               Top             =   1080
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   2
            End
            Begin VB.Label Label9 
               Caption         =   "Forecolor:"
               Height          =   375
               Left            =   120
               TabIndex        =   64
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label10 
               Caption         =   "Backcolor:"
               Height          =   255
               Left            =   2520
               TabIndex        =   63
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label Label4 
               Caption         =   "Font:"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "Size (0=Default):"
               Height          =   255
               Left            =   2520
               TabIndex        =   61
               Top             =   240
               Width           =   1455
            End
         End
         Begin SCIVBX.GroupBox gbStyle 
            Height          =   2175
            Left            =   0
            TabIndex        =   65
            Top             =   1080
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3836
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Style"
            Begin VB.TextBox txtStyleDesc 
               Height          =   315
               Left            =   120
               TabIndex        =   68
               Top             =   1680
               Width           =   3885
            End
            Begin VB.ListBox lstStyle 
               Height          =   1035
               ItemData        =   "frmOptions.frx":009A
               Left            =   120
               List            =   "frmOptions.frx":009C
               TabIndex        =   67
               Top             =   240
               Width           =   4935
            End
            Begin VB.CommandButton cmdAddStyle 
               Caption         =   "&Add Style"
               Height          =   315
               Left            =   4080
               TabIndex        =   66
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Style Description:"
               Height          =   195
               Left            =   120
               TabIndex        =   69
               Top             =   1440
               Width           =   1230
            End
         End
      End
      Begin VB.PictureBox picKeywords 
         BorderStyle     =   0  'None
         Height          =   5340
         Left            =   120
         ScaleHeight     =   5340
         ScaleWidth      =   5175
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtKeyword 
            Height          =   4215
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   1080
            Width           =   5175
         End
         Begin VB.ComboBox cmbKeyword 
            Height          =   315
            ItemData        =   "frmOptions.frx":009E
            Left            =   0
            List            =   "frmOptions.frx":00BA
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   360
            Width           =   5175
         End
         Begin VB.Label Label8 
            Caption         =   "Keywords:"
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label7 
            Caption         =   "Keyword Sets:"
            Height          =   375
            Left            =   0
            TabIndex        =   43
            Top             =   120
            Width           =   2535
         End
      End
   End
   Begin SCIVBX.ctlFrame ctlFrame2 
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10186
      Begin SCIVBX.ucTreeView tvMain 
         Height          =   4935
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   8281
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Language"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2295
      End
   End
   Begin SCIVBX.ctlFrame ctlFrame1 
      Height          =   5775
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10186
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5535
         ScaleWidth      =   5175
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   5175
         Begin SCIVBX.GroupBox gbMarkerColors 
            Height          =   1815
            Left            =   2520
            TabIndex        =   7
            Top             =   1080
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   3201
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Marker Colors"
            Begin SCIVBX.ArielColorBox clMarkerBack 
               Height          =   315
               Left            =   120
               TabIndex        =   8
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   4
            End
            Begin SCIVBX.ArielColorBox clMarkerFore 
               Height          =   315
               Left            =   120
               TabIndex        =   9
               Top             =   1200
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   4
            End
            Begin VB.Label lbMarkerBack 
               Caption         =   "Marker Back:"
               Height          =   375
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label lbMarkerFore 
               Caption         =   "Marker Fore:"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   1695
            End
         End
         Begin SCIVBX.GroupBox gbBookmarkColors 
            Height          =   1815
            Left            =   0
            TabIndex        =   12
            Top             =   1080
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   3201
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Bookmark Colors"
            Begin SCIVBX.ArielColorBox clBookBack 
               Height          =   315
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   4
            End
            Begin SCIVBX.ArielColorBox clBookFore 
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   1200
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   4
            End
            Begin VB.Label lbBookMarkBack 
               Caption         =   "Bookmark Back:"
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label lbBookMarkFore 
               Caption         =   "Bookmark Fore:"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   960
               Width           =   1695
            End
         End
         Begin SCIVBX.GroupBox gbIndentOptions 
            Height          =   975
            Left            =   2520
            TabIndex        =   17
            Top             =   0
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1720
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Indent Options"
            Begin VB.CheckBox chkTabIndents 
               Caption         =   "&Tab Indents"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   1575
            End
            Begin VB.CheckBox chkBackSpaceUnIndents 
               Caption         =   "&Backspace Unindents"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   1935
            End
         End
         Begin SCIVBX.GroupBox gbAutoClose 
            Height          =   975
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   1720
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Auto Close Options"
            Begin VB.CheckBox chkAutoCloseBraces 
               Caption         =   "&Auto Close Braces"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   1935
            End
            Begin VB.CheckBox chkAutoCloseQuotes 
               Caption         =   "Auto Close &Quotes"
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   600
               Width           =   1815
            End
         End
         Begin SCIVBX.GroupBox gbGeneralOptions 
            Height          =   2535
            Left            =   0
            TabIndex        =   23
            Top             =   3000
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4471
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "General Options"
            Begin VB.CheckBox chkOverType 
               Caption         =   "&Overtype"
               Height          =   255
               Left            =   2880
               TabIndex        =   31
               Top             =   1800
               Width           =   1935
            End
            Begin VB.CheckBox chkMaintainIndentation 
               Caption         =   "&Maintain Indentation"
               Height          =   255
               Left            =   2880
               TabIndex        =   30
               Top             =   1440
               Width           =   1815
            End
            Begin VB.CheckBox chkEndLastLine 
               Caption         =   "&End at Last Line"
               Height          =   255
               Left            =   2880
               TabIndex        =   29
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CheckBox chkClearUndoAfterSave 
               Caption         =   "&Clear Undo After Save"
               Height          =   255
               Left            =   2880
               TabIndex        =   28
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox chkHighlight 
               Caption         =   "&Highlight Braces"
               Height          =   255
               Left            =   2880
               TabIndex        =   27
               Top             =   360
               Width           =   1455
            End
            Begin VB.ComboBox cmbEOLMode 
               Height          =   315
               ItemData        =   "frmOptions.frx":00D6
               Left            =   1560
               List            =   "frmOptions.frx":00E3
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtCaretWidth 
               Height          =   285
               Left            =   1560
               TabIndex        =   25
               Text            =   "1"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtTabWidth 
               Height          =   285
               Left            =   1560
               TabIndex        =   24
               Text            =   "4"
               Top             =   360
               Width           =   1095
            End
            Begin SCIVBX.ArielColorBox clrEdgeColor 
               Height          =   315
               Left            =   1560
               TabIndex        =   32
               Top             =   1440
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   4
            End
            Begin SCIVBX.ArielColorBox clrCaretFore 
               Height          =   315
               Left            =   1560
               TabIndex        =   33
               Top             =   720
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Palette         =   4
            End
            Begin VB.Label lbEOLMode 
               Caption         =   "EOL Mode:"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   1830
               Width           =   1215
            End
            Begin VB.Label lbEdgeColor 
               Caption         =   "Edge Color:"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   1470
               Width           =   1215
            End
            Begin VB.Label lbCaretWidth 
               Caption         =   "Caret Width:"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   1095
               Width           =   1215
            End
            Begin VB.Label lbCaretFore 
               Caption         =   "Caret Forecolor:"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label lbTabWidth 
               Caption         =   "Tab Width:"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   375
               Width           =   975
            End
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hLighter() As Highlighter
Private Modified() As Boolean

Dim lStyle As Long
Public hlPath As String
Public hlMain As SCIHighlighter
Public WhatToDo As Long
Public strHoldDir
Dim lSelLang As Long
Dim Lexer() As String

Private Function GetUpper2(varArray() As Highlighter) As Long
Dim Upper As Integer
On Error Resume Next
Upper = UBound(varArray)
If Err.Number Then
     If Err.Number = 9 Then
          Upper = 0
     Else
          With Err
               MsgBox "Error:" & .Number & "-" & .Description
          End With
          Exit Function
     End If
Else
     Upper = UBound(varArray) + 1
End If
On Error GoTo 0
GetUpper2 = Upper
End Function


Private Sub clrBack_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleBack(lStyle) = clrBack.SelectedColor
  Modified(lSelLang) = True
End Sub

Private Sub clrFore_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleFore(lStyle) = clrFore.SelectedColor
  Modified(lSelLang) = True
End Sub

Private Sub cmbFont_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleFont(lStyle) = cmbFont.Text
  Modified(lSelLang) = True
End Sub

Private Sub cmbKeyword_Click()
  On Error Resume Next
  txtKeyword.Text = hLighter(lSelLang).Keywords(cmbKeyword.ListIndex)
End Sub

Private Sub cmdAdd_Click()
  Load frmNewLang
  frmNewLang.strDir = strHoldDir
  frmNewLang.show vbModal, Me
End Sub

Private Sub cmdAddStyle_Click()
  Dim msgRes As VbMsgBoxResult
  With frmAddStyle
    .show vbModal, Me
    If .HitOK Then
      ' The user hit ok.
      If GetUpper2(hLighter) >= lSelLang And lSelLang > -1 Then
        'Make sure we are dealing with an existing highlighter here.
        If hLighter(lSelLang).StyleName(.txtStyle.Text) <> "" Then
          ' We already have a style here.
          msgRes = MsgBox("The style you have entered already exists." & vbCrLf & "Do you wish to overwrite it's description?", vbYesNo + vbQuestion, "Overwrite")
          'Let's ask if they want to overwrite that style
          If msgRes = vbYes Then
            ' They said yes so do it
            hLighter(lSelLang).StyleName(.txtStyle.Text) = .txtDesc.Text
          Else
            ' They said no so ignore it and exit sub.
            Unload frmAddStyle
            Exit Sub
          End If
        End If
        ' None of the pitfal errors have occurred so let's just do it
        hLighter(lSelLang).StyleName(.txtStyle.Text) = .txtDesc.Text
        Unload frmAddStyle
        ' Redisplay the options with the newly added style included :)
        DispOpt True
      End If
    End If
  End With
End Sub

Private Sub cmdCancel_Click()
  WhatToDo = 0
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  Dim i As Long
  WriteSettings
  WhatToDo = 1
  Me.Hide
End Sub

Private Sub Form_Load()
  On Error Resume Next
  txtSize.ListIndex = 0
  With tvMain
    .Initialize
    .InitializeImageList
    Call .AddBitmap(LoadResPicture(102, vbResBitmap)) 'Folder Open
    Call .AddBitmap(LoadResPicture(103, vbResBitmap)) 'Page
    Call .AddBitmap(LoadResPicture(104, vbResBitmap)) 'Folder
  End With
  InitTreeView
  LoadFontList
  Me.Left = GetSetting("ScintillaClass", "Settings", "OptLeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaClass", "Settings", "OptTop", (Screen.Height - Me.Height) \ 2)
  tbMain_TabClick 1, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Erase Lexer()
  SaveSetting "ScintillaClass", "Settings", "OptLeft", Me.Left
  SaveSetting "ScintillaClass", "Settings", "OptTop", Me.Top
End Sub

Private Sub lstStyle_Click()
  lStyle = lstStyle.ItemData(lstStyle.ListIndex)
  Me.Caption = lStyle
  DispOpt
End Sub

Private Sub tbMain_TabClick(OldTab As Integer, NewTab As Integer)
  picStyles.visible = False
  picKeywords.visible = False
  Select Case NewTab
    Case 1
      picStyles.visible = True
    Case 2
      picKeywords.visible = True
  End Select
End Sub

Private Sub tvMain_NodeClick(ByVal hNode As Long)
  lSelLang = -1
  If Left(tvMain.GetNodeKey(hNode), 3) = "syn" Then
    picOptions.visible = False
    'picKeywords.visible = False
    lStyle = 0
    'picStyles.visible = True
    tbMain.visible = True
    lSelLang = Mid(tvMain.GetNodeKey(hNode), 4)
    DispOpt True
  End If
'  If Left(tvMain.GetNodeKey(hNode), 3) = "key" Then
'    picOptions.visible = False
'    lStyle = 0
'    picKeywords.visible = True
'    picStyles.visible = False
'    lSelLang = Mid(tvMain.GetNodeKey(hNode), 4)
'    DispOpt True
'  End If
  If Left(tvMain.GetNodeKey(hNode), 3) = "gen" Then
    picOptions.visible = True
    lStyle = 0
    tbMain.visible = False
    lSelLang = -1
  End If
'  DispOpt True
End Sub

Private Sub txtComment_Change()
  On Error Resume Next
  hLighter(lSelLang).strComment = txtComment.Text
  Modified(lSelLang) = True
End Sub

Private Sub txtFilter_Change()
  On Error Resume Next
  hLighter(lSelLang).strFilter = txtFilter.Text
  Modified(lSelLang) = True
End Sub

Private Sub txtKeyword_Change()
  On Error Resume Next
  hLighter(lSelLang).Keywords(cmbKeyword.ListIndex) = txtKeyword.Text
End Sub

Private Sub txtSize_Change()
  On Error Resume Next
  hLighter(lSelLang).StyleSize(lStyle) = txtSize.Text
  Modified(lSelLang) = True
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
  KeyAscii = IsNumericKey(KeyAscii)
End Sub


Private Sub chkBold_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleBold(lStyle) = chkBold.Value
  Modified(lSelLang) = True
End Sub

Private Sub chkEOL_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleEOLFilled(lStyle) = chkEOL.Value
  Modified(lSelLang) = True
End Sub

Private Sub chkItalic_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleItalic(lStyle) = chkItalic.Value
  Modified(lSelLang) = True
End Sub

Private Sub chkUnderline_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleUnderline(lStyle) = chkUnderline.Value
  Modified(lSelLang) = True
End Sub

Private Sub chkVisible_Click()
  On Error Resume Next
  hLighter(lSelLang).StyleVisible(lStyle) = chkVisible.Value
  Modified(lSelLang) = True
End Sub

Private Sub DispOpt(Optional ListStyles As Boolean = False)
  'On Error GoTo errHandler
  Dim lLexNum As Long, i As Long
  ' This is a basic function that will just set the options
  ' to the different options (checkboxes, textboxes, etc.),
  ' based on the highlighter selected.
  'If lstStyle.ListIndex = -1 Then Exit Sub
  'If lstStyle.ListIndex = -1 Then Exit Sub
  'lStyle = lstStyle.ItemData(lstStyle.ListIndex)
  lLexNum = hLighter(lSelLang).iLang
  If lLexNum < 0 Then Exit Sub  ' If we have nothing for this lexer lets
                                ' Kill this before it causes errors!
  Lexer() = Split(LexList(lLexNum), ":")
  txtFilter.Text = hLighter(lSelLang).strFilter
  cmbFont.Text = hLighter(lSelLang).StyleFont(lStyle)
  clrFore.SelectedColor = hLighter(lSelLang).StyleFore(lStyle)
  clrBack.SelectedColor = hLighter(lSelLang).StyleBack(lStyle)
  If lLexNum > 75 Then Exit Sub
  If hLighter(lSelLang).StyleName(lStyle) <> "" Then
    txtStyleDesc.Text = hLighter(lSelLang).StyleName(lStyle)
  Else
    If Lexer(lStyle) <> "" Then
      txtStyleDesc.Text = Lexer(lStyle)
    End If
  End If
    
  txtStyleDesc.Text = IIf(hLighter(lSelLang).StyleName(lStyle) <> "", hLighter(lSelLang).StyleName(lStyle), IIf(Lexer(lStyle) <> "", Lexer(lStyle), ""))
  txtComment.Text = hLighter(lSelLang).strComment
  txtSize.Text = hLighter(lSelLang).StyleSize(lStyle)
  chkBold.Value = hLighter(lSelLang).StyleBold(lStyle)
  chkEOL.Value = hLighter(lSelLang).StyleEOLFilled(lStyle)
  chkItalic.Value = hLighter(lSelLang).StyleItalic(lStyle)
  chkUnderline.Value = hLighter(lSelLang).StyleUnderline(lStyle)
  chkVisible.Value = hLighter(lSelLang).StyleVisible(lStyle)
  cmbKeyword.ListIndex = 0
  txtKeyword.Text = hLighter(lSelLang).Keywords(0)
  If lSelLang > -1 And ListStyles = True Then
    lstStyle.Clear
    
    If GetUpper(Lexer) > 0 Then
      
      For i = 0 To UBound(Lexer()) - 1
        'If LCase(hLighter(lSelLang).StyleName(i)) = "defau" Or LCase(hLighter(lSelLang).StyleName(i)) = "not set" Or LCase(hLighter(lSelLang).StyleName(i)) = "default" Or LCase(hLighter(lSelLang).StyleName(i)) = "defaul" Or LCase(hLighter(lSelLang).StyleName(i)) = "none" Then
        '  hLighter(lSelLang).StyleName(i) = ""
        'End If
        If UBound(hLighter(lSelLang).StyleName) > 0 And hLighter(lSelLang).StyleName(i) <> "" Then
          lstStyle.AddItem hLighter(lSelLang).StyleName(i)
          lstStyle.ItemData(lstStyle.ListCount - 1) = i
        ElseIf UBound(Lexer) >= i And Lexer(i) <> "" Then
          lstStyle.AddItem Lexer(i)
          lstStyle.ItemData(lstStyle.ListCount - 1) = i
        End If
      Next i
      If lstStyle.ListCount > 0 Then lstStyle.ListIndex = 0
    End If
  End If
errHandler:
  Exit Sub
End Sub

Private Sub WriteSettings()
  On Error Resume Next
  Dim i As Long, X As Long
  Dim strFile As String
  Dim strOutput As String
  If GetUpper2(hLighter) > 0 Then
    For i = 0 To UBound(hLighter) - 1
      If Modified(i) = True Then
        ' Save a little time here.  Basicly if no modification then
        ' Don't write it :)  With the default highlighters this shaves
        ' approximatly 107MS off the close time on the dialog.
        Open Left(hLighter(i).strFile, Len(hLighter(i).strFile) - 3) & "bin" For Binary Access Write As #1
          hLighter(i).strFile = Left(hLighter(i).strFile, Len(hLighter(i).strFile) - 3) & "bin"
          Put #1, , hLighter(i)
        Close #1
      End If
    Next i
  End If
End Sub

Private Sub txtStyleDesc_Change()
  On Error Resume Next
  hLighter(lSelLang).StyleName(lStyle) = txtStyleDesc.Text
  lstStyle.List(lstStyle.ListIndex) = txtStyleDesc.Text
  Modified(lSelLang) = True
End Sub

Public Sub InitTreeView()
  On Error Resume Next
  Dim pNode As Long, pMain As Long
  Dim i As Long
  With tvMain
    .Clear
    .HideSelection = False
    .HasRootLines = True
    .TrackSelect = True
    .HasButtons = True
    .HasLines = True
     pMain = .AddNode(, , "Main", "Settings", 2, 0, True)
    .AddNode pMain, , "gen", "General Options", 1, 1
    pNode = .AddNode(pMain, , "Syntax", "Syntax", 2, 0)
    ReDim hLighter(0 To hlCount)
    ReDim Modified(0 To hlCount)
    If hlCount > 0 Then
      For i = 0 To UBound(Highlighters) - 1
        hLighter(i) = Highlighters(i)
        Modified(i) = False
        .AddNode pNode, rLast, "syn" & i, hLighter(i).strName, 1, 1
      Next i
    End If
    .Expand pNode
'    pNode = .AddNode(pMain, , "Words", "Keywords    ", 2, 0)
'    If hlCount > 0 Then
'      For i = 0 To UBound(Highlighters) - 1
'        hLighter(i) = Highlighters(i)
'        .AddNode pNode, rLast, "key" & i, hLighter(i).strName, 1, 1
'      Next i
'    End If
    .Expand pMain
  End With
End Sub

Private Sub LoadFontList()
  On Error Resume Next
  GetFontList cmbFont
End Sub
