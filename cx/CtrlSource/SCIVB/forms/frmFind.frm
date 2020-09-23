VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox optMulti 
      Height          =   315
      Left            =   4725
      Picture         =   "frmFind.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtFind 
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin SCIVBX.GroupBox gbDir 
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Direction"
      Begin VB.OptionButton optDown 
         Caption         =   "&Down"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optUp 
         Caption         =   "&Up"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdMarkAll 
      Caption         =   "&Mark All"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox cmbFind 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkRegExp 
      Caption         =   "Regular &expression"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1500
      Width           =   1695
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &case"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox chkWhole 
      Caption         =   "Match &whole word only"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   2535
   End
   Begin VB.CheckBox chkWrap 
      Caption         =   "Wrap aroun&d"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bFind As Boolean
Public DoWhat As Integer
Public bMulti As Boolean
Private Const NormalHeight = 2250

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFind_Click()
  DoWhat = 1
  Me.Hide
  AddCombo cmbFind, cmbFind.Text
End Sub

Private Sub cmdMarkAll_Click()
  DoWhat = 2
  Me.Hide
End Sub

Private Sub Form_Activate()
  bMulti = GetSetting("ScintillaClass", "Settings", "optMulti", 0)
  If bMulti = True Then
    optMulti.Value = 1
  Else
    optMulti.Value = 0
  End If
End Sub

Private Sub Form_Load()
  DoWhat = 0
  
  FlatBorder txtFind.hwnd
  FlatBorder optMulti.hwnd
  'Flatten Me
  Me.Left = GetSetting("ScintillaClass", "Settings", "FindLeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaClass", "Settings", "FindTop", (Screen.Height - Me.Height) \ 2)
  chkCase.Value = GetSetting("ScintillaClass", "Settings", "FchkCase", 0)
  chkRegExp.Value = GetSetting("ScintillaClass", "Settings", "FchkRegEx", 0)
  chkWhole.Value = GetSetting("ScintillaClass", "Settings", "FchkWhole", 0)
  chkWrap.Value = GetSetting("ScintillaClass", "Settings", "FchkWrap", 1)
  optUp.Value = GetSetting("ScintillaClass", "Settings", "FOptUp", 0)
  ComboLoadHistory cmbFind
  If cmbFind.ListCount > 0 Then
    cmbFind.Text = cmbFind.List(0)
  End If
  
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "ScintillaClass", "Settings", "FindLeft", Me.Left
  SaveSetting "ScintillaClass", "Settings", "FindTop", Me.Top
  SaveSetting "ScintillaClass", "Settings", "FchkCase", chkCase.Value
  SaveSetting "ScintillaClass", "Settings", "FchkRegEx", chkRegExp.Value
  SaveSetting "ScintillaClass", "Settings", "FchkWhole", chkWhole.Value
  SaveSetting "ScintillaClass", "Settings", "FchkWrap", chkWrap.Value
  SaveSetting "ScintillaClass", "Settings", "FOptUp", optUp.Value
  SaveSetting "ScintillaClass", "Settings", "OptMulti", optMulti.Value
End Sub

Private Sub optMulti_Click()
  On Error Resume Next
  If optMulti.Value = 1 Then
    Me.Height = Me.Height + txtFind.Height - cmbFind.Height
    chkWrap.Top = chkWrap.Top + txtFind.Height - cmbFind.Height
    chkWhole.Top = chkWhole.Top + txtFind.Height - cmbFind.Height
    chkCase.Top = chkCase.Top + txtFind.Height - cmbFind.Height
    chkRegExp.Top = chkRegExp.Top + txtFind.Height - cmbFind.Height
    gbDir.Top = gbDir.Top + txtFind.Height - cmbFind.Height
    txtFind.visible = True
    txtFind.SetFocus
    bMulti = True
  Else
    Me.Height = Me.Height - (txtFind.Height - cmbFind.Height)
    chkWrap.Top = chkWrap.Top - (txtFind.Height - cmbFind.Height)
    chkWhole.Top = chkWhole.Top - (txtFind.Height - cmbFind.Height)
    chkCase.Top = chkCase.Top - (txtFind.Height - cmbFind.Height)
    chkRegExp.Top = chkRegExp.Top - (txtFind.Height - cmbFind.Height)
    gbDir.Top = gbDir.Top - (txtFind.Height - cmbFind.Height)
    txtFind.visible = False
    cmbFind.SetFocus
    bMulti = False
  End If
End Sub


