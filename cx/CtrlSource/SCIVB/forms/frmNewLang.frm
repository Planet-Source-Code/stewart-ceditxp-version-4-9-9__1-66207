VERSION 5.00
Begin VB.Form frmNewLang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Language"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmNewLang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SCIVBX.ArielColorBox clrBack 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
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
      SelectedColor   =   16777215
   End
   Begin SCIVBX.ArielColorBox clrFore 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
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
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox picSplit 
      Height          =   60
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   2835
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "10"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ComboBox cmbLexer 
      Height          =   315
      ItemData        =   "frmNewLang.frx":000C
      Left            =   120
      List            =   "frmNewLang.frx":00F4
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Language Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Def Backcolor:"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Def Forecolor:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Default Font Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Default Font:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblLexer 
      Caption         =   "Use Lexer:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmNewLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public strDir As String

Private Sub cmbLexer_Change()
  ComboAutoComplete cmbLexer
  txtName.Text = cmbLexer.Text
End Sub

Private Sub cmbLexer_Click()
  txtName.Text = cmbLexer.Text
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim strFile As String
  Dim strHold As String
  Dim lLang As Integer
  Dim strLang As String
  Dim msgRes As VbMsgBoxResult
  Dim i As Long
  If txtName.Text = "" Then
    MsgBox "Please enter a valid name for this language"
    Exit Sub
  End If
  lLang = cmbLexer.ListIndex
  If lLang < 0 Then
    For i = 0 To cmbLexer.ListCount - 1
      If LCase(cmbLexer.List(i)) = LCase(cmbLexer.Text) Then
        lLang = i
        Exit For
      End If
    Next
  End If
  If lLang < 0 Or lLang > 75 Then
    ' The user has entered a non existant name for a language.
    ' Let's ask them what to do.
    msgRes = MsgBox("You have entered an invalid lexer." & vbCrLf & "Do you wish to correct this or cancel?", vbOKCancel + vbQuestion, "Error")
    If msgRes = vbCancel Then
      Unload Me  ' The user has chosen to cancel the dialog so exit
    Else
      Exit Sub   ' No need to do anything than exit the sub and go back.
    End If
  End If
  'If lLang > 28 Then lLang = lLang + 2
  strLang = lLang
  Dim hTemp As Highlighter
  For i = 0 To 127
    hTemp.StyleBack(i) = clrBack.SelectedColor
    hTemp.StyleFore(i) = clrFore.SelectedColor
    hTemp.StyleFont(i) = cmbFont.Text
    hTemp.StyleVisible(i) = 1
    hTemp.StyleSize(i) = txtSize.Text
  Next i
  hTemp.iLang = lLang
  hTemp.strName = txtName.Text
  strFile = strDir & "\" & txtName.Text & ".bin"
  Open strFile For Binary Access Write As #1
    hTemp.strFile = strFile
    Put #1, , hTemp
  Close #1
  LoadHighlighter strFile
  frmOptions.InitTreeView
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Long
  For i = 0 To Screen.FontCount - 1
    cmbFont.AddItem Screen.Fonts(i)
  Next i
  clrBack.SelectedColor = vbWhite
  cmbFont.Text = "Courier New"
  cmbLexer.ListIndex = 3
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Sub
