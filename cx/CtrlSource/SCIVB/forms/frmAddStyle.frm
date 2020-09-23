VERSION 5.00
Begin VB.Form frmAddStyle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add A Style"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   Icon            =   "frmAddStyle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin SCIVBX.GroupBox gbStyle 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "Style Information "
      Begin VB.TextBox txtDesc 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtStyle 
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Style Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Style Num:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   465
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmAddStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HitOK As Boolean

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  If IsNumeric(txtStyle) = False Then
    MsgBox "You must enter a numeric value for the style." & vbCrLf & "The value must be between 0 and 127.", vbOKOnly + vbExclamation, "Alert"
    Exit Sub
  End If
  If txtStyle.Text > 127 Or txtStyle.Text < 0 Then
    MsgBox "You must enter a numeric value for the style." & vbCrLf & "The value must be between 0 and 127.", vbOKOnly + vbExclamation, "Alert"
    Exit Sub
  End If
  HitOK = True
  Me.Hide
End Sub

Private Sub Form_Load()
  HitOK = False
End Sub

Private Sub txtStyle_KeyPress(KeyAscii As Integer)
  KeyAscii = IsNumericKey(KeyAscii)
End Sub
