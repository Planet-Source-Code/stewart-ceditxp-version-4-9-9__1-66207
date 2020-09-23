VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Goto"
   ClientHeight    =   1425
   ClientLeft      =   5490
   ClientTop       =   5295
   ClientWidth     =   5175
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtColumn 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLine 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblColumn 
      AutoSize        =   -1  'True
      Caption         =   "Column: "
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblLineCount 
      AutoSize        =   -1  'True
      Caption         =   "Last Line:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lblCurLine 
      AutoSize        =   -1  'True
      Caption         =   "Current Line: "
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Column:"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   165
      Width           =   570
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Destination Line:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   165
      Width           =   1185
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iWhatToDo As Long

Private Sub cmdCancel_Click()
  iWhatToDo = 0
  Me.Hide
End Sub

Private Sub cmdGo_Click()
  iWhatToDo = 1
  Me.Hide
End Sub

Private Sub Form_Load()
  Me.Left = GetSetting("ScintillaClass", "Settings", "GotoLeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaClass", "Settings", "GotoTop", (Screen.Height - Me.Height) \ 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "ScintillaClass", "Settings", "GotoLeft", Me.Left
  SaveSetting "ScintillaClass", "Settings", "GotoTop", Me.Top
End Sub

Private Sub txtColumn_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Sub

Private Sub txtLine_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
