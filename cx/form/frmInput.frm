VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Input"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   1275
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      TabIndex        =   2
      Top             =   1275
      Width           =   990
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   4710
   End
   Begin VB.Image picIcon 
      Height          =   480
      Left            =   135
      Picture         =   "frmInput.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4635
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
  Result = ""
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Result = txtInput.Text
  Unload Me
End Sub

Private Sub Form_Load()
  LoadFormData Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
