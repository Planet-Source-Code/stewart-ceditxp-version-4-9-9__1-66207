VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3090
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Char/Word Data"
      Height          =   1245
      Left            =   135
      TabIndex        =   7
      Top             =   840
      Width           =   4065
      Begin VB.Label lblChar 
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   3300
      End
      Begin VB.Label lblWord 
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   585
         Width           =   3300
      End
      Begin VB.Label lblLine 
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   885
         Width           =   3300
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "File/Size Data"
      Height          =   1245
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4065
      Begin VB.Label lblSizeB 
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   885
         Width           =   3300
      End
      Begin VB.Label lblSizeK 
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   585
         Width           =   3300
      End
      Begin VB.Label lblFile 
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   3300
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   0
      Width           =   4305
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Caption         =   "Document: Untitled1"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "frmProperties.frx":1042
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Caption         =   "Document Properties"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  LoadFormData Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
