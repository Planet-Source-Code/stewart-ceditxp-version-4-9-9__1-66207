VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About cEditMX"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   53
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   3480
      Width           =   6855
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmAbout.frx":08CA
      Top             =   720
      Width           =   5775
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   120
      Picture         =   "frmAbout.frx":0B01
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
   Begin VB.Label lblTitle 
      Caption         =   "cEditXP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

