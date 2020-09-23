VERSION 5.00
Begin VB.Form frmSaveMacro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Macro"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaveMacro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cmbSave 
      Height          =   315
      ItemData        =   "frmSaveMacro.frx":000C
      Left            =   120
      List            =   "frmSaveMacro.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Save Macro to Macro Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSaveMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DoWhat As Long

Private Sub cmdCancel_Click()
  DoWhat = 0
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  DoWhat = 1
  Me.Hide
End Sub

Private Sub Form_Load()
  cmbSave.ListIndex = 0
End Sub

