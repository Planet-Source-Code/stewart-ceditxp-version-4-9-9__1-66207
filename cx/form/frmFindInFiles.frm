VERSION 5.00
Begin VB.Form frmFindInFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find in Files"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindInFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   ".."
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   285
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "&Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.ComboBox cmbFind 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Find in Folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Find Text:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmFindInFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFind_Click()
  Me.Hide
  frmMain.vs.FindInFiles txtFolder.Text, cmbFind.Text, chkCase.Value
End Sub

Private Sub cmdFolder_Click()
  txtFolder.Text = BrowseFolder("Find in Folder", Me)
End Sub

Private Sub Form_Load()
  LoadFormData Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
