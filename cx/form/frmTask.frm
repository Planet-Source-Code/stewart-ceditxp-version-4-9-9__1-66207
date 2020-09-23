VERSION 5.00
Begin VB.Form frmTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Task Editor"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   53
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   4920
      Width           =   4575
   End
   Begin VB.TextBox txtDesc 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1920
      Width           =   4575
   End
   Begin VB.HScrollBar hPer 
      Height          =   375
      LargeChange     =   5
      Left            =   120
      Max             =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblPer 
      Caption         =   "Percentage Completed:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Task Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public iItemNum As Integer
Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
  If bEdit = False Then
    With frmMain.lstTask.ListItems.Add
      .SubItems(1) = txtName.Text
      .SubItems(2) = lblNum.Caption
      .SubItems(3) = txtDesc.Text
    End With
  Else
    With frmMain.lstTask.ListItems(iItemNum)
      .SubItems(1) = txtName.Text
      .SubItems(2) = lblNum.Caption
      .SubItems(3) = txtDesc.Text
    End With
  End If
  Me.Hide
End Sub

Private Sub Form_Load()
  LoadFormData Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub

Private Sub hPer_Change()
  lblNum = hPer.Value * 10 & "%"
End Sub

Private Sub hPer_Scroll()
  lblNum = hPer.Value * 10 & "%"
End Sub
