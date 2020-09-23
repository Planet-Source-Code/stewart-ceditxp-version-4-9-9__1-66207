VERSION 5.00
Begin VB.Form frmChild 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmChild.frx":0000
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
