VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encryption"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objHost As Object              'Contains the parent object


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDecrypt_Click()
  On Error Resume Next
  Dim txt As String
  txt = objHost.ActiveForm.sciMain.Text
  txt = Base64Decode(txt)
  objHost.ActiveForm.sciMain.Text = txt
End Sub

Private Sub cmdEncrypt_Click()
  On Error Resume Next
  Dim txt As String
  txt = objHost.ActiveForm.sciMain.Text
  If objHost.ActiveForm.sciMain.eol > 0 Then
    If objHost.ActiveForm.sciMain.eol = 1 Then
      txt = Replace(txt, vbCr, vbCrLf)
    End If
    If objHost.ActiveForm.sciMain.eol = 2 Then
      txt = Replace(txt, vbLf, vbCrLf)
    End If
  End If
  txt = Base64Encode(txt)
  objHost.ActiveForm.sciMain.Text = txt
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set objHost = Nothing
End Sub
