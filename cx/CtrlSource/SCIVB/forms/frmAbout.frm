VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SCIVB"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000B&
      Height          =   1215
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":0000
      Top             =   720
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   3955
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   3955
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Label lblUpdates 
      AutoSize        =   -1  'True
      Caption         =   "For Details and Updates:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label lblURL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "http://scivb.sourceforge.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      MouseIcon       =   "frmAbout.frx":0006
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2280
      Width           =   2820
   End
   Begin VB.Label lblTop 
      Caption         =   "SCIVB"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0310
      Top             =   120
      Width           =   480
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

Private Sub Form_Load()
  Dim str As String
  str = "SCIVB Version: " & App.Major & "." & App.Minor & "." & App.Revision
  str = str & vbCrLf & vbCrLf & "SCIVB is an easy to use wrapper VB ActiveX control for Scintilla, an excellent opensource component available at http://www.scintilla.org which provides high quality syntax highlighting, code folding, code tips, code hints, and more."
  lblURL.ForeColor = vbBlue
  lblURL.Font.bold = True
  txtDesc.Text = str
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblURL.ForeColor = vbBlue
End Sub

Private Sub lblURL_Click()
  ShellDocument "http://scivb.sourceforge.net"
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (X > 0) And (X < lblURL.Width) And (Y > 0) And (Y < lblURL.Height) Then
    lblURL.ForeColor = vbRed
  Else
    lblURL.ForeColor = vbBlue
  End If
End Sub
