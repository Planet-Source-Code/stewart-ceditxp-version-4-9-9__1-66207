VERSION 5.00
Object = "*\A..\ScintillaWrapper.vbp"
Begin VB.Form frmHTML 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ScintillaWrapper.usrScintilla txtHTML 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   1693
      _ExtentY        =   1693
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  txtHTML.InitScintilla
End Sub

Private Sub Form_Resize()
  txtHTML.MoveIt 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

