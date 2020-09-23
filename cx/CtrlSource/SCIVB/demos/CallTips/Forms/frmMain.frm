VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.0#0"; "SCIVBX.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin SCIVBX.SCIHighlighter hlMain 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SCIVBX.SCIVB sciMain 
      Left            =   1440
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      FoldHi          =   16777215
      FoldLo          =   192
      EdgeColumn      =   80
      EdgeMode        =   2
      IndentWidth     =   25
      UseTabs         =   -1  'True
      FoldAtElse      =   -1  'True
      FoldCompact     =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  sciMain.InitScintilla Me.hWnd
  sciMain.LoadAPIFile App.Path & "\cpp.api"
  hlMain.LoadHighlighters App.Path
  hlMain.SetHighlighter sciMain, "CPP"
End Sub

Private Sub Form_Resize()
  sciMain.MoveSCI 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
