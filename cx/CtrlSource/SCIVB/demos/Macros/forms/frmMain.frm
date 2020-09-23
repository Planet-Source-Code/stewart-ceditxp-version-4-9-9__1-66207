VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.0#0"; "SCIVBX.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin SCIVBX.SCIHighlighter hlMain 
      Left            =   3000
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SCIVBX.SCIVB sciMain 
      Left            =   1680
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnuMacro 
      Caption         =   "&Macros"
      Begin VB.Menu cmdStart 
         Caption         =   "Start Recording"
      End
      Begin VB.Menu cmdStop 
         Caption         =   "Stop Recording"
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu cmdSave 
         Caption         =   "Save Macro"
      End
      Begin VB.Menu cmdLoad 
         Caption         =   "Load Macro"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdPlay 
         Caption         =   "Play Macro"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
  sciMain.LoadMacro App.Path & "\Macro.mac"
End Sub

Private Sub cmdPlay_Click()
  sciMain.PlayMacro
End Sub

Private Sub cmdSave_Click()
  sciMain.SaveMacro App.Path & "\Macro.mac"
End Sub

Private Sub cmdStart_Click()
  sciMain.StartMacroRecord
  sciMain.SetFocus
End Sub

Private Sub cmdStop_Click()
  sciMain.StopMacroRecord
  sciMain.SetFocus
End Sub

Private Sub Form_Load()
  sciMain.InitScintilla Me.hWnd
End Sub

Private Sub Form_Resize()
  sciMain.MoveSCI 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuPlay_Click()

End Sub

Private Sub mnuPaste_Click()
  sciMain.Paste
End Sub
