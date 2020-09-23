VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.0#0"; "SCIVBX.ocx"
Begin VB.Form frmMain 
   Caption         =   "Basic SCIVB Demo"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin SCIVBX.SCIHighlighter hlMain 
      Left            =   4200
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   6600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6660
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdGoto 
         Caption         =   "Goto"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdRedo 
         Caption         =   "Redo"
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Undo"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdLine 
         Caption         =   "Line Text"
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdGetSel 
         Caption         =   "Get SelText"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "SelText"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Set Text"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
   End
   Begin SCIVBX.SCIVB sciMain 
      Left            =   2520
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit


Private Sub cmdAbout_Click()
  sciMain.ShowAbout
End Sub

Private Sub cmdFind_Click()
  sciMain.DoFind
End Sub

Private Sub cmdGetSel_Click()
  MsgBox sciMain.SelText
  sciMain.SetFocus
End Sub

Private Sub cmdGoto_Click()
  sciMain.DoGoto
End Sub

Private Sub cmdLine_Click()
  MsgBox sciMain.GetLineText(sciMain.GetCurrentLine)
  sciMain.SetFocus
End Sub

Private Sub cmdOpen_Click()
  sciMain.LoadFile App.Path & "\editor.cxx"
  hlMain.SetHighlighterExt sciMain, App.Path & "\editor.cxx"
End Sub

Private Sub cmdOptions_Click()
  hlMain.DoOptions App.Path & "\highlighters"
  hlMain.SetStylesAndOptions sciMain, "CPP"
  sciMain.SetFocus
End Sub

Private Sub cmdRedo_Click()
  sciMain.Redo
End Sub

Private Sub cmdReplace_Click()
  sciMain.DoReplace
End Sub

Private Sub cmdSel_Click()
  sciMain.SelText = "Added This Text"
End Sub

Private Sub cmdText_Click()
  sciMain.Text = "int main(){" & vbCrLf & "  // main function" & vbCrLf & "  return 0;" & vbCrLf & "}"
End Sub

Private Sub cmdUndo_Click()
  sciMain.Undo
End Sub

Private Sub Form_Load()
  sciMain.InitScintilla Me.hWnd
  hlMain.LoadHighlighters App.Path & "\highlighters"
  hlMain.SetStylesAndOptions sciMain, "CPP"
End Sub

Private Sub Form_Resize()
  sciMain.MoveSCI 0, (picTop.Height \ Screen.TwipsPerPixelY), Me.ScaleWidth, Me.ScaleHeight - (picTop.Height \ Screen.TwipsPerPixelY)
End Sub

