VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmAddBuild 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Build Edit"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddBuild.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      Height          =   1965
      Left            =   240
      ScaleHeight     =   1905
      ScaleWidth      =   4395
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton optInput 
         Caption         =   "On"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optInput 
         Caption         =   "Off"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   22
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   19
         Top             =   360
         Width           =   2655
         Begin VB.OptionButton optRun 
            Caption         =   "On"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optRun 
            Caption         =   "Off"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   20
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   16
         Top             =   1560
         Width           =   2655
         Begin VB.OptionButton optCon 
            Caption         =   "On"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optCon 
            Caption         =   "Off"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   17
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Label lblLabel 
         Caption         =   "Run Application After Compiling:"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   26
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label lblLabel 
         Caption         =   "Input EXE Output Name:"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   25
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Show Compiler Output:"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmAddBuild.frx":1042
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picSettings 
      Height          =   1965
      Left            =   240
      ScaleHeight     =   1905
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "%s"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   ".."
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtCompiler 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboLanguage 
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         Caption         =   "Filename:"
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   "C Options:"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   "Compiler:"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLabel 
         Caption         =   "Extension:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblLabel 
         Caption         =   "Language:"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmAddBuild.frx":2084
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.TabStrip TB 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4524
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Object.ToolTipText     =   "Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Object.ToolTipText     =   "Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAddBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddNew()
  Dim UA() As String, LngCnt As Long
  cboLanguage.Clear
  cboLanguage.AddItem "C/C++"
  cboLanguage.AddItem "Basic"
  cboLanguage.AddItem "Java"
  cboLanguage.AddItem "Pascal"
  cboLanguage.AddItem "SQL"
  cboLanguage.AddItem "HTML"
  cboLanguage.AddItem "XML"
  cboLanguage.ListIndex = 0
  'UA = Split(Langs, Chr$(10))
  For LngCnt = 0 To UBound(UA) - 1: cboLanguage.AddItem UA(LngCnt): Next
  Erase UA
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdDone_Click()
  Dim file As String, dFile As String, FindItem As Long
  If txtFile.Text = "" Then
    MsgBox "You must enter a filename to write this to.", vbOKOnly + vbCritical, "No Filename"
    Exit Sub
  End If
  If txtCompiler.Text = "" Then
    MsgBox "You must choose a compiler to compile these applications.", vbOKOnly + vbCritical, "No Compiler"
    Exit Sub
  End If
  FindItem = InStr(1, txtFile.Text, ".")
  If FindItem <> 0 Then
    dFile = Mid(txtFile.Text, 1, FindItem - 1)
  Else
    dFile = txtFile.Text
  End If
  dFile = dFile & ".cmp"
  file = App.path & "\compile\" & dFile
  writeini "Compile", "Language", cboLanguage.Text, file
  writeini "Compile", "Extension", txtExt.Text, file
  writeini "Compile", "Compile", txtCompiler.Text, file
  writeini "Compile", "Variables", txtOptions.Text, file
  If optRun(0).Value = True Then
    writeini "Compile", "RunWhenComplete", "on", file
  Else
    writeini "Compile", "RunWhenComplete", "off", file
  End If
  If optInput(0).Value = True Then
    writeini "Compile", "InputForOutput", "on", file
  Else
    writeini "Compile", "InputForOutput", "off", file
  End If
  If optCon(0).Value = True Then
    writeini "Compile", "CaptureOutput", "on", file
  Else
    writeini "Compile", "CaptureOutput", "off", file
  End If
  frmBuild.LoadBuild
  Unload Me
End Sub

Private Sub cmdOpen_Click()
  frmMain.cd.FileName = ""
  frmMain.cd.DialogTitle = "Select Compiler"
  frmMain.cd.Filter = "Executables|*.exe"
  frmMain.cd.ShowOpen
  If frmMain.cd.FileName = "" Then Exit Sub
  txtCompiler.Text = frmMain.cd.FileName
End Sub

Private Sub Form_Load()
  FlatBorder cmdDone.hwnd
  LoadFormData Me
End Sub

Public Sub edit(strFile As String)
  Dim pStore As String, x As Integer
  strFile = App.path & "\compile\" & strFile
  pStore = ReadINI("Compile", "Language", strFile)
  Dim UA() As String, LngCnt As Long
  Dim I As Long
  cboLanguage.Clear
  'cboLanguage.ListIndex = 0
  ReDim UA(iLngCount)
  For I = 0 To iLngCount - 1
    UA(I) = frmMain.mnuHighlighter(I).Caption
  Next
  For LngCnt = 0 To UBound(UA) - 1: cboLanguage.AddItem UA(LngCnt): Next
  Erase UA
  
  For x = 0 To cboLanguage.ListCount - 1
    If cboLanguage.List(x) = pStore Then
      cboLanguage.ListIndex = x
      Exit For
    End If
  Next
  txtCompiler.Text = ReadINI("Compile", "Compile", strFile)
  txtOptions.Text = ReadINI("Compile", "Variables", strFile)
  txtExt.Text = ReadINI("Compile", "Extension", strFile)
  txtFile.Text = StripPath(strFile)
  If ReadINI("Compile", "RunWhenComplete", strFile) = "on" Then
    optRun(0).Value = True
  Else
    optRun(1).Value = True
  End If
  If ReadINI("Compile", "InputForOutput", strFile) = "on" Then
    optInput(0).Value = True
  Else
    optInput(1).Value = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub

Private Sub TB_Click()
  picSettings.Visible = False
  picOptions.Visible = False
  Select Case TB.SelectedItem.Index
    Case 1
      picSettings.Visible = True
    Case 2
      picOptions.Visible = True
  End Select
End Sub
