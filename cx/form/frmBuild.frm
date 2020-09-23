VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmBuild 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Build Settings"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmBuild.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgMain 
      Left            =   5400
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuild.frx":1042
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add New"
      Height          =   405
      Left            =   5370
      TabIndex        =   3
      Top             =   165
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   405
      Left            =   5370
      TabIndex        =   2
      Top             =   630
      Width           =   1035
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   405
      Left            =   5385
      TabIndex        =   1
      Top             =   1110
      Width           =   1035
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   405
      Left            =   5385
      TabIndex        =   0
      Top             =   1620
      Width           =   1035
   End
   Begin MSComctlLib.ListView lstMain 
      Height          =   3195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgMain"
      SmallIcons      =   "imgMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Extension"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Language"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Compiler"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Compiler Options"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "InputForOutput"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "RunWhenCompiled"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   5640
      Picture         =   "frmBuild.frx":2094
      Top             =   2280
      Width           =   240
   End
End
Attribute VB_Name = "frmBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub LoadBuild()
  Dim s As String, lst As ListItem
  Dim lang As String, Exe As String, Comp As String, Variables As String
  Dim RunComp As String, InForOut As String, file As String
  s = Dir(App.path & "\compile\")
  lstMain.ListItems.Clear
  Do While s <> ""
    If Right(s, 3) = "cmp" Then
      file = App.path & "\compile\" & s
      lang = ReadINI("Compile", "Language", file)
      Exe = ReadINI("Compile", "Extension", file)
      Comp = ReadINI("Compile", "Compile", file)
      Variables = ReadINI("Compile", "Variables", file)
      RunComp = ReadINI("Compile", "RunWhenComplete", file)
      InForOut = ReadINI("Compile", "InputForOutput", file)
      Set lst = lstMain.ListItems.Add(, , s, 1, 1)
      lst.SubItems(1) = Exe
      lst.SubItems(2) = lang
      lst.SubItems(3) = Comp
      lst.SubItems(4) = Variables
      lst.SubItems(5) = InForOut
      lst.SubItems(6) = RunComp
      Set lst = Nothing
    End If
    s = Dir
  Loop
  lstMain.Refresh
End Sub

Private Sub cmdDel_Click()
  DeleteFile App.path & "\compile\" & lstMain.SelectedItem.Text
  LoadBuild
End Sub

Private Sub cmdEdit_Click()
  frmAddBuild.edit lstMain.SelectedItem.Text
  frmAddBuild.Show vbModal, Me
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdNew_Click()
  frmAddBuild.AddNew
  frmAddBuild.Show vbModal, Me
End Sub

Private Sub Form_Load()
  FlatBorder cmdNew.hwnd
  LoadBuild
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
