VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mfrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Pinnable Tab Tester"
   ClientHeight    =   7755
   ClientLeft      =   4560
   ClientTop       =   3810
   ClientWidth     =   10605
   Icon            =   "mfrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   26
      Top             =   7440
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18177
            Text            =   "vbAccelerator Pinnable Tabs"
            TextSave        =   "vbAccelerator Pinnable Tabs"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10605
      TabIndex        =   25
      Top             =   0
      Width           =   10605
      Begin VB.CommandButton cmdTestFont 
         Caption         =   "Test Font"
         Height          =   315
         Left            =   60
         TabIndex        =   27
         Top             =   60
         Width           =   1335
      End
   End
   Begin vbalDTab6.vbalDTabControl tabLeft 
      Align           =   3  'Align Left
      Height          =   7065
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   12462
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Pinnable        =   -1  'True
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Index           =   1
         Left            =   900
         ScaleHeight     =   3795
         ScaleWidth      =   1695
         TabIndex        =   2
         Top             =   1980
         Visible         =   0   'False
         Width           =   1695
         Begin VB.CommandButton Command1 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   6
            Top             =   2880
            Width           =   1035
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   4
            Text            =   "vbAccelerator.com"
            Top             =   660
            Width           =   1515
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   3
            Text            =   "Pinnable Tabs"
            Top             =   60
            Width           =   1455
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   0
         Left            =   300
         ScaleHeight     =   3135
         ScaleWidth      =   1695
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         Begin VB.ListBox lstResults 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IntegralHeight  =   0   'False
            Left            =   60
            TabIndex        =   23
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Text            =   "Titles Only"
            Top             =   1500
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Text            =   "Platform SDK"
            Top             =   900
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Text            =   "Windows Hooks"
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "&Results:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   24
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "&Options:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   22
            Top             =   1260
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "&Filter By:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   21
            Top             =   660
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "&Look For:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1575
         End
      End
   End
   Begin vbalDTab6.vbalDTabControl tabRight 
      Align           =   4  'Align Right
      Height          =   7065
      Left            =   7935
      TabIndex        =   10
      Top             =   375
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   12462
      AllowScroll     =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Pinnable        =   -1  'True
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   3
         Left            =   360
         ScaleHeight     =   2775
         ScaleWidth      =   1695
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Text            =   "Pinnable Tabs"
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox Combo5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Text            =   "vbAccelerator.com"
            Top             =   780
            Width           =   1575
         End
         Begin VB.ComboBox Combo4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Text            =   "VB Code"
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Index           =   2
         Left            =   840
         ScaleHeight     =   3795
         ScaleWidth      =   1695
         TabIndex        =   11
         Top             =   1980
         Visible         =   0   'False
         Width           =   1695
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   15
            Text            =   "Pinnable Tabs"
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   14
            Text            =   "vbAccelerator.com"
            Top             =   660
            Width           =   1515
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1515
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   12
            Top             =   2880
            Width           =   1035
         End
      End
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   6480
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":13CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1680
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":17DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1934
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":1FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":2150
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":22AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mfrmMain.frx":2404
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New..."
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "mfrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lID As Long

Private Sub newDocument()
   m_lID = m_lID + 1
   Dim fD As New frmDocument
   fD.Caption = "Document " & m_lID
   fD.Show
End Sub

Private Sub cmdTestFont_Click()
Dim sF As New StdFont
   sF.Name = tabLeft.Font.Name
   If (tabLeft.Font.Size < 24) Then
      sF.Size = 24
      tabLeft.Font = sF
      sF.Bold = True
      tabLeft.SelectedFont = sF
   Else
      tabLeft.Font.Size = 8
      sF.Size = 8
      tabLeft.Font = sF
      sF.Bold = True
      tabLeft.SelectedFont = sF
   End If
End Sub

Private Sub MDIForm_Load()
   
   tabLeft.Pinned = False
   tabLeft.ImageList = ilsIcons
   Dim tabX As cTab
   Set tabX = tabLeft.Tabs.Add("EXPLORER", , "Solution Explorer", 0)
   tabX.Panel = picTab(0)
   tabX.ToolTipText = "Explore objects in your solution"
   Set tabX = tabLeft.Tabs.Add("CLASSVIEW", , "Class View", 1)
   Set tabX.Panel = picTab(1)
   tabX.ToolTipText = "Manage Classes in your project"
   tabX.Selected = True
   
   tabRight.Pinned = False
   tabRight.ImageList = ilsIcons
   Set tabX = tabRight.Tabs.Add("EXPLORER", , "Contents", 0)
   tabX.Panel = picTab(2)
   Set tabX = tabRight.Tabs.Add("CLASSVIEW", , "Search", 1)
   Set tabX.Panel = picTab(3)
   
   newDocument
     
   
End Sub

Private Sub MDIForm_Terminate()
   If Forms.Count = 0 Then
      UnloadApp
   End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
   Select Case Index
   Case 0
      newDocument
   Case 2
      Unload Me
   End Select
End Sub

Private Sub picTab_Resize(Index As Integer)
   On Error Resume Next ' may be too small
   Select Case Index
   Case 0
      Combo1.Move 2 * Screen.TwipsPerPixelX, Combo1.Top, picTab(0).ScaleWidth - 4 * Screen.TwipsPerPixelX
      Combo2.Move Combo1.Left, Combo2.Top, Combo1.Width
      Combo3.Move Combo1.Left, Combo3.Top, Combo1.Width
      lstResults.Move Combo1.Left, lstResults.Top, Combo1.Width, picTab(0).ScaleHeight - lstResults.Top - 2 * Screen.TwipsPerPixelY
      Dim i As Long
      For i = 0 To 3
         lblInfo(i).Move Combo1.Left, lblInfo(i).Top, Combo1.Width
      Next i
   Case 1
      Text1.Move 2 * Screen.TwipsPerPixelX, Text1.Top, picTab(1).ScaleWidth - 4 * Screen.TwipsPerPixelX
      Text2.Move Text1.Left, Text2.Top, Text1.Width
      List1.Move Text1.Left, List1.Top, Text1.Width
      Command1.Left = (picTab(1).ScaleWidth - Command1.Width) \ 2
   Case 2
      Text3.Move 2 * Screen.TwipsPerPixelX, Text3.Top, picTab(2).ScaleWidth - 4 * Screen.TwipsPerPixelX
      Text4.Move Text3.Left, Text4.Top, Text3.Width
      List2.Move Text3.Left, List2.Top, Text3.Width
      Command2.Left = (picTab(2).ScaleWidth - Command2.Width) \ 2
   Case 3
      Combo4.Move 2 * Screen.TwipsPerPixelX, Combo4.Top, picTab(3).ScaleWidth - 4 * Screen.TwipsPerPixelX
      Combo5.Move Combo4.Left, Combo5.Top, Combo4.Width
      Combo6.Move Combo4.Left, Combo6.Top, Combo4.Width
   End Select
End Sub

