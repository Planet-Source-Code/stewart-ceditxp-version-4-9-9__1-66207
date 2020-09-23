VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{A9E80832-87FC-4A90-A007-A87D78ADD7B3}#1.0#0"; "RevMDITabs.ocx"
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.1#0"; "SCIVBX.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "cEditMX"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   10515
   Icon            =   "mdiFrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin SCIVBX.SCIHighlighter Highlighters 
      Left            =   5520
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picSizeBot 
      Align           =   2  'Align Bottom
      Height          =   40
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   10515
      TabIndex        =   2
      Top             =   4845
      Width           =   10515
   End
   Begin VB.PictureBox picSizeLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3000
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4095
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   40
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5160
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6465
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img 
      Left            =   4440
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":136E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":1E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":1F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":2476
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":29C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":2F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":302C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":313E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":3250
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":3362
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":38B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":39C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":3F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":446A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":49BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":4F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":5460
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":59B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":5F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":6456
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":69A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":6EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":744C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":799E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":7EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":8442
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":8994
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":8EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":9438
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":998A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":9EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":A42E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":A980
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":AED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":B424
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":B976
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":BEC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":C41A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":C96C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":CEBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":DF10
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":E462
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":E9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":EF06
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":F458
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmMain.frx":F9AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbalDTab6.vbalDTabControl vsLeft 
      Align           =   3  'Align Left
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   7223
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
      ShowCloseButton =   0   'False
      Pinnable        =   -1  'True
      Begin VB.PictureBox picProject 
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   3120
         ScaleHeight     =   3495
         ScaleWidth      =   3615
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   3615
         Begin MSComctlLib.ImageList imgProj 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":FEFC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10056
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":101B0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":1030A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10464
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":105BE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10718
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10872
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":109CC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10B26
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10C80
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":10DDA
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView tvMain 
            Height          =   1770
            Left            =   2280
            TabIndex        =   14
            Top             =   840
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   3122
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   2
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imgProj"
            Appearance      =   1
         End
      End
      Begin VB.PictureBox picSnippet 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   -5760
         ScaleHeight     =   2655
         ScaleWidth      =   5520
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   5520
         Begin MSComctlLib.ListView lstSnippet 
            Height          =   1815
            Left            =   480
            TabIndex        =   12
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   3201
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "images"
            SmallIcons      =   "images"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            OLEDropMode     =   1
            NumItems        =   0
         End
      End
      Begin VB.PictureBox picFiles 
         BorderStyle     =   0  'None
         Height          =   3555
         Left            =   2520
         ScaleHeight     =   3555
         ScaleWidth      =   5520
         TabIndex        =   6
         Top             =   2760
         Visible         =   0   'False
         Width           =   5520
         Begin VB.PictureBox pic16 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2850
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   22
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox pic32 
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   4440
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   21
            Top             =   225
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.PictureBox picSize 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   50
            Left            =   240
            ScaleHeight     =   45
            ScaleWidth      =   495
            TabIndex        =   9
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   300
            TabIndex        =   8
            Top             =   0
            Width           =   2235
         End
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   315
            TabIndex        =   7
            Top             =   420
            Width           =   2220
         End
         Begin MSComctlLib.ListView File1 
            Height          =   1710
            Left            =   480
            TabIndex        =   10
            Top             =   2040
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   3016
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "File"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Path"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ImageList iml32 
            Left            =   360
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   -2147483644
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList iml16 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   -2147483644
            _Version        =   393216
         End
         Begin VB.Image imgSize 
            Height          =   50
            Left            =   360
            MouseIcon       =   "mdiFrmMain.frx":1132C
            MousePointer    =   99  'Custom
            Top             =   1920
            Width           =   2055
         End
      End
      Begin VB.PictureBox picTags 
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   -480
         ScaleHeight     =   2595
         ScaleWidth      =   6840
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   6840
         Begin MSComctlLib.ImageList images 
            Left            =   1500
            Top             =   390
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":1147E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":119D0
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TagsD 
            Height          =   1530
            Left            =   1725
            TabIndex        =   5
            Top             =   90
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   2699
            _Version        =   393217
            Indentation     =   5
            LineStyle       =   1
            Style           =   7
            ImageList       =   "images"
            Appearance      =   1
         End
      End
   End
   Begin ComCtl3.CoolBar cbMain 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1323
      BandCount       =   4
      _CBWidth        =   10515
      _CBHeight       =   750
      _Version        =   "6.7.9782"
      Child1          =   "tBar"
      MinWidth1       =   7500
      MinHeight1      =   330
      NewRow1         =   0   'False
      Child2          =   "tbSearch"
      MinHeight2      =   330
      Width2          =   4500
      NewRow2         =   0   'False
      Child3          =   "tbMacro"
      MinHeight3      =   330
      Width3          =   4455
      NewRow3         =   -1  'True
      Child4          =   "tbProgramming"
      MinHeight4      =   330
      Width4          =   6000
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar tbProgramming 
         Height          =   330
         Left            =   4650
         TabIndex        =   20
         Top             =   390
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tabl"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tabr"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cblock"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ublock"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tbmark"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nbmark"
               ImageIndex      =   22
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "pbmark"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cbmark"
               ImageIndex      =   24
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nline"
               ImageIndex      =   26
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "pline"
               ImageIndex      =   25
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ctag"
               ImageIndex      =   27
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbSearch 
         Height          =   330
         Left            =   7890
         TabIndex        =   18
         Top             =   30
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "findnext"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "findprev"
               ImageIndex      =   16
            EndProperty
         EndProperty
         Begin VB.ComboBox cmbFind 
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1695
         End
      End
      Begin MSComctlLib.Toolbar tbMacro 
         Height          =   330
         Left            =   165
         TabIndex        =   17
         Top             =   390
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac1"
               Object.ToolTipText     =   "Macro 1"
               ImageIndex      =   33
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac2"
               Object.ToolTipText     =   "Play Macro 2"
               ImageIndex      =   34
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac3"
               Object.ToolTipText     =   "Play Macro 3"
               ImageIndex      =   35
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac4"
               Object.ToolTipText     =   "Play Macro 4"
               ImageIndex      =   36
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac5"
               Object.ToolTipText     =   "Play Macro 5"
               ImageIndex      =   37
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac6"
               Object.ToolTipText     =   "Play Macro 6"
               ImageIndex      =   38
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac7"
               Object.ToolTipText     =   "Play Macro 7"
               ImageIndex      =   39
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac8"
               Object.ToolTipText     =   "Play Macro 8"
               ImageIndex      =   40
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac9"
               Object.ToolTipText     =   "Play Macro 9"
               ImageIndex      =   41
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mac10"
               Object.ToolTipText     =   "Play Macro 10"
               ImageIndex      =   42
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmac"
               Object.ToolTipText     =   "Create Macro"
               ImageIndex      =   48
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "smac"
               Object.ToolTipText     =   "Stop Macro Recording"
               ImageIndex      =   49
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tBar 
         Height          =   330
         Left            =   165
         TabIndex        =   16
         Top             =   30
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   26
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "New Document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open Document"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "close"
               Object.ToolTipText     =   "Close Open Document"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save Document"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "saveall"
               Object.ToolTipText     =   "Save All Open Documents"
               ImageIndex      =   44
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "saveas"
               Object.ToolTipText     =   "Save Document As"
               ImageIndex      =   45
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reload"
               Object.ToolTipText     =   "Reload Open Document"
               ImageIndex      =   46
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print Document"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               Object.ToolTipText     =   "Redo"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut Selected Text"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy Selected Text"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste Clipboard Contents"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "delete"
               Object.ToolTipText     =   "Delete Selected Text"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "prop"
               Object.ToolTipText     =   "Settings"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tilehor"
               Object.ToolTipText     =   "Tile Horizontaly"
               ImageIndex      =   28
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tilever"
               Object.ToolTipText     =   "Tile Verticly"
               ImageIndex      =   29
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cascade"
               Object.ToolTipText     =   "Cascade"
               ImageIndex      =   30
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "help"
               Object.ToolTipText     =   "Help"
               ImageIndex      =   31
            EndProperty
         EndProperty
      End
   End
   Begin vbalDTab6.vbalDTabControl vsBottom 
      Align           =   2  'Align Bottom
      Height          =   1575
      Left            =   0
      TabIndex        =   23
      Top             =   4890
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   2778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
      Begin VB.PictureBox picFFiles 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3120
         ScaleHeight     =   495
         ScaleWidth      =   1335
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         Begin cEditXP.VSFileSearch vs 
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
         End
      End
      Begin VB.PictureBox picTask 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   360
         ScaleHeight     =   735
         ScaleWidth      =   4215
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.PictureBox picFrame 
            Height          =   735
            Left            =   600
            ScaleHeight     =   675
            ScaleWidth      =   375
            TabIndex        =   27
            Top             =   0
            Width           =   430
            Begin cEditXP.ctlFrame cmdTBar 
               Height          =   735
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   1296
               Begin MSComctlLib.Toolbar tbBug 
                  Height          =   810
                  Left            =   75
                  TabIndex        =   29
                  Top             =   120
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   1429
                  ButtonWidth     =   423
                  ButtonHeight    =   476
                  AllowCustomize  =   0   'False
                  Style           =   1
                  ImageList       =   "imgMain"
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   3
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   1
                     EndProperty
                     BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   2
                     EndProperty
                     BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   3
                     EndProperty
                  EndProperty
               End
            End
         End
         Begin MSComctlLib.ListView lstTask 
            Height          =   780
            Left            =   1320
            TabIndex        =   30
            Top             =   0
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   1376
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Task ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Completed"
               Object.Width           =   2118
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ImageList imgMain 
            Left            =   3240
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483644
            ImageWidth      =   9
            ImageHeight     =   12
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":11F22
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":12404
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiFrmMain.frx":128E6
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picOutput 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   5760
         ScaleHeight     =   975
         ScaleWidth      =   2895
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtOut 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   25
            Top             =   0
            Width           =   1455
         End
      End
   End
   Begin RevMDITabs.RevMDITabsCtl MDITabs 
      Left            =   4200
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnuSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFTP 
         Caption         =   "&FTP"
         Begin VB.Menu mnuFTPOpen 
            Caption         =   "&Open From FTP"
         End
         Begin VB.Menu mnuSaveFTP 
            Caption         =   "&Save To FTP"
         End
      End
      Begin VB.Menu mnuSep24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export to HTML"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocProp 
         Caption         =   "&Document Properties"
      End
      Begin VB.Menu mnuSep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "&Recent Files"
         Begin VB.Menu mnuRec 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuRec 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRec 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRec 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRec 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRec 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSep28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "&Comment Block"
      End
      Begin VB.Menu mnuUncomment 
         Caption         =   "&Uncomment Block"
      End
      Begin VB.Menu mnuSep31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabLeft 
         Caption         =   "Tab &Right"
      End
      Begin VB.Menu mnuTabRight 
         Caption         =   "Tab &Left"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelLine 
         Caption         =   "Select &Line"
      End
      Begin VB.Menu mnuSep32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom &In"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom &Out"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDateTime 
         Caption         =   "&Date/Time"
      End
   End
   Begin VB.Menu mnuHigh 
      Caption         =   "&Highlighters"
      Begin VB.Menu mnuHighlighter 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindPrev 
         Caption         =   "Find &Prev"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindInFiles 
         Caption         =   "Find in F&iles"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggle 
         Caption         =   "&Toggle Bookmark"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuNextFlag 
         Caption         =   "&Next Bookmark"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPrevFlag 
         Caption         =   "&Prev Bookmark"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuDeleteFlags 
         Caption         =   "&Delete all Bookmarks"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNextLine 
         Caption         =   "&Next Line"
      End
      Begin VB.Menu mnuPrevLine 
         Caption         =   "&Previous Line"
      End
      Begin VB.Menu mnuSep37 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCountAll 
         Caption         =   "&Count All"
      End
   End
   Begin VB.Menu mnuCompile 
      Caption         =   "&Compile"
      Begin VB.Menu mnuBuild 
         Caption         =   "&Build/Compile"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigureBuild 
         Caption         =   "&Configure Build Settings"
      End
   End
   Begin VB.Menu mnuMacro 
      Caption         =   "&Macro"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Recording"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop &Recording"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveMacro 
         Caption         =   "Save &Macro"
      End
      Begin VB.Menu mnuLoadMacro 
         Caption         =   "&Load Macro"
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayMacro 
         Caption         =   "&Play Macro"
         Shortcut        =   ^{F5}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert String"
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPlugin 
         Caption         =   "No Plugins Installed"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSyntax 
         Caption         =   "&Configure Syntax"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuTileHo 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuTileVer 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuTileIcon 
         Caption         =   "Arrange &Icons"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuSep35 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "&Close All"
      End
      Begin VB.Menu mnuSep36 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowBrowser 
         Caption         =   "&Show File in Browser"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "&Window List"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuLinks 
      Caption         =   "&Links"
      Begin VB.Menu mnuCSite 
         Caption         =   "cEditMX Site"
      End
      Begin VB.Menu mnuVBA 
         Caption         =   "VBAccelerator"
      End
      Begin VB.Menu mnuPSC 
         Caption         =   "PlanetSourceCode"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Help &Index"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iChildCount As Long
Private strExt As String
Private bl As Boolean
Private Starting As Boolean
Private yBotTab As Long
Private m_CoolbarSaver As New clsCoolbarSaver
Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Sub Dir1_Change()
Dim path As String

Initialise
path = Dir1.path
FillFile1WithFiles path
GetAllIcons
ShowIcons
End Sub

Private Sub Highlighters_GetHighlighters(HighlighterName As String, HighlighterExt As String)
End Sub

Private Sub Highlighters_SetHighlighter(HighlighterName As String)
  On Error Resume Next
  Document(dnum).sciMain.LoadAPIFile App.path & "\api\" & HighlighterName & ".api"
End Sub

Private Sub Highlighters_AddHighlighter(HighlighterName As String, Filter As String)
  AddMenu HighlighterName, HighlighterName, iLngCount
  strExt = strExt & HighlighterExt
End Sub

Private Sub Highlighters_ClearHighlighters()
  iLngCount = 0
End Sub

Private Sub lstSnippet_DblClick()
  On Error Resume Next
  Dim fFile As Integer, str As String
  fFile = FreeFile()
  Open App.path & "\snippets\" & lstSnippet.SelectedItem.Text & ".snippet" For Input As #fFile
    str = Input(LOF(fFile), fFile)
  Close #fFile
  Call InsertString(Document(dnum).sciMain, str)
End Sub

Private Sub MDIForm_Activate()
  LoadCB
  If Starting = True Then
    doNew ""
    Starting = False
  End If
End Sub

Private Sub MDIForm_Load()
  bl = False
  Starting = True
'  fDock.GrabMain Me.hwnd
'  fDock.AddForm frmNav, tdDocked, tdAlignLeft, "frmNav", tdDockLeft Or tdDockFloat Or tdDockRight
'  fDock.Show
  ReadData
  LoadTasks
  LoadRecent
  LoadNav
  addTags
  AddPlugins Me
  Dim tabX As cTab
  With vsLeft
    .Pinned = False
    .ImageList = img
    Set tabX = .Tabs.Add("FILES", , "Files", 0)
    Set tabX.Panel = picFiles
    Set tabX = .Tabs.Add("TAGS", , "Tags", 1)
    Set tabX.Panel = picTags
    Set tabX = .Tabs.Add("PROJECTS", , "Project", 2)
    Set tabX.Panel = picProject
    Set tabX = .Tabs.Add("SNIPPETS", , "Snippets", 3)
    Set tabX.Panel = picSnippet
    vsLeft_Resize
  End With
  With vsBottom
    .ImageList = img
    Set tabX = .Tabs.Add("TODO", , "ToDo", 4)
    Set tabX.Panel = picTask
    Set tabX = .Tabs.Add("COMPILE", , "Compiler Output", 5)
    tabX.Panel = picOutput
    Set tabX = .Tabs.Add("FINDINFILES", , "Find in Files", 6)
    Set tabX.Panel = picFFiles
  End With
  Highlighters.LoadHighlighters App.path & "\highlighters"
  strExt = "All Files (*.*)|*.*|" & strExt
  cd.MaxFileSize = 5000
  
  FlatBorder picFrame.hwnd
  StopClose = False
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  StopClose = False
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  On Error Resume Next
  Dim frm As Form
  WriteData
  SaveCB
  For Each frm In Forms
    Unload frm
  Next
End Sub

Private Sub mnuAbout_Click()
  On Error Resume Next
  frmAbout.Show vbModal, Me
  Document(dnum).sciMain.SetFocus
End Sub

Private Sub mnuBuild_Click()
  On Error Resume Next
  Dim CaptureOut As String
  Dim s As String
  Dim lang As String, Exe As String, Comp As String, Variables As String
  Dim RunComp As String, InForOut As String, file As String, FileToCompile As String
  Dim Found As Boolean, VarRead As String
  'Dim dnum As Integer, Found As Boolean, VarRead As String, FileToCompile As String
  s = Dir(App.path & "\compile\")
  Found = False
  Do While s <> ""
    If Right(s, 3) = "cmp" Then
      file = App.path & "\compile\" & s
      lang = ReadINI("Compile", "Language", file)
      Exe = ReadINI("Compile", "Extension", file)
      If LCase(lang) = LCase(Document(dnum).sciMain.CurHigh) <> 0 Or GetExtension(Document(dnum).Caption) = LCase(Exe) Then
        Found = True
        Exit Do
      End If
    End If
    s = Dir
  Loop
  If Found = False Then
    MsgBox "No compiler found for this file type or language.", vbOKOnly + vbCritical, "Build"
    Exit Sub
  End If
  If Document(dnum).FTP = True Then
    Document(dnum).sciMain.SaveToFile App.path & "\data\tmp." & GetExtension(Document(dnum).FileName)
    FileToCompile = App.path & "\data\tmp." & GetExtension(Document(dnum).FileName)
  ElseIf Document(dnum).FTP = False And Document(dnum).FileName <> "" Then
    Document(dnum).sciMain.SaveToFile Document(dnum).FileName
    'DoSave
    FileToCompile = Document(dnum).FileName
    'Document(dnum).sciMain.SaveFile Document(dnum).filename, False
  Else
    FileToCompile = App.path & "\data\tmp." & Exe
    Document(dnum).sciMain.SaveToFile App.path & "\data\tmp." & Exe
  End If
  Comp = ReadINI("Compile", "Compile", file)
  Variables = ReadINI("Compile", "Variables", file)
  RunComp = ReadINI("Compile", "RunWhenComplete", file)
  InForOut = ReadINI("Compile", "InputForOutput", file)
  Variables = Replace(Variables, "%s", StrWrap(FileToCompile))
  CaptureOut = ReadINI("Compile", "CaptureOutput", file)
  If InForOut = "on" Then
    VarRead = InputStr("Enter the filename you would like this outputed to. (IE: hello.exe)", "Write Name")
    Variables = Replace(Variables, "%e", VarRead)
  End If
  If Dir(Comp) = "" Then
    MsgBox "Compiler not found.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  If CaptureOut = "on" Then
    txtOut.Text = "Compilation in progress..."
    DoEvents
    MDebugOutput.Checked = True
    ChDir Mid(FileToCompile, 1, InStrRev(FileToCompile, "\"))
    txtOut.Text = GetCommandOutput(StrWrap(Comp) & " " & Variables)
    txtOut.SelStart = Len(frmOutput.txtOut.Text)
  Else
    Shell StrWrap(Comp) & " " & Variables, vbNormalFocus
  End If
  If InForOut = "on" And RunComp = "on" Then
    Shell VarRead, vbNormalFocus
  End If

End Sub

Private Sub mnuCascade_Click()
  On Error Resume Next
  Me.Arrange vbCascade
End Sub

Private Sub mnuClose_Click()
  On Error Resume Next
  Unload Document(dnum)
End Sub

Private Sub mnuCloseAll_Click()
  On Error Resume Next
  CloseAllDoc
End Sub

Private Sub mnuComment_Click()
  On Error Resume Next
  'Highlighters.CommentBlock Document(dnum).sciMain
End Sub

Private Sub mnuConfigureBuild_Click()
  On Error Resume Next
  frmBuild.Show vbModal, Me
End Sub

Private Sub mnuCopy_Click()
  On Error Resume Next
  Document(dnum).sciMain.Copy
End Sub

Private Sub mnuCountAll_Click()
  On Error Resume Next
  Dim ua2() As String, us As Long, ut As Long
  ua2 = Split(Document(dnum).sciMain.Text, " ")
  us = Len(Document(dnum).sciMain.Text)
  ut = Document(dnum).sciMain.DirectSCI.GetLineCount
  MsgBox "Words: " & UBound(ua2) + 1 & Chr(10) & "Characters:" & us & Chr(10) & "Lines: " & ut, vbOKOnly + vbInformation, "Count All"
  Erase ua2
End Sub

Private Sub mnuCSite_Click()
  ShowSite "http://www.ceditmx.com"
End Sub

Private Sub mnuCut_Click()
  On Error Resume Next
  Document(dnum).sciMain.Cut
End Sub

Private Sub mnuDateTime_Click()
  On Error Resume Next
  Dim timedate As String
  timedate = Date & "/" & Time
  InsertString Document(dnum).sciMain, timedate
End Sub

Private Sub mnuDeleteFlags_Click()
  On Error Resume Next
  Document(dnum).sciMain.ClearBookmarks
End Sub

Private Sub mnuDocProp_Click()
  On Error Resume Next
  If ActiveForm.Name <> "frmDoc" Then Exit Sub
  Dim UA() As String, kB As Double
  kB = (Len(Document(dnum).sciMain.Text) / 1024)
  UA() = Split(Document(dnum).sciMain.Text, " ")
  With frmProperties
    .lblChar = "Characters: " & Len(Document(dnum).sciMain.Text)
    .lblLine = "Total Lines: " & Document(dnum).sciMain.DirectSCI.GetLineCount
    .lblWord = "Word Count: " & UBound(UA) + 1
    If Left(Document(dnum).Caption, 12) = "New Document" Then
      .lblFile = "File Name: " & "New Document"
    Else
      .lblFile = "File Name: " & Document(dnum).Caption
    End If
    .lblSizeK = "File Size(K): " & kB & " KBytes"
    .lblSizeB = "File Size(B): " & Len(Document(dnum).sciMain.Text) & " Bytes"
    .lblData(0).Caption = Document(dnum).Caption
    .Show vbModal, frmMain
  End With
  Erase UA
End Sub

Private Sub mnuExit_Click()
  End
    ' Showing that this method of subclassing doesn't seem to suffer
    ' from the same problems with crashing as the old method.
End Sub

Private Sub mnuExport_Click()
  On Error Resume Next
  With cd
    .Filter = "HTML Files (*.html)|*.html|"
    .ShowSave
    Call Highlighters.ExportToHTML(.FileName, Document(dnum).sciMain)
  End With
End Sub

Private Sub mnuFind_Click()
  On Error Resume Next
  Document(dnum).sciMain.DoFind
End Sub

Private Sub mnuFindInFiles_Click()
  frmFindInFiles.Show vbModal, Me
End Sub

Private Sub mnuFindNext_Click()
  On Error Resume Next
  Document(dnum).sciMain.FindNext
End Sub

Private Sub mnuFindPrev_Click()
  On Error Resume Next
  Document(dnum).sciMain.FindPrev
End Sub

Private Sub mnuFTPOpen_Click()
  On Error Resume Next
  frmFTP.Caption = "Open Document"
  frmFTP.cmdOpen.Caption = "&Open"
  frmFTP.Show , Me
End Sub

Private Sub mnuGoto_Click()
  On Error Resume Next
  Document(dnum).sciMain.DoGoto
End Sub

Private Sub mnuHelpIndex_Click()
  HHShowContents Me.hwnd
End Sub

Private Sub mnuHighlighter_Click(Index As Integer)
  On Error Resume Next
  Call Highlighters.SetStylesAndOptions(Document(dnum).sciMain, mnuHighlighter(Index).Tag)
  Call Document(dnum).sciMain.LoadAPIFile(App.path & "\apis\" & mnuHighlighter(Index).Caption & ".api")
  Document(dnum).SetCheck Index
End Sub

Private Sub mnuInsert_Click()
  On Error Resume Next
  Dim quicktag As String
  quicktag = InputStr("Enter the HTML tag to insert", "Quick Tag", "<>", 1)
  If quicktag <> "" Then Document(dnum).sciMain.SelText = quicktag
  Document(dnum).sciMain.SetFocus
End Sub

Private Sub mnuLoadMacro_Click()
  On Error Resume Next
  Load frmSaveMacro
  With frmSaveMacro
    .Caption = "Load Macro"
    .Show vbModal, Me
    If .DoWhat = 1 Then
      Document(dnum).sciMain.LoadMacro App.path & "\macros\" & .cmbSave.Text & ".mac"
    End If
  End With
  Unload frmSaveMacro
End Sub

Public Sub mnuNew_Click()
  doNew "NONE"
End Sub

Public Function AddMenu(sCaption As String, sTag As String, iIndex As Integer) As Integer
  Dim I As Long
  On Error Resume Next
  For I = 0 To iIndex - 1
    If mnuHighlighter(I).Caption = sCaption Then Exit Function
  Next I
  If iIndex > 0 Then Load mnuHighlighter(iIndex)
  mnuHighlighter(iIndex).Caption = sCaption ' sCaption we got from the "Identify" function on the plugin
  mnuHighlighter(iIndex).Visible = True
  mnuHighlighter(iIndex).Enabled = True
  mnuHighlighter(iIndex).Tag = sTag ' We store the interface to the plugin in here, to later use it on the event of a menu click
  iLngCount = iLngCount + 1
End Function

Private Sub mnuNextFlag_Click()
  On Error Resume Next
  Document(dnum).sciMain.NextBookmark
End Sub

Private Sub mnuNextLine_Click()
  On Error Resume Next
  Document(dnum).NextLine
End Sub

Private Sub mnuOpen_Click()
  On Error GoTo errHandle
  Dim vFiles() As String
  Dim lFile As Long
  With cd
    .FileName = ""
    .Filter = strExt
    .FLAGS = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Falgs, allows Multi select, Explorer style and hide the Read only tag
    .CancelError = True
    .ShowOpen
    vFiles = Split(.FileName, Chr(0)) 'Splits the filename up in segments
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
      DoOpen .FileName
    Else
      For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
        DoOpen vFiles(0) + "\" & vFiles(lFile)
      Next
    End If
  End With
  Err.Number = 0
errHandle:
  Erase vFiles()
  'exit with no effect cancel was probably pressed
End Sub

Private Sub mnuPaste_Click()
  On Error Resume Next
  Document(dnum).sciMain.Paste
End Sub

Private Sub mnuPlayMacro_Click()
  On Error Resume Next
  Document(dnum).sciMain.PlayMacro
End Sub

Private Sub mnuPrevFlag_Click()
  On Error Resume Next
  Document(dnum).sciMain.PrevBookmark
End Sub

Private Sub mnuPrevLine_Click()
  On Error Resume Next
  Document(dnum).PrevLine
End Sub

Private Sub mnuPrint_Click()
  On Error Resume Next
  Document(dnum).sciMain.PrintDoc
End Sub

Private Sub mnuPSC_Click()
  ShowSite "http://www.pscode.com"
End Sub

Private Sub mnuRec_Click(Index As Integer)
  On Error Resume Next
  DoOpen mnuRec(Index).Caption
End Sub

Private Sub mnuRedo_Click()
  On Error Resume Next
  Document(dnum).sciMain.Redo
End Sub

Private Sub mnuReplace_Click()
  On Error Resume Next
  Document(dnum).sciMain.DoReplace
End Sub

Private Sub mnuSave_Click()
  If Document(dnum).FTP = True And FState(dnum).Deleted = False Then
    frmUpload.cboAccount.Text = Document(dnum).FTPAccount
    frmUpload.cboAccount.Enabled = False
    DoEvents
    frmUpload.Show
    frmUpload.Refresh
    frmUpload.PutFile Document(dnum).FileName, Document(dnum).sciMain.Text, Document(dnum).FTPDir
    Document(dnum).Changed = False
    Document(dnum).FTP = True
    
    Unload frmUpload
    
  Else
    Save
  End If
End Sub

Public Sub Save()
  On Error Resume Next
  If Document(dnum).FileName = "" Then
    mnuSaveAs_Click
  Else
    Document(dnum).sciMain.SaveToFile Document(dnum).FileName
  End If
End Sub

Private Sub mnuSaveAll_Click()
  On Error Resume Next
  Dim x As Integer, Y As Integer
  Y = dnum
  For x = 1 To UBound(Document)
    Document(x).SetFocus
    Save
  Next
  Document(Y).SetFocus
End Sub

Private Sub mnuSaveAs_Click()
  doSaveAs
End Sub

Public Sub SaveAs()
  On Error GoTo errhandler
  Dim mResult As VbMsgBoxResult
Start:
  With cd
    .CancelError = True
    .Filter = strExt
    .ShowSave
    If Dir(.FileName) <> "" Then
      mResult = MsgBox("File: " & .FileName & vbCrLf & "This file already exists.  Do you wish to overwrite?", vbYesNo)
      If mResult = vbNo Then
        GoTo Start
      Else
        Document(dnum).FileName = .FileName
        Document(dnum).sciMain.SaveToFile .FileName
        Document(dnum).Caption = .FileName
        Highlighters.SetHighlighterExt Document(dnum).sciMain, .FileName
      End If
    Else
      Document(dnum).FileName = .FileName
      Document(dnum).sciMain.SaveToFile .FileName
      Document(dnum).Caption = .FileName
      Highlighters.SetHighlighterExt Document(dnum).sciMain, .FileName
    End If
  End With
  Document(dnum).sciMain.SetFocus
errhandler:
  'exit do nothing
End Sub

Private Sub mnuSaveFTP_Click()
  On Error Resume Next
  frmFTP.Caption = "Save Document"
  frmFTP.cmdOpen.Caption = "&Save"
  frmFTP.SaveString = ActiveForm.sciMain.Text
  frmFTP.Show
End Sub

Private Sub mnuSaveMacro_Click()
  On Error Resume Next
End Sub

Private Sub mnuSelAll_Click()
  On Error Resume Next
  Document(dnum).sciMain.SelectAll
End Sub

Private Sub mnuSelLine_Click()
  On Error Resume Next
  Document(dnum).sciMain.SelectLine
End Sub

Private Sub mnuShowBrowser_Click()
  On Error Resume Next
  ShowSite "about:" & Document(dnum).sciMain.Text
End Sub

Private Sub mnuStart_Click()
  On Error Resume Next
  Document(dnum).sciMain.StartMacroRecord
End Sub

Private Sub mnuStop_Click()
  On Error Resume Next
  Document(dnum).sciMain.StopMacroRecord
  Load frmSaveMacro
  With frmSaveMacro
    .Caption = "Save Macro"
    .Show vbModal, Me
    If .DoWhat = 1 Then
      Document(dnum).sciMain.SaveMacro App.path & "\macros\" & .cmbSave.Text & ".mac"
    End If
  End With
  Unload frmSaveMacro
  
End Sub

Private Sub mnuSyntax_Click()
'  On Error Resume Next
  Dim I As Integer
  bl = True
  strExt = ""
  Highlighters.DoOptions (App.path & "\highlighters")
  strExt = "All Files (*.*)|*.*|" & strExt
  For I = 1 To UBound(Document)
    If FState(dnum).Deleted = False Then
      Highlighters.SetStylesAndOptions Document(I).sciMain, Document(I).sciMain.CurHigh
    End If
  Next I
  If FState(dnum).Deleted = False Then
    Document(dnum).sciMain.SetFocus
  End If
End Sub

Private Sub mnuTabLeft_Click()
  On Error Resume Next
  'Document(dnum).sciMain.tableft
End Sub

Private Sub mnuTabRight_Click()
  On Error Resume Next
  Document(dnum).sciMain.TabRight
End Sub

Private Sub mnuTileHo_Click()
  On Error Resume Next
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileIcon_Click()
  On Error Resume Next
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuTileVer_Click()
  On Error Resume Next
  Me.Arrange vbTileVertical
End Sub

Private Sub mnuToggle_Click()
  On Error Resume Next
  Document(dnum).sciMain.ToggleMarker
End Sub

Private Sub mnuUncomment_Click()
  On Error Resume Next
  'Highlighters.UncommentBlock Document(dnum).sciMain
End Sub

Private Sub mnuUndo_Click()
  On Error Resume Next
  Document(dnum).sciMain.Undo
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub

Private Sub mnuVBA_Click()
  ShowSite "http://www.vbaccelerator.com"
End Sub

Private Sub mnuZoomIn_Click()
  On Error Resume Next
  'If Document(dnum).sciMain.Zoom < 150 Then
  '  Document(dnum).sciMain.Zoom = Document(dnum).sciMain.Zoom + 1
  'End If
End Sub

Private Sub mnuZoomOut_Click()
  On Error Resume Next
  'If Document(dnum).sciMain.Zoom > -150 Then
  '  Document(dnum).sciMain.Zoom = Document(dnum).sciMain.Zoom - 1
  'End If
End Sub

Private Sub tbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1
      mnuUndo_Click
    Case 2
      mnuRedo_Click
    Case 4
      mnuCut_Click
    Case 5
      mnuCopy_Click
    Case 6
      mnuPaste_Click
    Case 8
      mnuFind_Click
    Case 9
      mnuFindNext_Click
    Case 10
      mnuFindPrev_Click
    Case 12
      mnuToggle_Click
    Case 13
      mnuNextFlag_Click
    Case 14
      mnuPrevFlag_Click
    Case 15
      mnuDeleteFlags_Click
  End Select
End Sub

Private Sub tbGeneral_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Select Case Button.Index
    Case 1
      mnuTabLeft_Click
    Case 2
      mnuTabRight_Click
  End Select
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1
      mnuNew_Click
    Case 2
      mnuOpen_Click
    Case 4
      mnuSave_Click
    Case 5
      mnuSaveAs_Click
    Case 7
      mnuClose_Click
    Case 9
      mnuHelpIndex_Click
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Public Sub HHShowContents(lhWnd As Long)
    HTMLHelp lhWnd, App.path & "\cEdit.chm" & "", HH_DISPLAY_TOC, 0
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub

Private Sub picFFiles_Resize()
  On Error Resume Next
  vs.Move 0, 0, picFFiles.ScaleWidth, picFFiles.ScaleHeight
End Sub

Private Sub picFiles_Resize()
  On Error Resume Next
  Resize
End Sub

Private Sub picLeft_Click()

End Sub

Private Sub picOutput_Resize()
  On Error Resume Next
  txtOut.Move 0, 0, picOutput.ScaleWidth, picOutput.ScaleHeight
End Sub

Private Sub picProject_Resize()
  On Error Resume Next
  Resize
End Sub

Private Sub picSizeBot_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  If Button = vbLeftButton Then
    vsBottom.Height = (vsBottom.Height - Y)
    If vsBottom.Top > stbMain.Top Then stbMain.Top = vsBottom.Top + vsBottom.Height + 1000
  End If
End Sub

Private Sub picSizeLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picSizeLeft.BackColor = &H8000000C
End Sub

Private Sub picSizeLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = vbLeftButton Then
    vsLeft.Width = picSizeLeft.Left + x
  End If
End Sub

Private Sub picSizeLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picSizeLeft.BackColor = &H8000000F
End Sub

Private Sub picSnippet_Resize()
  On Error Resume Next
  Resize
End Sub

Private Sub picTags_Resize()
  On Error Resume Next
  Resize
End Sub

Private Sub picTask_Resize()
  On Error Resume Next
  picFrame.Move 20, 20, picFrame.Width, picTask.ScaleHeight - 45
  cmdTBar.Move 0, 0, picFrame.ScaleWidth, picFrame.ScaleHeight
  lstTask.Move picFrame.Width + 10, 10, picTask.ScaleWidth - picFrame.Width - 30, picTask.ScaleHeight - 45
  lstTask.ColumnHeaders(4).Width = lstTask.Width - lstTask.ColumnHeaders(2).Width - lstTask.ColumnHeaders(3).Width - 80
End Sub


Private Sub TagsD_DblClick()
  Dim timedate As String
  On Error Resume Next
  timedate = TagsD.SelectedItem.Text
  Document(dnum).sciMain.SelText = timedate
  Document(dnum).sciMain.SetFocus
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim quicktag As String
  Select Case Button.Key
    Case "new"
      mnuNew_Click
    Case "close"
      Unload Document(dnum)
    Case "prop"
      Highlighters.DoOptions App.path & "\highlighters\"
    Case "reload"
      If Document(dnum).isFile = False Then Exit Sub
      Document(dnum).sciMain.LoadFile Document(dnum).Caption
    Case "find"
      Document(dnum).sciMain.DoFind
    Case "findnext"
      Document(dnum).sciMain.FindNext
    Case "findprev"
      Document(dnum).sciMain.FindPrev
    Case "undo"
      Document(dnum).sciMain.Undo
    Case "saveas"
      mnuSaveAs_Click
    Case "delete"
      'Document(dnum).sciMain.DirectSC.SendEditor 2180  'Kind of hacked in
    Case "saveall"
      mnuSaveAll_Click
    Case "redo"
      Document(dnum).sciMain.Redo
    Case "tilever"
      Me.Arrange vbTileVertical
    Case "tilehor"
      Me.Arrange vbTileHorizontal
    Case "cascade"
      Me.Arrange vbCascade
    Case "cut"
      Document(dnum).sciMain.Cut
    Case "paste"
      Document(dnum).sciMain.Paste
    Case "copy"
      Document(dnum).sciMain.Copy
    Case "open"
      mnuOpen_Click
    Case "print"
      Call Document(dnum).sciMain.PrintDoc
    Case "save"
      mnuSave_Click
    Case "tabl"
      mnuTabLeft_Click
      'Document(dNum).sciMain.ExecuteCmd cmCmdIndentSelection
    Case "tabr"
      mnuTabRight_Click
      'Document(dNum).sciMain.ExecuteCmd cmCmdUnindentSelection
    Case "cblock"
      'Document(dNum).CommentBlock
    Case "ublock"
      'Document(dNum).UncommentBlock
    Case "tbmark"
      mnuToggle_Click
    Case "nbmark"
      mnuNextFlag_Click
    Case "pbmark"
      mnuPrevFlag_Click
    Case "cbmark"
      mnuDeleteFlags_Click
    Case "pline"
      'Document(dNum).PrevLine
    Case "nline"
      'Document(dNum).NextLine
    Case "ctag"
      quicktag = InputStr("Enter the HTML tag to insert", "Quick Tag", "<>", 1)
      If quicktag <> "" Then Document(dnum).sciMain.SelText = quicktag
    Case "help"
      HHShowContents Me.hwnd
  End Select

End Sub

Private Sub tbBug_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim iFile, iVal As Integer
  Select Case Button.Index
    Case 1
      frmTask.bEdit = False
      frmTask.txtName.Text = ""
      frmTask.txtDesc.Text = ""
      frmTask.hPer.Value = 0
      frmTask.Show vbModal, frmMain
    Case 2
      frmTask.bEdit = True
      frmTask.iItemNum = lstTask.SelectedItem.Index
      frmTask.txtName.Text = lstTask.SelectedItem.SubItems(1)
      frmTask.txtDesc.Text = lstTask.SelectedItem.SubItems(3)
      iVal = InStr(1, lstTask.SelectedItem.SubItems(2), "%")
      frmTask.hPer.Value = (Left(lstTask.SelectedItem.SubItems(2), iVal - 1) \ 10)
      frmTask.Show vbModal, frmMain
    Case 3
      lstTask.ListItems.Remove (lstTask.SelectedItem.Index)
  End Select
  iFile = FreeFile
  Open App.path & "\data\tasks.dat" For Output As #iFile
  For iVal = 0 To lstTask.ListItems.Count
      Print #iFile, lstTask.ListItems(iVal).SubItems(1) + "|" + lstTask.ListItems(iVal).SubItems(2) + "|" + lstTask.ListItems(iVal).SubItems(3)
  Next
  Close #iFile
End Sub

Private Sub tbMacro_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Select Case LCase(Button.Key)
    Case "mac1"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\0.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac2"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\1.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac3"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\2.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac4"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\3.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac5"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\4.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac6"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\5.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac7"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\6.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac8"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\7.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac9"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\8.mac"
      Document(dnum).sciMain.PlayMacro
    Case "mac10"
      Document(dnum).sciMain.LoadMacro App.path & "\macros\9.mac"
      Document(dnum).sciMain.PlayMacro
    Case "cmac"
      mnuStart_Click
    Case "smac"
      mnuStop_Click
  End Select
End Sub

Private Sub tbProgramming_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim quicktag As String
  Select Case Button.Key
    Case "tabl"
      Document(dnum).sciMain.TabRight
    Case "tabr"
      Document(dnum).sciMain.TabLeft
    Case "cblock"
      Highlighters.CommentBlock Document(dnum).sciMain
    Case "ublock"
      Highlighters.UncommentBlock Document(dnum).sciMain
    Case "tbmark"
      mnuToggle_Click
    Case "nbmark"
      mnuNextFlag_Click
    Case "pbmark"
      mnuPrevFlag_Click
    Case "cbmark"
      mnuDeleteFlags_Click
    Case "nline"
      Document(dnum).PrevLine
    Case "pline"
      Document(dnum).NextLine
    Case "ctag"
      quicktag = InputStr("Enter the HTML tag to insert", "Quick Tag", "<>", 1)
      If quicktag <> "" Then Document(dnum).sciMain.SelText = quicktag
      Document(dnum).sciMain.SetFocus
  End Select
End Sub

Private Sub tbSearch_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "find"
      If cmbFind.Text <> "" Then
        Document(dnum).sciMain.FindText cmbFind.Text
      Else
        Document(dnum).sciMain.DoFind
      End If
    Case "findnext"
      Document(dnum).sciMain.FindNext
    Case "findprev"
      Document(dnum).sciMain.FindPrev
  End Select
End Sub

Private Sub tvMain_DblClick()
  If Dir(tvMain.Nodes(tvMain.SelectedItem.Index).Key) <> "" Then DoOpen tvMain.Nodes(tvMain.SelectedItem.Index).Key
End Sub

Private Sub vs_DblClick(SelectedFile As String, LineNumber As String)
  DoOpen SelectedFile
  Document(dnum).sciMain.GotoLine Int(LineNumber) - 1
  Document(dnum).sciMain.SetFocus
End Sub

Private Sub vsLeft_Pinned()
  picSizeLeft.Visible = True
End Sub

Private Sub vsLeft_Resize()
  On Error Resume Next
  Resize
  MDITabs.RedrawControl
End Sub

Private Sub vsLeft_UnPinned()
  picSizeLeft.Visible = False
End Sub

Private Sub Resize()
  On Error Resume Next
  imgSize.Left = 0
  imgSize.Width = picFiles.ScaleWidth
  picSize.Move 0, imgSize.Top, imgSize.Width, imgSize.Height
  Drive1.Move 0, 30, picFiles.ScaleWidth
  Dir1.Move 0, Drive1.Top + Drive1.Height + 30, picFiles.ScaleWidth, imgSize.Top - Dir1.Top
  If Dir1.Height > (picFiles.ScaleHeight - 1500) Then Dir1.Height = picFiles.ScaleHeight - 1500
  imgSize.Move 0, Dir1.Top + Dir1.Height, picFiles.ScaleWidth
  File1.Move 0, imgSize.Top + imgSize.Height, picFiles.ScaleWidth, picFiles.Height - (imgSize.Top + imgSize.Height)
  TagsD.Move 0, 30, picTags.ScaleWidth, picTags.ScaleHeight - 30
  lstSnippet.Move 0, 30, picSnippet.ScaleWidth, picSnippet.ScaleHeight - 30
  tvMain.Move 0, 30, picProject.ScaleWidth, picProject.ScaleHeight - 30
End Sub

Private Sub LoadTasks()
  On Error Resume Next
  Dim iFile As Integer, iPos As Integer
  Dim strStore As String, strTemp As String
  Dim strHold(2) As String
  
  If FileExists(App.path & "\data\tasks.dat") = False Then Exit Sub
  iFile = FreeFile
  Open App.path & "\data\tasks.dat" For Input As #iFile
  Do Until EOF(iFile)
    Input #iFile, strStore
    iPos = InStr(1, strStore, "|")
    strTemp = Left(strStore, iPos - 1)
    strHold(0) = strTemp
    strTemp = Mid(strStore, iPos + 1)
    iPos = InStr(1, strTemp, "|")
    strHold(1) = Left(strTemp, iPos - 1)
    strHold(2) = Mid(strTemp, iPos + 1)
    With lstTask.ListItems.Add
      .SubItems(1) = strHold(0)
      .SubItems(2) = strHold(1)
      .SubItems(3) = strHold(2)
    End With
  Loop
End Sub


Public Sub LoadRecent()
  Dim FreeFileNum As Integer
  FreeFileNum = FreeFile()
  Open App.path & "\temp\recent.rct" For Binary Access Read As #FreeFileNum
    Get #FreeFileNum, , Recnt
  Close #FreeFileNum
  With frmMain
    If Recnt.Recent1 <> "" Then
      .mnuRec(0).Caption = Recnt.Recent1
      .mnuRec(0).Visible = True
    End If
    If Recnt.Recent2 <> "" Then
      .mnuRec(1).Caption = Recnt.Recent2
      .mnuRec(1).Visible = True
    End If
    If Recnt.Recent3 <> "" Then
      .mnuRec(2).Caption = Recnt.Recent3
      .mnuRec(2).Visible = True
    End If
    If Recnt.Recent4 <> "" Then
      .mnuRec(3).Caption = Recnt.Recent4
      .mnuRec(3).Visible = True
    End If
    If Recnt.Recent5 <> "" Then
      .mnuRec(4).Caption = Recnt.Recent5
      .mnuRec(4).Visible = True
    End If
    If Recnt.Recent6 <> "" Then
      .mnuRec(5).Caption = Recnt.Recent6
      .mnuRec(5).Visible = True
    End If
  End With
End Sub


Private Sub LoadNav()
  pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
  pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
  pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
  pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY
  imgSize.Top = 1920
  Initialise
  AddSnippets
  Dir1_Change
End Sub

Sub FillFile1WithFiles(ByVal path As String)
'-------------------------------------------
'Scan the selected folder for files
'and add then to the listview
'-------------------------------------------
Dim Item As ListItem
Dim s As String

path = CheckPath(path)    'Add '\' to end if not present
s = Dir(path, vbNormal)
Do While s <> ""
  Set Item = File1.ListItems.Add(, , s)
  Item.Key = path & s
  'Item.SmallIcon = "Folder"
  Item.Text = s
  Item.SubItems(1) = path
  s = Dir
Loop

End Sub


Private Sub Initialise()
'-----------------------------------------------
'Initialise the controls
'-----------------------------------------------
On Local Error Resume Next

'Break the link to iml lists
File1.ListItems.Clear
File1.Icons = Nothing
File1.SmallIcons = Nothing

'Clear the image lists
iml32.ListImages.Clear
iml16.ListImages.Clear

End Sub


Private Sub Drive1_Change()
  Dir1.path = Drive1.Drive
End Sub



Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In File1.ListItems
  FileName = Item.SubItems(1) & Item.Text
  GetIcon FileName, Item.Index
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection



'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    ImageList_Draw hLIcon, ShInfo.iIcon, pic32.hdc, 0, 0, ILD_TRANSPARENT
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    ImageList_Draw hSIcon, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function
Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the File1
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With File1
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub AddSnippets()
  Dim s As String
  s = Dir(App.path & "\snippets\")
  lstSnippet.ListItems.Clear
  Do Until s = ""
    If Right(s, 7) = "snippet" Then
      lstSnippet.ListItems.Add , , Left(s, Len(s) - 8), 2, 2
    End If
    s = Dir
  Loop
End Sub


Private Sub File1_DblClick()
  DoOpen Dir1.path & "\" & File1.SelectedItem.Text
End Sub

Private Sub SaveCB()
  ' This saves the coolbar layout (Thanks to Abdalla Mahmoud)
  On Error Resume Next
  Call m_CoolbarSaver.DoSave
  Set m_CoolbarSaver = Nothing
End Sub

Private Sub LoadCB()
  ' This loads the coolbar layout (Thanks to Abdalla Mahmoud)
  On Error Resume Next
  m_CoolbarSaver.FileName = App.path & "\settings\settings.ini"
  Call m_CoolbarSaver.SubClass(cbMain)
  Call m_CoolbarSaver.DoRead
End Sub


Private Sub mnuPlugin_Click(Index As Integer)
  On Error Resume Next
  Call RunPlugin(mnuPlugin(Index).Tag, Me) ' Execute the plug-in
End Sub


'**************************************************************
'* The following functions are for use with the plugin code   *
'**************************************************************

Public Sub AddText(str As String)
  If dnum = 0 Then Exit Sub
  InsertString Document(dnum).sciMain, str
End Sub

Public Sub MessageBox(Optional msgStr As String, Optional msgStyle As VbMsgBoxStyle, Optional msgTitle As String)
  MsgBox msgStr, msgStyle, msgTitle
End Sub

Private Sub CloseAllDoc()
  On Error Resume Next
  LockWindowUpdate Me.hwnd
  Dim x As Integer
  For x = 1 To UBound(Document)
    'Document(X).Visible = False
    Unload Document(x)
    
    If StopClose = True Then
      StopClose = False
      Exit For
    End If
  Next
  LockWindowUpdate 0
End Sub

