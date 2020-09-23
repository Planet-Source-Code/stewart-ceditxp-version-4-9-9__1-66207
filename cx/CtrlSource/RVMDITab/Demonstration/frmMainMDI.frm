VERSION 5.00
Object = "{8C55C216-8E63-4117-B4EF-291DEC2D6A91}#2.0#0"; "RevMDITabs.ocx"
Begin VB.MDIForm frmMainMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Revelatek MDITabs Control Demonstration"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8955
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8955
      TabIndex        =   2
      Top             =   4740
      Width           =   8955
      Begin VB.Label Label5 
         Caption         =   "Copyright © 2004 Andrea Batina. All rights reserved."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   7455
      End
   End
   Begin VB.PictureBox picBar 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   6060
      ScaleHeight     =   4740
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox chkDrawIcons 
         BackColor       =   &H00F8F8F8&
         Caption         =   "Draw Icons"
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
         Left            =   180
         TabIndex        =   10
         Top             =   2820
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkDrawFocusRect 
         BackColor       =   &H00F8F8F8&
         Caption         =   "Draw Focus Rect"
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
         Left            =   180
         TabIndex        =   9
         Top             =   2520
         Width           =   1995
      End
      Begin VB.ComboBox cboStyle 
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
         ItemData        =   "frmMainMDI.frx":0000
         Left            =   180
         List            =   "frmMainMDI.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdchangecaption 
         Caption         =   "Change Caption"
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
         TabIndex        =   1
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Revelatek MDI Tabs control provides you with the ability to have Visual Studio.NET, Office 2003 and Office 2000 style tabs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MDI Tabs Control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Revelatek"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MDI Tabs drawing style:"
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
         Left            =   180
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Shape shpBorder 
         BackColor       =   &H00F8F8F8&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   4635
         Left            =   60
         Top             =   60
         Width           =   2775
      End
   End
   Begin RevMDITabs.RevMDITabsCtl RevMDITabsCtl1 
      Left            =   120
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
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
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindowTop 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboStyle_Click()
    RevMDITabsCtl1.Style = cboStyle.ListIndex
End Sub

Private Sub chkDrawFocusRect_Click()
    RevMDITabsCtl1.DrawFocusRect = chkDrawFocusRect.Value
End Sub

Private Sub chkDrawIcons_Click()
    RevMDITabsCtl1.DrawIcons = chkDrawIcons.Value
End Sub

Private Sub cmdchangecaption_Click()
    Me.ActiveForm.Caption = "NewCaption.txt *"
End Sub

Private Sub MDIForm_Load()
    mnuFileNew_Click
    cboStyle.ListIndex = 2
End Sub

Private Sub mnuFileNew_Click()
    Static lDocID As Long
    Dim frm As New frmChild
    lDocID = lDocID + 1
    frm.Caption = "Document " & lDocID
    frm.Show
End Sub
Private Sub mnuFileClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Unload ActiveForm
End Sub
Private Sub mnuPopupClose_Click()
    mnuFileClose_Click
End Sub

Private Sub picBar_Resize()
    shpBorder.Move 15, 15, picBar.Width - 30, picBar.Height - 15
End Sub

Private Sub RevMDITabsCtl1_ColorChanged(NewColor As stdole.OLE_COLOR)
    picBar.BackColor = NewColor
End Sub
Private Sub RevMDITabsCtl1_TabBarClick(Button As Integer, X As Long, Y As Long)
    Debug.Print "TabBarClick (" & Button & ", " & X & ", " & Y & ")"
End Sub
Private Sub RevMDITabsCtl1_TabClick(TabHwnd As Long, Button As Integer, X As Long, Y As Long)
    Debug.Print "TabClick (" & TabHwnd & ", " & Button & ", " & X & ", " & Y & ")"
    If Button = vbRightButton Then
        PopupMenu mnuPopup, vbPopupMenuLeftAlign, (X + 3) * Screen.TwipsPerPixelX, (Y + 22) * Screen.TwipsPerPixelY
    End If
End Sub
