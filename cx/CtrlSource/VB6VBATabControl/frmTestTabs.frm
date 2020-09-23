VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestTabs 
   Caption         =   "Visual Studio Tab Tester"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestTabs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelTab 
      Caption         =   "&Remove Tab"
      Height          =   435
      Left            =   4800
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdTabProperties 
      Caption         =   "&Tab Info"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   60
      Width           =   975
   End
   Begin VB.CheckBox chkNoClose 
      Caption         =   "No &Close Button"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   60
      Width           =   1515
   End
   Begin VB.TextBox txtTest 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   1740
      Width           =   1155
   End
   Begin VB.PictureBox picEvents 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   2820
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   5
      Top             =   720
      Width           =   2655
      Begin VB.ListBox lstEvents 
         Height          =   540
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label lblEvents 
         Caption         =   "Events:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.CheckBox chkAllowScroll 
      Caption         =   "Allow &Scroll"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkNoTabs 
      Caption         =   "&No Tabs"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   60
      Width           =   1515
   End
   Begin VB.CheckBox chkIcons 
      Caption         =   "&Icons"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkTabBottom 
      Caption         =   "&Tabs at Bottom"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   4020
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTabs.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTabs.frx":13CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTabs.frx":1526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTabs.frx":1680
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTabs.frx":17DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTabs.frx":1934
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbalDTab6.vbalDTabControl tabTest 
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   3942
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
   End
   Begin VB.Menu mnuContextTOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "&Tabs At Bottom"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Icons"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Allow &Scroll"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuContext 
         Caption         =   "No &Close Button"
         Index           =   3
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&No Tabs"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmTestTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' remove border from ListBox
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Enum EWindowLongIndexes
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum
' General window styles:
Private Enum EExWindowStyles
     WS_EX_DLGMODALFRAME = &H1
     WS_EX_NOPARENTNOTIFY = &H4
     WS_EX_TOPMOST = &H8
     WS_EX_ACCEPTFILES = &H10
     WS_EX_TRANSPARENT = &H20
     WS_EX_MDICHILD = &H40
     WS_EX_TOOLWINDOW = &H80
     WS_EX_WINDOWEDGE = &H100
     WS_EX_CLIENTEDGE = &H200
     WS_EX_CONTEXTHELP = &H400
     WS_EX_RIGHT = &H1000
     WS_EX_LEFT = &H0
     WS_EX_RTLREADING = &H2000
     WS_EX_LTRREADING = &H0
     WS_EX_LEFTSCROLLBAR = &H4000
     WS_EX_RIGHTSCROLLBAR = &H0
     WS_EX_CONTROLPARENT = &H10000
     WS_EX_STATICEDGE = &H20000
     WS_EX_APPWINDOW = &H40000
     WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
     WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum

Private Sub LogEvent(ByVal sMsg As String)
   lstEvents.AddItem sMsg
   lstEvents.ListIndex = lstEvents.ListCount - 1
End Sub

Private Sub chkAllowScroll_Click()
   If (chkAllowScroll.Value = vbChecked) Then
      tabTest.AllowScroll = True
      mnuContext(2).Checked = True
   Else
      tabTest.AllowScroll = False
      mnuContext(2).Checked = True
   End If
End Sub

Private Sub chkIcons_Click()
   If (chkIcons.Value = vbChecked) Then
      tabTest.ImageList = ilsIcons
      mnuContext(1).Checked = True
   Else
      tabTest.ImageList = 0
      mnuContext(1).Checked = False
   End If
End Sub

Private Sub chkNoClose_Click()
   If (chkNoClose.Value = vbChecked) Then
      tabTest.ShowCloseButton = False
      mnuContext(3).Checked = False
   Else
      tabTest.ShowCloseButton = True
      mnuContext(3).Checked = True
   End If
End Sub

Private Sub chkNoTabs_Click()
   If (chkNoTabs.Value = vbChecked) Then
      tabTest.ShowTabs = False
      mnuContext(5).Checked = True
   Else
      tabTest.ShowTabs = True
      mnuContext(5).Checked = True
   End If
End Sub

Private Sub chkTabBottom_Click()
   If (chkTabBottom.Value = vbChecked) Then
      tabTest.TabAlign = TabAlignBottom
      mnuContext(0).Checked = True
   Else
      tabTest.TabAlign = TabAlignTop
      mnuContext(0).Checked = False
   End If
End Sub

Private Sub cmdDelTab_Click()
   Dim cT As cTab
   Set cT = tabTest.SelectedTab
   If Not (cT Is Nothing) Then
      tabTest.Tabs.Remove cT.Key
   End If
End Sub

Private Sub cmdTabProperties_Click()
   Dim sMsg As String
   sMsg = "Tab Count: " & tabTest.Tabs.Count
   sMsg = sMsg & vbCrLf & vbCrLf & "Selected Tab:"
   Dim cT As cTab
   Set cT = tabTest.SelectedTab
   If cT Is Nothing Then
      sMsg = sMsg & vbCrLf & "None"
   Else
      sMsg = sMsg & vbCrLf & vbTab & "Caption:" & cT.Caption
      sMsg = sMsg & vbCrLf & vbTab & "IconIndex:" & cT.IconIndex
      sMsg = sMsg & vbCrLf & vbTab & "Tag:" & cT.Tag
      sMsg = sMsg & vbCrLf & vbTab & "ItemData:" & cT.ItemData
      sMsg = sMsg & vbCrLf & vbTab & "Key:" & cT.Key
      sMsg = sMsg & vbCrLf & vbTab & "CanClose:" & cT.CanClose
      sMsg = sMsg & vbCrLf & vbTab & "Enabled:" & cT.Enabled
      If (cT.Panel Is Nothing) Then
         sMsg = sMsg & vbCrLf & vbTab & "Tab has no Panel"
      Else
         sMsg = sMsg & vbCrLf & vbTab & "Panel: " & cT.Panel.Name
      End If
   End If
   sMsg = sMsg & vbCrLf
   MsgBox sMsg, vbInformation
   
   ' Check item objects:
   Dim i As Long
   With tabTest.Tabs
      For i = 1 To .Count
         Set cT = .Item(i)
         Debug.Print "Item "; i; " Caption="; cT.Caption; " Key="; cT.Key
      Next i
   End With
End Sub

Private Sub Form_Load()

   'Dim lStyle As Long
   'lStyle = GetWindowLong(lstEvents.hWnd, GWL_EXSTYLE)
   'lStyle = lStyle And Not WS_EX_CLIENTEDGE
   'SetWindowLong lstEvents.hWnd, GWL_EXSTYLE, lStyle

   Dim c As cTab
   With tabTest
      .ImageList = ilsIcons
      Set c = .Tabs.Add("SOLUTION", , "Solution Explorer")
      c.IconIndex = 0
      c.Panel = picEvents
      Set c = .Tabs.Add("CLASS", , "Class View")
      c.IconIndex = 1
      c.CanClose = False
      c.Panel = txtTest
      Set c = .Tabs.Add("CONTENTS")
      c.IconIndex = 2
      c.Caption = "Contents"
      Set c = .Tabs.Add("INDEX", , "Index")
      c.IconIndex = 3
      Set c = .Tabs.Add("SEARCH", , "Search")
      c.IconIndex = 4
   End With
   
   txtTest.Text = "vbAccelerator VS Tab Control Demonstration"
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next ' in case form is too small
   tabTest.Move _
      2 * Screen.TwipsPerPixelX, _
      tabTest.Top, _
      Me.ScaleWidth - 4 * Screen.TwipsPerPixelX, _
      Me.ScaleHeight - tabTest.Top - 2 * Screen.TwipsPerPixelY
End Sub

Private Sub Form_Terminate()
   If Forms.Count = 0 Then
      UnloadApp
   End If
End Sub

Private Sub mnuContext_Click(Index As Integer)
   Select Case Index
   Case 0
      If (chkTabBottom.Value = vbChecked) Then
         chkTabBottom.Value = vbUnchecked
      Else
         chkTabBottom.Value = vbChecked
      End If
   Case 1
      If (chkIcons.Value = vbChecked) Then
         chkIcons.Value = vbUnchecked
      Else
         chkIcons.Value = vbChecked
      End If
   Case 2
      If (chkAllowScroll.Value = vbChecked) Then
         chkAllowScroll.Value = vbUnchecked
      Else
         chkAllowScroll.Value = vbChecked
      End If
   Case 3
      If (chkNoClose.Value = vbChecked) Then
         chkNoClose.Value = vbUnchecked
      Else
         chkNoClose.Value = vbChecked
      End If
   Case 5
      If (chkNoTabs.Value = vbChecked) Then
         chkNoTabs.Value = vbUnchecked
      Else
         chkNoTabs.Value = vbChecked
      End If
   End Select
End Sub

Private Sub picEvents_Resize()
   On Error Resume Next ' in case it gets too small
   lblEvents.Width = picEvents.ScaleWidth - lstEvents.Left * 2
   lstEvents.Move lstEvents.Left, lstEvents.Top, picEvents.ScaleWidth - lstEvents.Left * 2, picEvents.ScaleHeight - lstEvents.Top - 2 * Screen.TwipsPerPixelY
End Sub

Private Sub tabTest_Resize()
   LogEvent "Resize"
End Sub


Private Sub tabTest_TabBarClick(ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
   LogEvent "TabBarClick: Button = " & iButton & ", Shift = " & Shift & ", X = " & x & ", Y = " & y
End Sub

Private Sub tabTest_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
   LogEvent "TabClick: '" & theTab.Caption & "', Button = " & iButton & ", Shift = " & Shift & ", X = " & x & ", Y = " & y
   If (iButton = vbRightButton) Then
      Me.PopupMenu mnuContextTOP, , x + tabTest.Left, y + tabTest.Top
   End If
End Sub

Private Sub tabTest_TabClose(theTab As vbalDTab6.cTab, bCancel As Boolean)
   LogEvent "TabClose: '" & theTab.Caption & "'"
   If (vbNo = MsgBox("Are you sure you want to close the " & theTab.Caption & " tab?", vbQuestion Or vbYesNo)) Then
      bCancel = True
   End If
End Sub

Private Sub tabTest_TabDoubleClick(theTab As vbalDTab6.cTab)
   LogEvent "TabDoubleClick: '" & theTab.Caption & "'"
End Sub

Private Sub tabTest_TabSelected(theTab As vbalDTab6.cTab)
   If Not (theTab Is Nothing) Then
      LogEvent "TabSelected: '" & theTab.Caption & "'"
   Else
      LogEvent "No Tab Selected"
   End If
End Sub
