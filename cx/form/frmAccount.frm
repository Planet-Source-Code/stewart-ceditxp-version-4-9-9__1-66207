VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4590
      TabIndex        =   4
      Text            =   "21"
      Top             =   2280
      Width           =   765
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   375
      Width           =   2775
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "ftp.microsoft.com"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "guest@unknow.com"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "anonymous"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox chkAnonym 
      Caption         =   "Anonymous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   3840
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
            Picture         =   "frmAccount.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstAccounts 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   375
      Left            =   4575
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line LineSep 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   135
      X2              =   5430
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LineSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   150
      X2              =   5400
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line LineSep 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   2625
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line LineSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2640
      X2              =   5400
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label lblConnection 
      Caption         =   "Connection:"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblUser 
      Caption         =   "Username:"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblURL 
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Changed As Boolean
Private Sub chkAnonym_Click()
  If -chkAnonym.Value = 0 Then
    txtUserName.Text = "anonymous"
    txtPassword.Text = "cEdit@cEdit.com"
  End If
  Changed = True
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  Kill App.path & "\Accounts\" & lstAccounts.SelectedItem.SubItems(1) & ".ftp"
  lstAccounts.ListItems.Remove lstAccounts.SelectedItem.Index
End Sub

Private Sub cmdNew_Click()
  txtConnect.Text = ""
  chkAnonym.Value = 1
  txtUserName.Text = "Anonymous"
  txtPassword.Text = "cEdit@cEdit.com"
  txtURL.Text = ""
End Sub

Private Sub cmdSave_Click()
  Dim FTPInfo As FTP, fFile As Integer
  FTPInfo.Anonymous = chkAnonym.Value
  FTPInfo.Name = txtConnect.Text
  FTPInfo.UserName = txtUserName.Text
  FTPInfo.Password = Base64Encode(txtPassword.Text)
  FTPInfo.URL = txtURL.Text
  FTPInfo.PortNum = txtPort.Text
  fFile = FreeFile()
  Open App.path & "\Accounts\" & FTPInfo.Name & ".ftp" For Binary Access Write As #fFile
    Put #fFile, , FTPInfo
  Close #fFile
  LoadAccounts
  Changed = False
End Sub

Private Sub Form_Load()
  On Error Resume Next
  lstAccounts.ColumnHeaders(2).Width = lstAccounts.Width - 90 - lstAccounts.ColumnHeaders(1).Width
  LoadFormData Me
  Changed = False
  LoadAccounts
  lstAccounts.ListItems(1).Selected = True
  lstAccounts_Click
End Sub
Private Sub LoadAccounts()
  Dim s As String, lstItem As ListItem
  lstAccounts.ListItems.Clear
  s = Dir(App.path & "\Accounts\")
  Do Until LenB(s) = 0
    If LCase(Right(s, 3)) = "ftp" Then
      Set lstItem = lstAccounts.ListItems.Add(, , , 1, 1)
      lstItem.SubItems(1) = Left(s, Len(s) - 4)
    End If
    s = Dir
  Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
  GetAccounts frmFTP.cboAccount
  SaveFormData Me
End Sub

Private Sub lstAccounts_Click()
  On Error Resume Next
  Dim fFile As Integer, FTPInfo As FTP
  fFile = FreeFile()
  Open App.path & "\Accounts\" & lstAccounts.SelectedItem.SubItems(1) & ".ftp" For Binary Access Read As #fFile
    Get #fFile, , FTPInfo
  Close #fFile
  txtConnect.Text = FTPInfo.Name
  txtUserName.Text = FTPInfo.UserName
  txtPassword.Text = Base64Decode(FTPInfo.Password)
  txtURL.Text = FTPInfo.URL
  txtPort.Text = FTPInfo.PortNum
  chkAnonym.Value = FTPInfo.Anonymous
  Changed = False
End Sub

Private Sub txtConnect_Change()
  Changed = True
End Sub

Private Sub txtURL_Change()
  Changed = True
End Sub

Private Sub txtUserName_Change()
  Changed = True
End Sub

Private Sub txtPassword_Change()
  Changed = True
End Sub

Private Sub txtPort_Change()
  Changed = True
End Sub

Private Sub txtConnect_GotFocus()
  txtConnect.BackColor = 14073525
  SelAll txtConnect
End Sub
Private Sub txtPort_GotFocus()
  txtPort.BackColor = 14073525
  SelAll txtPort
End Sub

Private Sub txtPort_LostFocus()
  txtPort.BackColor = vbWindowBackground
End Sub
Private Sub txtConnect_LostFocus()
  txtConnect.BackColor = vbWindowBackground
End Sub
Private Sub txtURL_GotFocus()
  txtURL.BackColor = 14073525
  SelAll txtURL
End Sub

Private Sub txtURL_LostFocus()
  txtURL.BackColor = vbWindowBackground
End Sub

Private Sub txtUserName_GotFocus()
  txtUserName.BackColor = 14073525
  SelAll txtUserName
End Sub

Private Sub txtUserName_LostFocus()
  txtUserName.BackColor = vbWindowBackground
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.BackColor = 14073525
  SelAll txtPassword
End Sub

Private Sub txtPassword_LostFocus()
  txtPassword.BackColor = vbWindowBackground
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim FTPInfo As FTP, fFile As Integer
  If Changed = True Then
    If MsgBox("Do you wish to save this FTP account?", vbYesNo + vbQuestion, "Save") = vbYes Then
      FTPInfo.Anonymous = chkAnonym.Value
      FTPInfo.Name = txtConnect.Text
      FTPInfo.UserName = txtUserName.Text
      FTPInfo.Password = Base64Encode(txtPassword.Text)
      FTPInfo.URL = txtURL.Text
      FTPInfo.PortNum = txtPort.Text
      fFile = FreeFile()
      Open App.path & "\Accounts\" & FTPInfo.Name & ".ftp" For Binary Access Write As #fFile
        Put #fFile, , FTPInfo
      Close #fFile
    End If
  End If
  Unload Me
End Sub

Private Sub SelAll(txt As TextBox)
  txt.SelStart = 0
  txt.SelLength = Len(txt.Text)
End Sub
