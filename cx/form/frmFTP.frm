VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cEdit - FTP"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin cEditXP.ctlFrame ctlFrame2 
      Height          =   630
      Left            =   120
      TabIndex        =   13
      Top             =   30
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1111
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   150
         Width           =   4635
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Browse"
         Height          =   390
         Left            =   5760
         TabIndex        =   14
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblAcnt 
         Caption         =   "Account :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   195
         Width           =   1005
      End
   End
   Begin cEditXP.ctlFrame ctlFrame1 
      Height          =   3975
      Left            =   5520
      TabIndex        =   8
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   7011
      Begin VB.TextBox txtFilter 
         Height          =   315
         Left            =   735
         TabIndex        =   10
         Text            =   "*.*"
         Top             =   3090
         Width           =   840
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "Close when complete"
         Height          =   345
         Left            =   60
         TabIndex        =   9
         Top             =   3495
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin MSComctlLib.Toolbar tbMain 
         Height          =   2970
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   5239
         ButtonWidth     =   2805
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Accounts"
               Key             =   "setup"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Make Directory"
               Key             =   "create"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Rename"
               Key             =   "rename"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Delete"
               Key             =   "delete"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&CHMOD (Unix)"
               Key             =   "chmod"
               Object.ToolTipText     =   "CHMOD (Unix)"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Command"
               Key             =   "command"
               Object.ToolTipText     =   "Execute Command"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Refresh"
               Key             =   "refresh"
               Object.ToolTipText     =   "Refresh Directory Listing"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   60
         TabIndex        =   12
         Top             =   3090
         Width           =   645
      End
   End
   Begin VB.TextBox txtFile 
      Height          =   345
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "You may type files in here. Syntax: ""File1"" ""File2"". Quotes must be around each file."
      Top             =   4800
      Width           =   4770
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   345
      Left            =   5985
      TabIndex        =   2
      Top             =   4800
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5985
      TabIndex        =   1
      Top             =   5265
      Width           =   1230
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   345
      Left            =   1095
      TabIndex        =   0
      Top             =   5265
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lstMain 
      Height          =   3840
      Left            =   120
      TabIndex        =   6
      Top             =   900
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   6773
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imgMain"
      SmallIcons      =   "imgMain"
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
         Text            =   "Filename"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Created"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Modified"
         Object.Width           =   3678
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   2760
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":1594
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":1AE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1320
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":2038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":2198
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":25EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":2A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":2B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":2CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":2E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":3E5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   690
      Width           =   4965
   End
   Begin VB.Label lblFile 
      Caption         =   "File Name:"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4890
      Width           =   3060
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress:"
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   5355
      Width           =   3060
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuAccounts 
         Caption         =   "Accounts"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakeDir 
         Caption         =   "&Make Directory"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuChmod 
         Caption         =   "&CHMOD (Unix)"
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "&Command"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim URL As String, Port As String, User As String, Pass As String, SiteName As String
Public SaveString As String
Dim TopDir As String, DoRun As Long

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdConnect_Click()
  Dim file As String
  Dim fFile As Integer, FTPInfo As FTP
  DoEvents
  If cboAccount.Text = "" Then
    MsgBox "Please select an account to access first.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  file = App.path & "\accounts\" & cboAccount.Text & ".ftp"
  If Dir(file) = "" Then
    MsgBox "There was an error reading the FTP file."
    Exit Sub
  End If
  DoEvents
  fFile = FreeFile()
  Open App.path & "\Accounts\" & cboAccount.Text & ".ftp" For Binary Access Read As #fFile
    Get #fFile, , FTPInfo
  Close #fFile
  URL = FTPInfo.URL
  Port = FTPInfo.PortNum
  User = FTPInfo.UserName
  Pass = Base64Decode(FTPInfo.Password)
  DoEvents
  If URL = "" Or Port = "" Or User = "" Or Pass = "" Then
    MsgBox "There was an error reading the FTP directory."
    Exit Sub
  End If
  DoEvents
  SiteName = cboAccount.Text
  hSession = InternetOpen(SiteName, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
  DoEvents
  If hSession <> 0 Then
    hConnect = InternetConnect(hSession, URL, Port, User, Pass, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, &H0)
    If hConnect <> 0 Then
      TopDir = GetFTPDirectory(hConnect)
      ListDir lstMain
      tbMain.Buttons(3).Enabled = True
      tbMain.Buttons(4).Enabled = True
      tbMain.Buttons(5).Enabled = True
      tbMain.Buttons(9).Enabled = True
      tbMain.Buttons(6).Enabled = True
      tbMain.Buttons(7).Enabled = True
    Else
      
      InternetCloseHandle hConnect
      InternetCloseHandle hSession
      MsgBox "Unable to connect.", vbOKOnly + vbCritical, "Error"
    End If
  Else
    InternetCloseHandle hSession
    MsgBox "Unable to connect.", vbOKOnly + vbCritical, "Error"
  End If
  If FTPInfo.LastDir <> "" Then
    FtpSetCurrentDirectory hConnect, FTPInfo.LastDir
    ListDir lstMain
  End If
End Sub

Private Sub cmdOpen_Click()
  On Error Resume Next
  Dim OpenFile As String, Ret As String, FileStr As String, NewOpen As Boolean
  Dim fFile As Integer, FTPInfo As FTP
  NewOpen = False
  Dim file As String
  If hConnect = 0 Or hSession = 0 Then
    If cboAccount.Text = "" Then
      MsgBox "Please select an account to access first.", vbOKOnly + vbCritical, "Error"
      Exit Sub
    End If
    file = App.path & "\accounts\" & cboAccount.Text & ".ftp"
    If Dir(file) = "" Then
      MsgBox "There was an error reading the FTP file."
      Exit Sub
    End If
    fFile = FreeFile()
    Open App.path & "\Accounts\" & cboAccount.Text & ".ftp" For Binary Access Read As #fFile
      Get #fFile, , FTPInfo
    Close #fFile
    URL = FTPInfo.URL
    Port = FTPInfo.PortNum
    User = FTPInfo.UserName
    Pass = Base64Decode(FTPInfo.Password)
    
    If URL = "" Or Port = "" Or User = "" Or Pass = "" Then
      MsgBox "There was an error reading the FTP directory."
      Exit Sub
    End If
 
    SiteName = cboAccount.Text
    hSession = InternetOpen(SiteName, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    If hSession <> 0 Then
      hConnect = InternetConnect(hSession, URL, Port, User, Pass, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, &H0)
      If hConnect <> 0 Then
        NewOpen = True
      Else
        InternetCloseHandle hConnect
        InternetCloseHandle hSession
        MsgBox "Unable to connect.", vbOKOnly + vbCritical, "Error"
        Exit Sub
      End If
    Else
      InternetCloseHandle hSession
      MsgBox "Unable to connect.", vbOKOnly + vbCritical, "Error"
      Exit Sub
    End If
  If FTPInfo.LastDir <> "" Then
    FtpSetCurrentDirectory hConnect, FTPInfo.LastDir
    ListDir lstMain
  End If
  End If
  
  If cmdOpen.Caption = "&Open" Then
    LockWindowUpdate frmMain.hwnd
    FileStr = Replace(txtFile.Text, " ", "")
    OpenFile = SplitStr(FileStr, Ret)
    Dim FTPDir As String
    FTPDir = GetFTPDirectory(hConnect)
    If OpenFile = "" Then
      OpenFTP GetFile(txtFile.Text), txtFile.Text, FTPDir, cboAccount.Text
    Else
      Do Until OpenFile = ""
        OpenFTP GetFile(OpenFile), OpenFile, FTPDir, cboAccount.Text
        OpenFile = SplitStr(Ret, Ret)
      Loop
    End If
    Dim cftp As FTP
    cftp.Name = SiteName
    cftp.UserName = User
    cftp.URL = URL
    cftp.PortNum = Port
    cftp.Password = Base64Encode(Pass)
    cftp.LastDir = FTPDir
    Open App.path & "\accounts\" & cftp.Name & ".ftp" For Binary Access Write As #1
      Put #1, , cftp
    Close #1
    LockWindowUpdate 0
  Else
    FileStr = Replace(txtFile.Text, " ", "")
    OpenFile = SplitStr(FileStr, Ret)
    If OpenFile = "" Then
      MsgBox OpenFile
      UploadAsString txtFile.Text, Document(dnum).sciMain.Text
    Else
      MsgBox OpenFile
      UploadAsString OpenFile, Document(dnum).sciMain.Text
    End If
    cftp.Name = SiteName
    cftp.UserName = User
    cftp.URL = URL
    cftp.PortNum = Port
    cftp.Password = Base64Encode(Pass)
    cftp.LastDir = GetFTPDirectory(hConnect)
    Open App.path & "\accounts\" & cftp.Name & ".ftp" For Binary Access Write As #1
      Put #1, , cftp
    Close #1
    
  End If
  
  If NewOpen = True Then
    InternetCloseHandle hConnect
    InternetCloseHandle hSession
    hConnect = 0: hSession = 0
  End If
  If chkClose = 1 Then Unload Me
End Sub

Private Sub Form_Load()
  GetAccounts cboAccount
  LoadFormData Me
  FlatBorder PB.hwnd
  If cboAccount.ListCount > 0 Then cboAccount.ListIndex = 0
  If cmdOpen.Caption = "&Open" Then
    lstMain.MultiSelect = True
  Else
    lstMain.MultiSelect = False
  End If
End Sub

Private Sub ListDir(lst As ListView, Optional ItemDir As String)
  'On Error GoTo errHandler
  Dim dt As WIN32_FIND_DATA
  Dim hFile As Long, sFile As Long
  Dim LstData As ListItem
  lst.ListItems.Clear
  Dim FTPDir As String
  If hConnect = 0 Or hSession = 0 Then
    MsgBox "You are not connected.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  DoEvents
  If ItemDir = "" Then
    FTPDir = GetFTPDirectory(hConnect)
  Else
    FTPDir = ItemDir
  End If
  DoEvents
  LockWindowUpdate lst.hwnd
  lst.ListItems.Add , , "Up a level", 1, 1
  hFile = FtpFindFirstFile(hConnect, txtFilter.Text, dt, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  DoEvents
  If hFile Then
    sFile = 1
  
    Do Until sFile = 0
    DoEvents
      If (dt.dwFileAttributes And vbDirectory) Then
        Set LstData = lst.ListItems.Add(, , StripCrap(dt.cFileName), 2, 2)
        LstData.SubItems(1) = "Directory"
        LstData.SubItems(2) = Win32ToVbTime(dt.ftCreationTime)
        LstData.SubItems(3) = Win32ToVbTime(dt.ftLastWriteTime)
        Set LstData = Nothing
      End If
      
      sFile = InternetFindNextFile(hFile, dt)
    Loop
    InternetCloseHandle hFile
    InternetCloseHandle sFile
  End If
  hFile = FtpFindFirstFile(hConnect, txtFilter.Text, dt, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  DoEvents
  If hFile Then
  
    sFile = 1
    Do Until sFile = 0
    DoEvents
      If (dt.dwFileAttributes And Not vbDirectory) Then
        Set LstData = lst.ListItems.Add(, , StripCrap(dt.cFileName), 3, 3)
        LstData.SubItems(1) = dt.nFileSizeLow
        LstData.SubItems(2) = Win32ToVbTime(dt.ftCreationTime)
        LstData.SubItems(3) = Win32ToVbTime(dt.ftLastWriteTime)
        Set LstData = Nothing
      End If
      
      sFile = InternetFindNextFile(hFile, dt)
    Loop
  End If
  DoEvents
  lblDir.Caption = GetFTPDirectory(hConnect)
  InternetCloseHandle hFile
  InternetCloseHandle sFile
  LockWindowUpdate 0
  Exit Sub
errhandler:
  FTPError Err.LastDllError, "ListDir"
  LockWindowUpdate 0
  Exit Sub
End Sub

Private Function StripCrap(str As String) As String
  StripCrap = Left(str, InStr(1, str, vbNullChar) - 1)
End Function

Private Sub Form_Unload(Cancel As Integer)
  Dim cftp As FTP
  SaveFormData Me
  If hConnect <> 0 Then
    cftp.Name = SiteName
    cftp.UserName = User
    cftp.URL = URL
    cftp.PortNum = Port
    cftp.Password = Base64Encode(Pass)
    cftp.LastDir = GetFTPDirectory(hConnect)
    Open App.path & "\accounts\" & cftp.Name & ".ftp" For Binary Access Write As #1
      Put #1, , cftp
    Close #1
  End If
  InternetCloseHandle hConnect
  InternetCloseHandle hSession
  hConnect = 0: hSession = 0
End Sub

Private Sub lstMain_Click()
  Dim x As Long, strFiles As String
  strFiles = ""
  
  For x = 1 To lstMain.ListItems.Count
    If lstMain.ListItems(x).SubItems(1) <> "" And lstMain.ListItems(x).Selected = True Then
      If strFiles <> "" Then strFiles = strFiles & " "
      strFiles = strFiles & StrWrap(lstMain.ListItems(x))
    End If
  Next
  txtFile.Text = strFiles
End Sub

Private Sub lstMain_DblClick()
  Dim FTPDir As String
  If hSession = 0 Or hConnect = 0 Then
    MsgBox "Error: You are not connected.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  If lstMain.SelectedItem.Index = 1 Then
    FTPDir = GetFTPDirectory(hConnect)
    If FTPDir <> "/" And FTPDir <> "\" Then
      If InStrRev(FTPDir, "/") = 1 Then
        FtpSetCurrentDirectory hConnect, "/"
      Else
        FtpSetCurrentDirectory hConnect, Left(FTPDir, InStrRev(FTPDir, "/") - 1)
      End If
      ListDir lstMain
    End If
    Exit Sub
  End If
  If lstMain.SelectedItem.SubItems(1) = "Directory" Then
    FtpSetCurrentDirectory hConnect, lstMain.SelectedItem.Text
    ListDir lstMain
  Else
    cmdOpen_Click
  End If
End Sub

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuFile
End Sub

Private Sub mnuAccounts_Click()
  frmAccount.Show vbModal, Me
End Sub

Private Sub mnuChmod_Click()
  Dim sCommand As String, sFile As String
  sCommand = InputStr("Please enter the chmod value. IE: 777", "CHMOD")
  If IsNumeric(sCommand) = False Then
    MsgBox "You must enter a numeric value.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  sFile = lstMain.SelectedItem.Text
  If sFile = "" Then
    MsgBox "You must select a file.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  FtpCommand hConnect, False, FTP_TRANSFER_TYPE_ASCII, "site chmod " & sCommand & " " & sFile, 0, 0
  ListDir lstMain
End Sub

Private Sub mnuCommand_Click()
  Dim sCommand As String, sFile As String
  sCommand = InputStr("Please enter the command", "Comment")
  sFile = lstMain.SelectedItem.Text
  If sFile = "" Then
    MsgBox "You must select a file.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  FtpCommand hConnect, False, FTP_TRANSFER_TYPE_ASCII, sCommand & " " & sFile, 0, 0
  ListDir lstMain
End Sub

Private Sub mnuDelete_Click()
      Dim x As Long
      If lstMain.SelectedItem.Text = "" Then Exit Sub
      For x = 1 To lstMain.ListItems.Count
        If lstMain.ListItems(x).Selected = True Then
          If lstMain.SelectedItem.SubItems(1) = "Directory" Then
            DoRun = FtpRemoveDirectory(hConnect, lstMain.ListItems(x))
          Else
            DoRun = FtpDeleteFile(hConnect, lstMain.ListItems(x))
          End If
        End If
      Next
      ListDir lstMain
End Sub

Private Sub mnuMakeDir_Click()
      Dim strCreate As String
      strCreate = InputStr("Please enter the new Directory's title.", "New Directory")
      If strCreate = "" Then Exit Sub
      DoRun = FtpCreateDirectory(hConnect, strCreate)
      If DoRun = 0 Then
        FTPError Err.LastDllError, "Create Directory"
        Exit Sub
      End If
            
      ListDir lstMain
End Sub

Private Sub mnuRefresh_Click()
ListDir lstMain
End Sub

Private Sub mnuRename_Click()
      Dim strRename As String
      strRename = InputStr("Please enter the new filename.", "Rename File", lstMain.SelectedItem.Text)
      If strRename = "" Then Exit Sub
      DoRun = FtpRenameFile(hConnect, lstMain.SelectedItem.Text, strRename)
      If DoRun = 0 Then
        FTPError Err.LastDllError, "Rename"
        Exit Sub
      End If
      ListDir lstMain
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "setup"
      frmAccount.Show vbModal, Me
    Case "refresh"
      ListDir lstMain
    Case "rename"
      mnuRename_Click
    Case "create"
      mnuMakeDir_Click
    Case "delete"
      mnuDelete_Click
    Case "chmod"
      mnuChmod_Click
    Case "command"
      mnuCommand_Click
  End Select
End Sub

Private Function GetFile(str As String) As String
  On Error Resume Next
  Dim hFile As Long, sBuffer As String, Ret As Long, strStore As String, tSize As Long
  tSize = ReturnSize(str)
  PB.Value = 0
  PB.Max = tSize
  If hSession = 0 Or hConnect = 0 Then
    MsgBox "Error: Not connected to a server.", vbOKOnly + vbCritical, "Error"
    Exit Function
  End If
  hFile = FtpOpenFile(hConnect, str, GENERIC_READ, FTP_TRANSFER_TYPE_ASCII, 0)
  If hFile = 0 Then
    MsgBox "Unable to open requested file from FTP server.", vbOKOnly + vbCritical, "Error"
    Exit Function
  End If
  sBuffer = Space(sReadBuffer)
  Do
    InternetReadFile hFile, sBuffer, sReadBuffer, Ret
    If PB.Value + Ret > PB.Max Then PB.Max = PB.Value + Ret
    If Ret < sReadBuffer And PB.Max - Ret > PB.Value Then PB.Max = (PB.Value + Ret)
    PB.Value = PB.Value + Ret
    If Ret <> sReadBuffer Then
      sBuffer = Left$(sBuffer, Ret)
    End If
    strStore = strStore & sBuffer
  Loop Until Ret <> sReadBuffer
  InternetCloseHandle hFile
  GetFile = strStore

End Function

Private Sub UploadAsString(File1 As String, Data As String)
  Dim hFile As Long, sizeLeft As Long, sBuffer As String, Ret As Long
  Dim SaveString As String, currBytes As Long, totalBytes As Long
  SaveString = Data
  currBytes = 0
  totalBytes = Len(Data)
  hFile = FtpOpenFile(hConnect, File1, GENERIC_WRITE, FTP_TRANSFER_TYPE_ASCII, 0)
  Do
    If hFile = 0 Then
      FTPError Err.LastDllError, "UploadAsFile"
      Exit Sub
    End If
    If Len(SaveString) >= sReadBuffer Then
      sBuffer = Left$(SaveString, sReadBuffer)
      SaveString = Mid(SaveString, sReadBuffer + 1)
    Else
      sBuffer = Left$(SaveString, Len(SaveString))
      SaveString = ""
    End If
    sizeLeft = Len(sBuffer)
    If sizeLeft = sReadBuffer Then
      If InternetWriteFile(hFile, sBuffer, sReadBuffer, Ret) = 0 Then
        Exit Do
      End If
    Else
      If InternetWriteFile(hFile, sBuffer, sizeLeft, Ret) = 0 Then
        Exit Do
      End If
    End If
    currBytes = currBytes + Ret
    If currBytes > totalBytes Then totalBytes = currBytes
    DoEvents
    PB.Max = totalBytes
    PB.Value = currBytes
  Loop Until currBytes >= totalBytes
  InternetCloseHandle (hFile)
  Document(dnum).Changed = False
  Document(dnum).Caption = txtFile.Text
  Document(dnum).FTPAccount = cboAccount.Text
  Document(dnum).FileName = txtFile.Text
  Document(dnum).FTP = True
  Document(dnum).FTPDir = CurDir
  'RaiseEvent Message(MUPLOADED)
End Sub
