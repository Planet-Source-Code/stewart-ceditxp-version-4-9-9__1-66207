VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uploading..."
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Picture         =   "frmUpload.frx":1042
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   90
      Width           =   240
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox cboAccount 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   135
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while file is uploaded....."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim URL As String, Port As String, User As String, Pass As String, SiteName As String
Public Sub PutFile(file As String, SaveString As String, Direc As String)
  'On Error Resume Next
  Dim hFile As Long, hDir As Boolean, FileINI As String
  Dim fFile As Integer, FTPInfo As FTP
  Dim hSize As Long, sBuffer As String, sizeLeft As Long, Ret As Long
  If hConnect = 0 Or hSession = 0 Then
    If cboAccount.Text = "" Then
      MsgBox "Please select an account to access first.", vbOKOnly + vbCritical, "Error"
      Exit Sub
    End If
    FileINI = App.path & "\accounts\" & cboAccount.Text & ".ftp"
    If Dir(FileINI) = "" Then
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
      Else
        InternetCloseHandle hConnect
        InternetCloseHandle hSession
        FTPError Err.LastDllError, "Put File"
        Exit Sub
      End If
    Else
      InternetCloseHandle hSession
      FTPError Err.LastDllError, "Put File"
      Exit Sub
    End If
  End If
  
  If hSession = 0 Or hConnect = 0 Then
    FTPError Err.LastDllError, "Put File"
    InternetCloseHandle hSession
    InternetCloseHandle hConnect
    hSession = 0: hConnect = 0
    Exit Sub
  End If
  
  hSize = Len(SaveString)
  PB.Max = hSize
  hFile = 0
  hDir = FtpSetCurrentDirectory(hConnect, Direc)
  If hDir = False Then
    hSession = 0: hConnect = 0
    MsgBox "Unable to set directory."
    Exit Sub
  End If
  hFile = FtpOpenFile(hConnect, file, GENERIC_WRITE, FTP_TRANSFER_TYPE_ASCII, 0)
  If hFile = 0 Then
    FTPError Err.LastDllError, "Put File"
    Exit Sub
  End If
  Do
    
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
        FTPError Err.LastDllError, "Put File"
        Exit Do
      End If
    Else
      If InternetWriteFile(hFile, sBuffer, sizeLeft, Ret) = 0 Then
        FTPError Err.LastDllError, "Put File"
        Exit Do
      End If
    End If
    If PB.Value + Ret > PB.Max Then PB.Max = PB.Value + Ret
    If Ret < sReadBuffer And PB.Max - Ret > PB.Value Then PB.Max = (PB.Value + Ret)
    PB.Value = PB.Value + Ret
  Loop Until SaveString = ""
  InternetCloseHandle hSession
  InternetCloseHandle hConnect
  hSession = 0: hConnect = 0
End Sub

Private Sub Form_Load()
  FlatBorder PB.hwnd
  LoadFormData Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
