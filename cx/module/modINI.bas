Attribute VB_Name = "modINI"
Option Explicit
#If Win16 Then


Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String) As Integer


Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String) As Integer
#Else

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Function ReadINI(Section, KeyName, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Function writeini(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    WritePrivateProfileString sSection, sKeyName, sNewString, sFileName
End Function


