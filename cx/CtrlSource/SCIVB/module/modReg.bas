Attribute VB_Name = "modReg"
Option Explicit
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
'MakeFileType "txt", "Text Document", "C:\windows\notepad.exe,0", "open", "C:\windows\notepad.exe %1", False, True

Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function


Public Function ReadSetting(hKey As Long, strPath As String, strValue As String, DefaultStr As Long) As String
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    RegOpenKey hKey, strPath, keyhand
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                ReadSetting = Left$(strBuf, intZeroPos - 1)
            Else
                ReadSetting = strBuf
            End If
        End If
    End If
    If strBuf = "" Then ReadSetting = DefaultStr
End Function


Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    RegCreateKey hKey, strPath, keyhand
    RegSetValueEx keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata)
    RegCloseKey keyhand
End Sub
