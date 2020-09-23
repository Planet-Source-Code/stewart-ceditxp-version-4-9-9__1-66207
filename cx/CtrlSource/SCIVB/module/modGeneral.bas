Attribute VB_Name = "modGeneral"
Option Explicit

Public Enum dcShiftDirection
    lLeft = -1
    lRight = 0
End Enum


Public Function FileExists(strFile As String) As Boolean
  ' This is a generic function that uses the dir command
  ' to return a boolean value (true/false) on if a file exists.
  If Dir(strFile) = "" Then
    FileExists = False
  Else
    FileExists = True
  End If
End Function

Public Function IsNumericKey(KeyAscii As Integer) As Integer
  IsNumericKey = KeyAscii
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Function

Public Function Shift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long, ByVal lDirectionToShift As dcShiftDirection) As Long

    Const ksCallname As String = "Shift"
    On Error GoTo Procedure_Error
    Dim LShift As Long

    If lDirectionToShift Then 'shift left
        LShift = lValue * (2 ^ lNumberOfBitsToShift)
    Else 'shift right
        LShift = lValue \ (2 ^ lNumberOfBitsToShift)
    End If

    
Procedure_Exit:
    Shift = LShift
    Exit Function
    
Procedure_Error:
    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Procedure_Exit
End Function

Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long

    Const ksCallname As String = "LShift"
    On Error GoTo Procedure_Error
    LShift = Shift(lValue, lNumberOfBitsToShift, lLeft)
    
Procedure_Exit:
    Exit Function
    
Procedure_Error:
    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Procedure_Exit
End Function

Public Function GET_X_LPARAM(ByVal lParam As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)
    GET_X_LPARAM = CLng("&H" & Right(hexstr, 4))
End Function

Public Function GET_Y_LPARAM(ByVal lParam As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)
    GET_Y_LPARAM = CLng("&H" & Left(hexstr, 4))
End Function

' This function is utilized to return the modified position of the
' mousecursor on a window
Public Function GetWindowCursorPos(Window As Long) As POINTAPI
  Dim lP As POINTAPI
  Dim rct As RECT
  GetCursorPos lP
  GetWindowRect Window, rct
  GetWindowCursorPos.X = lP.X - rct.Left
  If GetWindowCursorPos.X < 0 Then GetWindowCursorPos.X = 0
  GetWindowCursorPos.Y = lP.Y - rct.Top
  If GetWindowCursorPos.Y < 0 Then GetWindowCursorPos.Y = 0
End Function

Function GetSHIFT() As Long

    'This function returns the state of the
    '     SHIFT, CONTROL and ALT keys
    'It does not distinguish the difference
    '     in left or right
    'Return value:
    'Bit 0=1 if pressed)
    Dim KS As Long
    Dim RetVal As Long
    KS = 0
    RetVal = GetKeyState(VK_SHIFT)


    If (RetVal And 32768) <> 0 Then
        KS = KS Or 1
    End If

    GetSHIFT = KS
End Function

Public Function piGetShiftState() As Integer
Dim iR As Integer
Dim lR As Long
Dim lKey As Long
    iR = iR Or (-1 * pbKeyIsPressed(VK_SHIFT))
    iR = iR Or (-2 * pbKeyIsPressed(VK_MENU))
    iR = iR Or (-4 * pbKeyIsPressed(VK_CONTROL))
    piGetShiftState = iR

End Function
Private Function pbKeyIsPressed( _
        ByVal nVirtKeyCode As KeyCodeConstants _
    ) As Boolean
Dim lR As Long
    lR = GetAsyncKeyState(nVirtKeyCode)
    If (lR And &H8000&) = &H8000& Then
        pbKeyIsPressed = True
    End If
End Function

Private Sub pGetHiWordLoWord( _
        ByVal lValue As Long, _
        ByRef lHiWord As Long, _
        ByRef lLoWord As Long _
    )
    lHiWord = lValue \ &H10000
    lLoWord = (lValue And &HFFFF&)
End Sub

Public Function Max(a As Long, b As Long) As Long
  If a > b Then
    Max = a
  Else
    Max = b
  End If
End Function


Public Function Byte2Str(bVal() As Byte) As String
  Dim i As Long
  If GetUpper(bVal) <> 0 Then
    For i = 0 To UBound(bVal())
      Byte2Str = Byte2Str & Chr(bVal(i))
    Next i
  End If
End Function

Public Function ShellDocument(sDocName As String, _
                    Optional ByVal action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState) As Boolean
    Dim Response
    Response = ShellExecute(&O0, action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
End Function


