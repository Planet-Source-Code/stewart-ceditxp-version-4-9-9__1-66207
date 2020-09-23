Attribute VB_Name = "modComboAutoComplete"
Option Explicit

Private Declare Function SendMessageB Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Const CB_SHOWDROPDOWN = &H14F
Const CB_FINDSTRING = &H14C
Const CB_GETLBTEXTLEN = &H149
Const CB_GETDROPPEDWIDTH = &H15F
Const CB_SETDROPPEDWIDTH = &H160

Type SIZE
    cx As Long
    cy As Long
End Type

Public Sub LockWindow(ByVal hwnd As Long)
Dim lRet As Long
    lRet = LockWindowUpdate(hwnd)
End Sub
Public Sub ReleaseWindow()
Dim lRet As Long
    lRet = LockWindowUpdate(0)
End Sub

Public Sub ComboDropdown(ByRef comboObj As ComboBox)
    Call SendMessageB(comboObj.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Public Sub ComboRetract(ByRef comboObj As ComboBox)
    Call SendMessageB(comboObj.hwnd, CB_SHOWDROPDOWN, 0, ByVal 0&)
End Sub

Public Function ComboAutoComplete(ByRef comboObj As ComboBox) As Boolean
Dim lngItemNum As Long
Dim lngSelectedLength As Long
Dim lngMatchLength As Long
Dim strCurrentText As String
Dim strSearchText As String
Dim sTypedText As String
Const CB_LOCKED = &H255

    With comboObj
        If .Text = Empty Then
            Exit Function
        End If
        Call LockWindow(.hwnd)
        If ((InStr(1, .Text, .Tag, vbTextCompare) <> 0 And Len(.Tag) = Len(.Text) - 1) Or (Left(.Text, 1) <> Left(.Tag, 1) And .Tag <> "")) And .Tag <> CStr(CB_LOCKED) Then
        
            strSearchText = .Text
            lngSelectedLength = Len(strSearchText)
        
            lngItemNum = SendMessageB(.hwnd, CB_FINDSTRING, -1, ByVal strSearchText)
            ComboAutoComplete = Not (lngItemNum = -1)
        
            If ComboAutoComplete Then
                lngMatchLength = Len(.List(lngItemNum)) - lngSelectedLength
                .Tag = CB_LOCKED
                sTypedText = strSearchText
                .Text = .Text & Right(.List(lngItemNum), lngMatchLength)
                .Tag = sTypedText
                .SelStart = lngSelectedLength
                .SelLength = lngMatchLength
            End If
        ElseIf .Tag <> CStr(CB_LOCKED) Then
            .Tag = .Text
        End If
        Call ReleaseWindow
    End With
End Function

Public Sub ComboDropWidth(ByRef comboObj As ComboBox)
Dim nCount As Long
Dim lNewDropDownWidth As Long
Dim lLongestString As Long

    On Error GoTo e_Trap
    For nCount = 0 To comboObj.ListCount - 1
        lNewDropDownWidth = comboObj.Parent.TextWidth(comboObj.List(nCount))
        If comboObj.Parent.ScaleMode = vbTwips Then
            lNewDropDownWidth = lNewDropDownWidth / Screen.TwipsPerPixelX  ' if twips change to pixels
        End If
        If lNewDropDownWidth > lLongestString Then
            lLongestString = lNewDropDownWidth
        End If
    Next nCount
    Call SendMessageB(comboObj.hwnd, CB_SETDROPPEDWIDTH, lLongestString + 25, 0)
    Exit Sub
e_Trap:
    Exit Sub
End Sub




