VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoolbarSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//Name        : clsCoolbarSaver
'//Author      : Abdalla Mahmoud
'//E-Mail      : la3toot@hotmail.com
'//Description : With this class you can
            'save the GUI of the coolbar control
Option Explicit

Private m_Coll     As New Collection
Private m_Extender As CoolBar
Private m_FileName As String

Public Sub SubClass(ByRef XCoolbar As Object)
    If Not TypeOf XCoolbar Is CoolBar Then Exit Sub
    If Not m_Extender Is Nothing Then Exit Sub
    Dim I As Long
    Dim tmpObj As Object
    Set m_Extender = XCoolbar
    For I = 1 To m_Extender.Bands.Count
        Set tmpObj = m_Extender.Bands(I).Child
        If Not tmpObj Is Nothing Then Call m_Coll.Add(tmpObj, tmpObj.Name)
    Next
    Set tmpObj = Nothing
End Sub

Public Function DoSave() As Boolean
  On Error Resume Next
    Dim PropBag As New PropertyBag
    Dim I As Long
    Dim FN As Long
    Dim lStr As Variant
    Dim CIndex As Long
    Err.Clear
    With m_Extender.Bands
        If .Count = 0 Then Exit Function
        For I = 1 To .Count
            With .Item(I)
                CIndex = IndexFromKey(.Child.Name)
                Call PropBag.WriteProperty("Width" & CIndex, .Width)
                Call PropBag.WriteProperty("NewRow" & CIndex, .NewRow)
                Call PropBag.WriteProperty("Visible" & CIndex, .Visible)
                Call PropBag.WriteProperty("Position" & CIndex, .Position)
                Call PropBag.WriteProperty("Child" & I, .Child.Name)
            End With
        Next
    End With
    lStr = PropBag.Contents
    FN = FreeFile
    Call Kill(m_FileName)
    Open m_FileName For Binary As FN
        Seek #FN, 1
        Put #FN, , lStr
    Close FN
    lStr = vbNullString
    Set PropBag = Nothing
    DoSave = CBool(Err.Number)
End Function

Public Function DoRead() As Boolean
    On Error GoTo FunErr
    Dim PropBag As New PropertyBag
    Dim I As Long
    Dim FN As Long
    Dim lStr As Variant
    Dim lVar As Variant
    Dim bytArr() As Byte
    Dim tmpChild As Object
    Dim CIndex As Long
    Err.Clear
    FN = FreeFile
    Open m_FileName For Binary As #FN
        Seek #FN, 1
        Get #FN, , lStr
    Close FN
    lVar = lStr
    bytArr = lStr 'lVar
    PropBag.Contents = bytArr
    If PropBag.ReadProperty("Width1", -1) = -1 Then
        Set PropBag = Nothing
        Exit Function
    End If
    With m_Extender.Bands
        For I = 1 To m_Coll.Count
            
            CIndex = PropBag.ReadProperty("Position" & I)
            With .Item(CIndex)
                .Visible = PropBag.ReadProperty("Visible" & I)
                Set tmpChild = m_Coll(I)
                Set tmpChild.Container = m_Extender
                Set .Child = tmpChild
                If CIndex <> 1 Then .NewRow = PropBag.ReadProperty("NewRow" & I)
                .Width = PropBag.ReadProperty("Width" & I)
            End With
        Next
    End With
    lStr = vbNullString
    Set PropBag = Nothing
    Set tmpChild = Nothing
    Erase bytArr
    DoRead = CBool(Err.Number)
    Exit Function
FunErr:
    Set PropBag = Nothing
    Set tmpChild = Nothing
    Erase bytArr
    DoRead = False
End Function

Private Function IndexFromKey(ByVal Key As String) As Long
    On Error Resume Next
    Dim lObj As Object
    Call Err.Clear
    Set lObj = m_Coll(Key)
    If Err.Number Then
        Call Err.Clear
        IndexFromKey = 0
    Else
        Dim I As Long
        For I = 1 To m_Coll.Count
            If m_Coll(I) Is lObj Then
                IndexFromKey = I
                Exit For
            End If
        Next
    End If
    Set lObj = Nothing
End Function

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
    m_FileName = vNewValue
End Property

Private Sub Class_Terminate()
    Set m_Extender = Nothing
    Set m_Coll = Nothing
End Sub
