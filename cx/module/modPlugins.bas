Attribute VB_Name = "modPlugins"
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long



Public Sub AddPlugins(FormX As Object)

' This generic function will look for all plugins in a spesified directory.
' It will then query the plugin for identification and add the plugin
' to the main form.

Dim objTemp As Object
Dim sTemp As String
Dim sPlugin As String

'Now, we loop through all the plugin files and add them to the menus.
' In addition to this, we call a common function on the plugins that
' Identifies the plugins for us.
Dim s As String

s = Dir(App.path & "\plugin\")
Do Until s = ""
  If Right(s, 4) = ".dll" Then
    sPlugin = Mid(s, 1, Len(s) - 4) & ".clsPluginInterface"
    Set objTemp = CreateObject(sPlugin)
    sTemp = objTemp.Identify ' Run the function on the plugin to get the identification
    'add the plugin to the form's menus.
    AddMenu FormX, sTemp, sPlugin
    Set objTemp = Nothing
  End If
  s = Dir()
Loop

End Sub
Public Sub RunPlugin(sPlugin As String, FormX As Form)

'On Error GoTo Error_H

    'Declare a clean object to use
    Dim objPlugIn As Object
    Dim strResponse As String
    ' Run the Plugin
    'Set objPlugIn = CreateObject(Combo1.Text)
    Set objPlugIn = CreateObject(sPlugin)
    strResponse = objPlugIn.Run(FormX)
    'MsgBox FormX.Name
    'if the plug-in returns an error, let us know
    If strResponse <> vbNullString Then
        MsgBox strResponse
    End If
    
Exit Sub

Error_H:

MsgBox sPlugin & " - Error executing the plugin" & vbCrLf & Err.Description

End Sub


Public Function AddMenu(FormX As Object, sCaption As String, sTag As String) As Integer

'On Error Resume Next
Dim iIndex As Integer

iIndex = (FormX.mnuPlugin.Count - 1) ' Get the position (Index) of where the plugin must go.
If FormX.mnuPlugin(0).Enabled = True Then iIndex = iIndex + 1
With FormX
  If iIndex <> 0 Then Load .mnuPlugin(iIndex)
  .mnuPlugin(iIndex).Caption = sCaption ' sCaption we got from the "Identify" function on the plugin
  .mnuPlugin(iIndex).Visible = True
  .mnuPlugin(iIndex).Enabled = True
  .mnuPlugin(iIndex).Tag = sTag ' We store the interface to the plugin in here, to later use it on the event of a menu click
End With

End Function

'Public Sub LoadTemplates()
'  Dim s As String
'  s = Dir(App.path & "\templates\")
'  Do Until s = ""
'    If Right(s, 4) = ".tmp" Then
'      AddMenuTemp Left(s, Len(s) - 4), s
'    End If
'    s = Dir
'  Loop
'End Sub

'Private Sub AddMenuTemp(str As String, fle As String)
'  Dim iIndex As Integer
'  iIndex = frmMain.mnuTemplate.Count - 1
'  With frmMain
'
'    .mnuTemplate(iIndex).Caption = str
'    .mnuTemplate(iIndex).Visible = True
'    .mnuTemplate(iIndex).Enabled = True
'    .mnuTemplate(iIndex).Tag = fle
'  End With
'End Sub
'
'Public Function LoadTemplate(str As String)
'  Dim fNum As Integer, lang As String, txt As String, strText
'  fNum = FreeFile
'  Open App.path & "\templates\" & str For Input As #fNum
'    Input #fNum, lang
'  Close #fNum
'  doNew lang
'  fNum = FreeFile
'  Open App.path & "\templates\" & str For Input As #fNum
'    Input #fNum, strText
'    Do Until EOF(fNum)
'      Input #fNum, strText
'      txt = txt & strText & vbCrLf
'    Loop
'  Close #fNum
'  Document(fIndex).rt.Text = txt
'End Function

