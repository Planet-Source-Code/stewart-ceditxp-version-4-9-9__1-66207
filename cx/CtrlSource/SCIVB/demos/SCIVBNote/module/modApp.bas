Attribute VB_Name = "modApp"
Option Explicit

Public Sub SetupStatus()
  ' Thank you to Carles P.V.  for his Excellent and simple to use
  ' ucStatusBar.  Because of it's style it can readily be included
  ' in the project so it has no dependancies.
  With frmMain.stbMain
    '-- Initialize statusbar
    Call .Initialize(SizeGrip:=True, ToolTips:=True)
    '-- Initialize icons list
    Call .InitializeIconList
    '-- Add icons
    'Call .AddIcon(LoadResPicture("MAIL", vbResIcon))
    'Call .AddIcon(LoadResPicture("USER", vbResIcon))
    'Call .AddIcon(LoadResPicture("TIP", vbResIcon))
    '-- Add panels
    Call .AddPanel(, , , [sbSpring], "Panel #1", , 0)
    Call .AddPanel(, 0, , [sbContents], "Panel #2", , 1)
    Call .AddPanel(, , , [sbSpring], "Last panel")
  End With
End Sub
