Attribute VB_Name = "modEnumFonts"
Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long


Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long


Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Public Const TmPF_FIXED_PITCH = &H1
    Public Const TmPF_TRUETYPE = &H4
    Public Const RASTER_FONTTYPE = &H1
    Public Const TRUETYPE_FONTTYPE = &H4
    Public ShowFontType As Integer
    Public SelectedFont As String
    Public SelectedStyle As String
    Public SelectedSize As Integer
    Public fUnderline As Boolean
    Public fStrikethru As Boolean
    Public Const Lf_FACESIZE = 32


Type LOGFONT
    LfHeight As Long
    LfWidth As Long
    LfEscapement As Long
    LfOrientation As Long
    LfWeight As Long
    LfItalic As Byte
    LfUnderline As Byte
    LfStrikeOut As Byte
    LfCharSet As Byte
    LfOutPrecision As Byte
    LfClipPrecision As Byte
    LfQuality As Byte
    LfPitchAndFamily As Byte
    LfFaceName(Lf_FACESIZE) As Byte
    End Type

Type NEWTEXTMETRIC
    TmHeight As Long
    TmAscent As Long
    TmDescent As Long
    TmInternalLeading As Long
    TmExternalLeading As Long
    TmAveCharWidth As Long
    TmMaxCharWidth As Long
    TmWeight As Long
    TmOverhang As Long
    TmDigitizedAspectX As Long
    TmDigitizedAspectY As Long
    TmFirstChar As Byte
    TmLastChar As Byte
    TmDefaultChar As Byte
    TmBreakChar As Byte
    TmItalic As Byte
    TmUnderlined As Byte
    TmStruckOut As Byte
    TmPitchAndFamily As Byte
    TmCharSet As Byte
    NTmFlags As Long
    NTmSizeEM As Long
    NTmCellHeight As Long
    NTmAveWidth As Long
    End Type

Private Function EnumFontFamTypeProc(LFont As LOGFONT, Ntm As NEWTEXTMETRIC, ByVal FontType As Long, lParam As ComboBox) As Long

    Dim FontFaceName As String

    If ShowFontType = FontType Then
        FontFaceName = StrConv(LFont.LfFaceName, vbUnicode)
        lParam.AddItem Left$(FontFaceName, InStr(FontFaceName, vbNullChar) - 1)
    End If


EnumFontFamTypeProc = 1
End Function


Public Sub GetFontList(ListBx As ComboBox)

    ListBx.Clear
    ShowFontType = 4


  EnumFontFamilies GetDC(ListBx.hwnd), vbNullString, AddressOf EnumFontFamTypeProc, ListBx
End Sub
