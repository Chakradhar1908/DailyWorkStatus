Imports System.Runtime.InteropServices
Imports stdole
Public Class CommonDlgs
    '@NO-LINT-TYPE
    '
    'Common Dialog Open/Save file and Choose font without the CommonDialog OCX.
    '
    'No dependencies.
    '
    Private Const LOGPIXELSY As Integer = 90
    Private Const LF_FACESIZE As Integer = 32
    Dim f As New FontDialog

    Public Enum FileOpenConstants
        'ShowOpen, ShowSave constants.
        cdlOFNAllowMultiselect = &H200&
        cdlOFNCreatePrompt = &H2000&
        cdlOFNExplorer = &H80000
        cdlOFNExtensionDifferent = &H400&
        cdlOFNFileMustExist = &H1000&
        cdlOFNHideReadOnly = &H4&
        cdlOFNLongNames = &H200000
        cdlOFNNoChangeDir = &H8&
        cdlOFNNoDereferenceLinks = &H100000
        cdlOFNNointegerNames = &H40000
        cdlOFNNoReadOnlyReturn = &H8000&
        cdlOFNNoValidate = &H100&
        cdlOFNOverwritePrompt = &H2&
        cdlOFNPathMustExist = &H800&
        cdlOFNReadOnly = &H1&
        cdlOFNShareAware = &H4000&
    End Enum
    ''Case-preserving hack:
    '#If False Then
    'Dim cdlOFNAllowMultiselect, cdlOFNCreatePrompt, cdlOFNExplorer, cdlOFNExtensionDifferent
    'Dim cdlOFNFileMustExist, cdlOFNHideReadOnly, cdlOFNintegerNames, cdlOFNNoChangeDir
    'Dim cdlOFNNoDereferenceLinks, cdlOFNNointegerNames, cdlOFNNoReadOnlyReturn
    'Dim cdlOFNNoValidate, cdlOFNOverwritePrompt, cdlOFNPathMustExist, cdlOFNReadOnly
    'Dim cdlOFNShareAware
    '#End If

    Public Enum FontsConstants
        'ShowFont constants.
        cdlCFANSIOnly = &H400&
        cdlCFApply = &H200&
        cdlCFBoth = &H3&
        cdlCFEffects = &H100&
        cdlCFFixedPitchOnly = &H4000&
        cdlCFForceFontExist = &H10000
        cdlCFLimitSize = &H2000&
        cdlCFInitFont = &H40& 'Loads our Font property values into the dialog as defaults.
        cdlCFNoScriptSel = &H800000
        cdlCFNoFaceSel = &H80000
        cdlCFNoSimulations = &H1000&
        cdlCFNoSizeSel = &H200000
        cdlCFNoStyleSel = &H100000
        cdlCFNoVectorFonts = &H800&
        cdlCFPrinterFonts = &H2&
        cdlCFScalableOnly = &H20000
        cdlCFScreenFonts = &H1&
        cdlCFTTOnly = &H40000
        cdlCFWYSIWYG = &H8000&
    End Enum
    'Case-preserving hack:
    '#If False Then
    'Dim cdlCFANSIOnly, cdlCFApply, cdlCFBoth, cdlCFEffects, cdlCFFixedPitchOnly
    'Dim cdlCFForceFontExist, cdlCFLimitSize, cdlCFInitFont, cdlCFNoScriptSel
    'Dim cdlCFNoFaceSel, cdlCFNoSimulations, cdlCFNoSizeSel, cdlCFNoStyleSel
    'Dim cdlCFNoVectorFonts, cdlCFPrinterFonts, cdlCFScalableOnly, cdlCFScreenFonts
    'Dim cdlCFTTOnly, cdlCFWYSIWYG
    '#End If

    Private Structure LOGFONT
        Dim lfHeight As Integer
        Dim lfWidth As Integer
        Dim lfEscapement As Integer
        Dim lfOrientation As Integer
        Dim lfWeight As Integer
        Dim lfItalic As Byte
        Dim lfUnderline As Byte
        Dim lfStrikeOut As Byte
        Dim lfCharSet As Byte
        Dim lfOutPrecision As Byte
        Dim lfClipPrecision As Byte
        Dim lfQuality As Byte
        Dim lfPitchAndFamily As Byte
        'Dim lfFaceName(LF_FACESIZE - 1) As Byte
        Dim lfFaceName() As Byte
    End Structure

    Private Structure CHOOSEFONTType
        Dim lStructSize As Integer
        Dim hwndOwner As Integer
        Dim hDC As Integer
        'Dim lpLogFont As integer
        Dim lpLogFont As Integer
        Dim iPointSize As Integer
        Dim Flags As Integer
        'Dim rgbColors As integer
        Dim rgbColors As Color
        Dim lCustData As Integer
        Dim lpfnHook As Integer
        Dim lpTemplateName As String
        Dim hInstance As Integer
        'Dim hInstance As IntPtr
        Dim lpszStyle As String
        Dim nFontType As Integer
        Dim MISSING_ALIGNMENT As Integer
        Dim nSizeMin As Integer
        Dim nSizeMax As Integer
    End Structure

    Private Structure OPENFILENAME
        Dim lStructSize As Integer
        Dim hwndOwner As Integer
        Dim hInstance As Integer
        Dim lpstrFilter As String
        Dim lpstrCustomFilter As String
        Dim nMaxCustFilter As Integer
        Dim nFilterIndex As Integer
        Dim lpstrFile As String
        Dim nMaxFile As Integer
        Dim lpstrFileTitle As String
        Dim nMaxFileTitle As Integer
        Dim lpstrInitialDir As String
        Dim lpstrTitle As String
        Dim Flags As FileOpenConstants
        Dim nFileOffset As Integer
        Dim nFileExtension As Integer
        Dim lpstrDefExt As String
        Dim lCustData As Integer
        Dim lpfnHook As Integer
        Dim lpTemplateName As String
    End Structure

    Private Declare Function ChooseFont Lib "comdlg32" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTType) As Integer
    Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Integer
    Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Integer
    Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Integer, ByVal nNumerator As Integer, ByVal nDenominator As Integer) As Integer

    'Shared properties, used with multiple dialog types.
    Public Flags As Integer

    'ShowOpen, ShowSave properties.
    Public DefaultExt As String 'Value excludes period.
    Public DialogTitle As String
    Public FileName As String
    Public FileTitle As String
    Public Filter As String
    Public FilterIndex As Integer
    Public InitDir As String
    Public MaxFileSize As Integer

    'ShowFont properties.
    Public Charset As Integer
    'Public Color As VBRUN.ColorConstants
    Public Color As Color
    Public FontBold As Boolean
    Public FontItalic As Boolean
    Public FontName As String
    Public FontSize As Single
    Public FontStrikeThru As Boolean
    Public FontUnderline As Boolean
    Public Max As Single
    Public Min As Single

    Private CF As CHOOSEFONTType
    Private Lf As LOGFONT
    'Private Of As OPENFILENAME
    Private Off As OPENFILENAME
    Private Declare Function VarPtrAny Lib "msvbvm60.dll" Alias "VarPtr" (ByRef lpObject As Object) As Integer

    Private Sub InitChooseFont(ByVal hwnd As Integer, ByVal hDC As Integer)
        Dim bytFaceName() As Byte
        'Dim bytFaceName(0) As String
        'Dim bytFaceName(0) As Byte
        Dim intByte As Integer

        If FontSize = 0 Then
            Lf.lfHeight = 0
        Else
            Lf.lfHeight = -MulDiv(FontSize, GetDeviceCaps(hDC, LOGPIXELSY), 72)
        End If
        Lf.lfWidth = 0
        Lf.lfWeight = IIf(FontBold, 700, 400)
        Lf.lfItalic = IIf(FontItalic, 1, 0)
        Lf.lfUnderline = IIf(FontUnderline, 1, 0)
        Lf.lfStrikeOut = IIf(FontStrikeThru, 1, 0)
        Lf.lfCharSet = Charset
        'bytFaceName(0) = StrConv(Left(FontName & New String("0", LF_FACESIZE), LF_FACESIZE), VBA.VbStrConv.vbFromUnicode)
        bytFaceName = Text.Encoding.Default.GetBytes(Left(FontName & New String("0", LF_FACESIZE), LF_FACESIZE))
        Lf.lfFaceName = bytFaceName
        For intByte = 0 To LF_FACESIZE - 1
            Lf.lfFaceName(intByte) = bytFaceName(intByte)
        Next

        CF.hDC = hDC
        CF.hwndOwner = hwnd
        CF.nSizeMax = Max
        CF.nSizeMin = Min
        CF.rgbColors = Color
        If (Flags And FontsConstants.cdlCFBoth) = 0 Then
            CF.Flags = Flags Or FontsConstants.cdlCFScreenFonts
        Else
            CF.Flags = Flags
        End If
    End Sub

    Private Sub ExtractChooseFont(ByVal hDC As Integer)
        FontSize = -MulDiv(Lf.lfHeight, 72, GetDeviceCaps(hDC, LOGPIXELSY))
        FontBold = Lf.lfWeight >= 600
        FontItalic = CBool(Lf.lfItalic)
        FontUnderline = CBool(Lf.lfUnderline)
        FontStrikeThru = CBool(Lf.lfStrikeOut)
        Charset = Lf.lfCharSet
        'FontName = StrConv(Lf.lfFaceName.ToString, VBA.VbStrConv.vbUnicode)
        'FontName = Text.Encoding.Default.GetChars(Lf.lfFaceName)
        FontName = f.Font.Name
        'FontName = Left(FontName, InStr(FontName, vbNullChar) - 1)
        'FontName = Left(FontName, InStr(FontName, "0") - 1)

        Color = CF.rgbColors
        Flags = CF.Flags
    End Sub

    Public Function ShowFont(ByVal hwnd As Integer, ByVal hDC As Integer) As Boolean
        'Returns False on Cancel or error.
        InitChooseFont(hwnd, hDC)
        'ShowFont = ChooseFont(CF) <> 0
        'If ShowFont Then
        '    ExtractChooseFont(hDC)
        'End If

        'Dim f As New FontDialog
        f.ShowColor = True
        If f.ShowDialog = DialogResult.OK Then
            ExtractChooseFont(hDC)
            ShowFont = True
        Else
            ShowFont = False
        End If
    End Function

    Private Sub InitOpenFile(ByVal hwnd As Integer)
        Off.hwndOwner = hwnd
        Off.lpstrFilter = Replace(Filter, "|", vbNullChar) & vbNullChar
        Off.nFilterIndex = FilterIndex
        Off.lpstrFile = FileName & New String("0", 256 - Len(FileName))
        Off.nMaxFile = MaxFileSize
        Off.lpstrFileTitle = New String("0", 256)
        Off.nMaxFileTitle = 256
        Off.lpstrInitialDir = InitDir
        Off.lpstrTitle = DialogTitle
        Off.Flags = Flags
        Off.lpstrDefExt = DefaultExt
    End Sub

    Private Sub ExtractOpenFile()
        FileName = Off.lpstrFile
        FileTitle = ""
        If (Flags And FileOpenConstants.cdlOFNAllowMultiselect) = 0 Then
            FileName = Left(FileName, InStr(FileName, vbNullChar) - 1)
            If (Flags And FileOpenConstants.cdlOFNNoValidate) = 0 Then
                FileTitle = Left(Off.lpstrFileTitle, InStr(Off.lpstrFileTitle, vbNullChar) - 1)
            End If
        End If
        Flags = Off.Flags
    End Sub

    Public Function ShowOpen(ByVal hwnd As Integer) As Boolean
        'Returns False on Cancel or error.
        InitOpenFile(hwnd)
        ShowOpen = GetOpenFileName(Off) <> 0
        If ShowOpen Then
            ExtractOpenFile()
        End If
    End Function

    Public Function ShowSave(ByVal hwnd As Integer) As Boolean
        'Returns False on Cancel or error.
        InitOpenFile(hwnd)
        ShowSave = GetSaveFileName(Off) <> 0
        If ShowSave Then
            ExtractOpenFile()
        End If
    End Function

    Public Sub New()
        'On Error Resume Next
        'Dim lfFaceName(LF_FACESIZE - 1) As Byte
        ReDim Lf.lfFaceName(LF_FACESIZE - 1)

        CF.lStructSize = Len(CF)
        'CF.hInstance = App.hInstance

        'CF.hInstance = Marshal.GetHINSTANCE(GetType(CommonDlgs).Module)
        'Dim i As IntPtr
        CF.hInstance = Marshal.GetHINSTANCE(GetType(Application).Module).ToInt32


        'CF.lpLogFont = VarPtrAny(Lf)
        'CF.lpLogFont = VarPtrArray(Lf)


        'Dim Handle As GCHandle
        'Handle.Free()
        'Dim g As Integer
        'Handle = GCHandle.Alloc(Lf, GCHandleType.Pinned)
        'Handle = GCHandle.Alloc(g, GCHandleType.Pinned)
        'Handle = GCHandle.Alloc(CF)
        'i = Handle.AddrOfPinnedObject
        'CF.lpLogFont = Handle.AddrOfPinnedObject.ToInt32
        'Handle.Free()
        'Dim i As IntPtr
        'i = Handle.AddrOfPinnedObject
        'CF.lpLogFont = i
        Dim NF As New StdFont
        Charset = NF.Charset
        Color = Color.Black
        FontName = NF.Name
        FontSize = NF.Size

        Off.lStructSize = Len(Off)
        'Off.hInstance = App.hInstance
        'Off.hInstance = Marshal.GetHINSTANCE(GetType(CommonDlgs).Module)
        Off.hInstance = Marshal.GetHINSTANCE(GetType(Application).Module).ToInt32

        Filter = "All files (*.*)|*.*"
        FilterIndex = 1
        MaxFileSize = 256
    End Sub
End Class


