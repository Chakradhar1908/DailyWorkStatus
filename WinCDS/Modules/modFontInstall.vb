Module modFontInstall
    Declare Function AddFontResource Lib "GDI32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Integer
    Private Const HWND_BROADCAST As Integer = &HFFFF
    Private Const WM_FONTCHANGE As Integer = &H1D

    Public Function FontExists(ByVal FontName As String) As Boolean
        On Error Resume Next
        'Dim oFN As String
        Dim oFN As Font

        'oFN = MainMenu.FontName
        oFN = New Font(MainMenu.Font.Name, MainMenu.Font.Style)
        'MainMenu.FontName = FontName
        MainMenu.Font = New Font(FontName, MainMenu.Font.Style)
        'FontExists = (LCase(MainMenu.FontName) = LCase(FontName))
        FontExists = (LCase(MainMenu.Font.Name) = LCase(FontName))
        'MainMenu.FontName = oFN
        MainMenu.Font = New Font(oFN, MainMenu.Font.Style)
    End Function

    Public Function InstallFontTTFToWindows(ByVal SrcFile As String, ByVal FontName As String, Optional ByVal ForceReplace As Boolean = False) As Boolean
        Dim N As String, W As String, T As String
        On Error GoTo FontRegisterFail
        N = GetFileName(SrcFile)
        W = SpecialFolder(FolderEnum.feWindowsFonts) & "\"
        T = W & N
        If Not FileExists(SrcFile) Then Exit Function   ' file to be installed does not exist

        If FileExists(T) And ForceReplace Then          ' Copy if not there.  Replace only if forced.  Otherwise, just register.
            Kill(T)
            If FileExists(T) Then Exit Function
            FileCopy(SrcFile, T)
        End If

        If Not FileExists(T) Then Exit Function         ' If can't create the file, can't register
        InstallFontTTFToWindows = Install_TTF(T, FontName)

        Exit Function
FontRegisterFail:
        '
    End Function

    Public Function Install_TTF(ByVal FontFile As String, ByVal FontName As String) As Boolean
        Dim RET As String, S As String, SFF As String, RegFN As String

        RET = AddFontResource(FontFile)
        SFF = GetFileName(FontFile)
        RegFN = FontName & " (TrueType)"

        '  S = TempFile(, , ".bat")
        '  WriteFile S, "@ECHO OFF"
        '  WriteFile S, "XCOPY " & SFF & " %systemroot%\fonts"
        '  WriteFile S, "@REG ADD ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"" /v """ & RegFN & """ /t REG_SZ /d """ & SFF & """ /f"
        '  ShellOut.RunFileAsAdmin S
        On Error Resume Next
        Kill(S)
        On Error GoTo 0

        SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0) 'Notify other open applications of the font change
        Install_TTF = FontExists(FontName)
    End Function
End Module
