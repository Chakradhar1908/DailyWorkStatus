Module modFontInstall
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

End Module
