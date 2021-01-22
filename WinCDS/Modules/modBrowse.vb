Module modBrowse
    Public Function cdgFile() As AxMSComDlg.AxCommonDialog
        '::::cdgFile
        ':::SUMMARY
        ': Used to get Common Dialog File.
        ':::DESCRIPTION
        ': This function is used to set cdg file and also used as error handler.
        ':::PARAMETERS
        ':::RETURN
        ': CommonDialog
        On Error Resume Next
        cdgFile = MainMenu.cdgFile
    End Function

End Module
