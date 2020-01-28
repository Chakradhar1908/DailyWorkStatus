Module modHTMLHelp
    Public Function OpenCHM(Optional ByVal HelpContextID As Integer = 0, Optional ByVal AltFile As String = "") As Boolean
        'WinHelp(frmMain.hWnd, App.HelpFile, HELP_CONTEXT, ByVal CLng(1234)) will show the topic that has Topic ID 1234.
        'WinHelp(frmMain.hWnd, App.HelpFile, HELP_FINDER, ByVal 0&) will show the table of contents that you created
        Const HELP_CONTEXT As Long = 0
        Const HELP_FINDER As Long = 0

        Dim Res As Long, cmd As Long
        'If AltFile = "" Then AltFile = App.HelpFile
        If AltFile = "" Then AltFile = WinCDSHelpFile()
        'cmd = IIf(HelpContextID = 0, HH_DISPLAY_TOC, HH_HELP_CONTEXT)
        'Res = HTMLHelp(MainMenu.hwnd, AltFile, cmd, HelpContextID)
        OpenCHM = (Res = 0)
    End Function

End Module
