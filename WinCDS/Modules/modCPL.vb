Module modCPL
    Public Sub ViewPrinters()
        '::::ViewPrinters
        ':::SUMMARY
        ': Open Printers Control Panel
        ':::DESCRIPTION
        ': This fucntion is used to open rquired file that allows us to view available printers, using path given below.
        Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL PrintersFolder")
    End Sub

    Public Sub OpenCalculator()
        '::::OpenCalculator
        ':::SUMMARY
        ': System Calculator
        ':::DESCRIPTION
        ': This function is used to Opens Calculator.
        Shell("Calc")
    End Sub

End Module
