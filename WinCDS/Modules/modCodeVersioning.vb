Module modCodeVersioning
    Public Const ProgramMajorVersion as integer = 11 ' if you're changing this, consider changing ...?
    Public Const ProgramMinorVersion as integer = 0

    Public Function WinCDSMajorVersion() As Integer
        WinCDSMajorVersion = ProgramMajorVersion
    End Function

    Public Function WinCDSRevisionNumber() As Integer
        If IsIDE() Then WinCDSRevisionNumber = GetWinCDSRevisionNumber() : Exit Function ' If development, get next, not from EXE
        WinCDSRevisionNumber = VersionInformation(WinCDSEXEFile(True)).FileVersion_vbBuild '.FileVersion_vbRevision
    End Function

    Public Function WinCDSMinorVersion() As Integer
        WinCDSMinorVersion = ProgramMinorVersion
    End Function

    Public Function WinCDSBuildNumber() As Integer
        If IsIDE() Then WinCDSBuildNumber = GetWinCDSBuildNumber() : Exit Function ' If development, get next, not from EXE
        WinCDSBuildNumber = VersionInformation(WinCDSEXEFile(True)).FileVersion_vbRevision '.FileVersion_vbBuild
    End Function

    Public Function GetWinCDSRevisionNumber() As Integer
        'BFH20140304
        '  Revision number is number of days since 1/1/2012, plus 2000,
        '  to keep conflict with previous versioning.
        '    1 /1 /2001 == 2000
        '    1/2/2001 == 2001
        '    ...
        '    11/25/2033 == 9999
        '  Theoretically, this will likely never need to be reset... However, it can
        '  be re-set at any update of the major or minor version numbers.
        '
        '  To keep the version fitting within the file identifier, however, the value should
        '  probably be kept to 4 digits:  XXXX
        GetWinCDSRevisionNumber = (2000 + DateDiff("d", #1/1/2012#, Today))
    End Function

    Public Function GetWinCDSBuildNumber() As Integer
        Dim S As String, F As String
        If Not IsIDE() Then GetWinCDSBuildNumber = WinCDSBuildNumber() : 
        Exit Function
        S = ReadFile(WinCDSBuildNumberFile)
        GetWinCDSBuildNumber = Val(S)
        If GetWinCDSBuildNumber = 0 Then GetWinCDSBuildNumber = Val(VersionShouldBe(True))
    End Function

    Private Function WinCDSBuildNumberFile() As String
        WinCDSBuildNumberFile = AppFolder() & "Build.txt"
    End Function

    Public Function VersionShouldBe(Optional ByVal EXEBackOne As Boolean = False, Optional ByVal DevChannel As String = EXE_CHANNEL_PRODUCTION) As String
        VersionShouldBe = WinCDSMajorVersion() & "." & WinCDSMinorVersion() & "." & IIf(ExeChannelDescriptor(DevChannel) = "", GetWinCDSRevisionNumber, Format(ExeChannelDescriptor(DevChannel), "0000")) & "." & (GetWinCDSBuildNumber() - IIf(EXEBackOne, 1, 0))
    End Function

    Public Sub CheckCompanyInformation()
        Const Cmt As String = ProgramName & " is created and maintained by " & CompanyName & ".  For information, sales, issues, or feature requests, contact " & AdminContactName & " at " & AdminContactPhone & "."
        Const Dsc As String = ProgramDesc
        Const Tmk As String = ProgramTradeMk
        Const Prd As String = ProgramTag

        'If Application.CompanyName <> CompanyName Then DevErr("App.CompanyName <> CompanyName", vbCritical, "Developer Error - CheckCompanyInformation")
        If My.Application.Info.CompanyName <> CompanyName Then DevErr("App.CompanyName <> CompanyName", vbCritical, "Developer Error - CheckCompanyInformation")
        If My.Application.Info.Description <> Cmt Then DevErr("App.Comments <> Comments", vbCritical, "Developer Error - CheckCompanyInformation")
        If My.Application.Info.Title <> Dsc Then DevErr("App.FileDescription <> Dsc", vbCritical, "Developer Error - CheckCompanyInformation")
        If My.Application.Info.ProductName <> Prd Then DevErr("App.ProductName <> Dsc", vbCritical, "Developer Error - CheckCompanyInformation")
        If My.Application.Info.Trademark <> Tmk Then DevErr("App.LegalTrademarks <> Tmk", vbCritical, "Developer Error - CheckCompanyInformation")
        If My.Application.Info.Copyright <> SoftwareCopyright(True) Then
            If My.Application.Info.Copyright = Replace(SoftwareCopyright(True), CopyrightYear, CopyrightYear - 1) Then
                DevErr("Hey Developer!!!" & vbCrLf2 & "You missed one!!" & vbCrLf2 & "Please enter Project Menu, Projet Properties, 'Make' Tab, and edit the copyright information." & vbCrLf2 & "Current Copyright Notice:" & vbCrLf & My.Application.Info.Copyright, vbCritical, "Developer Error - CheckCompanyInformation")
            Else
                DevErr("App.LegalTrademarks <> Tmk", vbCritical, "Developer Error - CheckCompanyInformation")
            End If
        End If
    End Sub

    Public Sub CheckCopyrightDate()
        Dim Dev As String
        If Not IsDevelopment() Then Exit Sub
        If Val(CopyrightYear) = Val(Year(Today)) - 1 And Month(Today) = 1 And DateAndTime.Day(Today) <= 10 Then
            Dev = "Hey Developer!!"
            Dev = Dev & vbCrLf2 & "Go to modSetup and change CopyrightYear to " & Year(Today) & "."
            Dev = Dev & vbCrLf & "Also Go to the Project menu, select WinCDS Properties, click on the Make Tab and change the Legal Copyright notice to reflect the new year."
            Dev = Dev & vbCrLf2 & "This message is only visible to developers between January 1st and January 10th."
            Dev = Dev & vbCrLf & "This message will go away automatically as soon as the copyright year is udpated."
            If MessageBox.Show(Dev, "Developer New Year Notice", MessageBoxButtons.RetryCancel) = DialogResult.Cancel Then End
        End If
    End Sub

    Public Function CheckCertificateExpiration() As Boolean
        Dim S As String, Expiration As Date

        ' IDE only, of course.
        If Not IsIDE() Then Exit Function

        Expiration = CodeSignCertificateExpiration
        CheckCertificateExpiration = True

        If DateAfter(Today, DateAdd("m", -1, Expiration)) Then
            S = ""
            S = S & "!!!!!!!!!!!!!!!!!!  HEY DEVELOPER !!!!!!!!!!!!!!" & vbCrLf2
            S = S & "Your Code Signing Certificate is about to expire" & vbCrLf
            S = S & "On " & Expiration & "." & vbCrLf2
            S = S & "*WARNING: If you do not renew and have a working" & vbCrLf
            S = S & "certificate by then you will not be publish code" & vbCrLf
            S = S & "that is signed, and will cause client errors." & vbCrLf2
            S = S & "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            MessageBox.Show(S, "DEVELOPER WARNING -- IDE START-UP ONLY", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
            CheckCertificateExpiration = False
        End If
    End Function

    Public Function CodeSignCertificateExpiration() As Date
        '  CodeSignCertificateExpiration = #5/10/2018# ' Renewed 4/19/2018
        CodeSignCertificateExpiration = #5/10/2021#
    End Function

    Public Function CurrentVersion() As String
        CurrentVersion = VersionShouldBe(True)
    End Function
End Module
