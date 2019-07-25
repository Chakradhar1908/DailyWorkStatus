Module modCodeVersioning
    Public Const ProgramMajorVersion as integer = 11 ' if you're changing this, consider changing ...?
    Public Const ProgramMinorVersion as integer = 0
    Public Function WinCDSMajorVersion() as integer
        WinCDSMajorVersion = ProgramMajorVersion
    End Function
    Public Function WinCDSRevisionNumber() as integer
        If IsIDE() Then WinCDSRevisionNumber = GetWinCDSRevisionNumber : Exit Function ' If development, get next, not from EXE
        WinCDSRevisionNumber = VersionInformation(WinCDSEXEFile(True)).FileVersion_vbBuild '.FileVersion_vbRevision
    End Function
    Public Function WinCDSMinorVersion() as integer
        WinCDSMinorVersion = ProgramMinorVersion
    End Function
    Public Function WinCDSBuildNumber() as integer
        If IsIDE() Then WinCDSBuildNumber = GetWinCDSBuildNumber : Exit Function ' If development, get next, not from EXE
        WinCDSBuildNumber = VersionInformation(WinCDSEXEFile(True)).FileVersion_vbRevision '.FileVersion_vbBuild
    End Function
    Public Function GetWinCDSRevisionNumber() as integer
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
    Public Function GetWinCDSBuildNumber() as integer
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

End Module
