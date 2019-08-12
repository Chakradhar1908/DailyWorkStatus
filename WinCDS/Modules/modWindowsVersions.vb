Imports Microsoft.VisualBasic.Interaction
Module modWindowsVersions
    Public Function IsWinXP() As Boolean
        IsWinXP = GetWinVer() = "WinXP" And GetWinVerNumber() = WINVER_WINXP
    End Function
    Public Function GetWinVer(Optional ByVal wNumbers As Boolean = False) As String 'returns a string representing the version, ie "95", "98", "NT4", "WinXP"
        Dim OSV As OSVERSIONINFO
        Dim R as integer
        '  Dim Pos As Integer
        Dim sVer As String
        Dim sBuild As String

        Dim X As String, Y As String
        On Error Resume Next

        OSV.OSVSize = Len(OSV)
        If GetVersionEx(OSV) = 1 Then
            'PlatformId contains a value representing the OS

            X = Switch(OSV.PlatformID = VER_PLATFORM_WIN32s, "32s", OSV.PlatformID = VER_PLATFORM_WIN32_NT, "NT", OSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS, "Win", True, "UNK PLATFORM")
            X = X & " v" & OSV.dwVerMajor & "." & OSV.dwVerMinor
            X = X & " (build: " & OSV.dwBuildNumber & ")"
            Y = CStr(Left(OSV.szCSDVersion, InStr(OSV.szCSDVersion, Chr(0)) - 1))
            If Len(Y) > 0 Then X = X & Y
            GetWinVer = "[" & X & "]"

            Select Case OSV.PlatformID
'                Case VER_PLATFORM_WIN32s: GetWinVer = "32s"
                Case VER_PLATFORM_WIN32_NT
                    'dwVerMajor = NT version.
                    'dwVerMinor = minor version
                    Select Case OSV.dwVerMajor
                        Case 3
                            Select Case OSV.dwVerMinor
                                Case 0 : GetWinVer = "NT3"
                                Case 1 : GetWinVer = "NT3.1"
                                Case 5 : GetWinVer = "NT3.5"
                                Case 51 : GetWinVer = "NT3.51"
                            End Select
                        Case 4
                            Select Case OSV.dwVerMinor
                                Case 0 : GetWinVer = "Win95"
                                Case 1 : GetWinVer = "Win98"
                                Case 90 : GetWinVer = "WinME"
                            End Select
                        Case 5
                            Select Case OSV.dwVerMinor
                                Case 0 : GetWinVer = "Win2000"
                                Case 1 : GetWinVer = "WinXP"
                                Case 2 : GetWinVer = "WinServer2003 or WinXPx64" & " " & Y
                            End Select
                        Case 6
                            Select Case OSV.dwVerMinor
                                Case 0 : GetWinVer = "WinVista"
                                Case 1 : GetWinVer = "Win7"
                                Case 2 : GetWinVer = "Win8"
                                Case 3 : GetWinVer = "Win8.1 or WinServer2012R2"
                            End Select
                        Case 10
                            GetWinVer = "Win10"
                    End Select
                Case VER_PLATFORM_WIN32_WINDOWS
                    'dwVerMinor bit tells if its 95 or 98.
                    Select Case OSV.dwVerMinor
                        Case 0 : GetWinVer = "Win95"
                        Case 10 : GetWinVer = "Win98"
                        Case 90 : GetWinVer = "WinME"
                    End Select
            End Select
        End If
        If wNumbers Then GetWinVer = GetWinVer & " [" & GetWinVerNumber() & "]"
    End Function

    Public Function IsWin5() As Boolean  ' XP or Server2003
        'IsWin5 = GetWinVerMajor <= 5

    End Function

End Module
