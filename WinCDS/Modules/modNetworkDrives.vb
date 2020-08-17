Module modNetworkDrives
    Private Declare Function IsNetDrive Lib "shell32" (ByVal iDrive As Integer) As Integer
    Public Function RestoreIDrive() As Boolean
        '::::RestoreIDrive
        ':::SUMMARY
        ': Restore a previously mapped I:\ drive.
        ':::DESCRIPTION
        ': Shortcut to RestoreNetDrive(8).
        ':::RETURN
        ': Boolean
        If IsIDE() Then Exit Function ' Inventory computer was gone for a day, and load was taking too long
        RestoreIDrive = RestoreNetDrives(8)
    End Function

    Public Function RestoreNetDrives(Optional ByVal nDrive As Integer = 8) As Boolean
        '::::RestoreNetDrives
        ':::SUMMARY
        ': Restore Mapped Network Drives
        ':::DESCRIPTION
        ': Re-connect already mapped drives which are frequenly disconnected on system reboot.
        ': Runs at software startup.
        ':::PARAMETERS
        ': - nDrive - The drive letter (zero-based).  Thus, A:\ == 0.  I:\ == 8
        ':::RETURN
        ': Boolean
        On Error GoTo RestoreFail
        Dim I As Integer, D As String, P As String, L As String
        If nDrive < 0 Then
            For I = 0 To 25
                RestoreNetDrives(I)
            Next
            RestoreNetDrives = True
            Exit Function
        ElseIf nDrive > 25 Then
            Exit Function
        Else
            If IsNetDrive(nDrive) > 0 Then
                L = Chr(Asc("A") + nDrive)
                D = L & ":"
                P = GetCurrentUserSetting("Network\" & L, "RemotePath")
                DoMapDrive(D, P)
            End If
        End If

        RestoreNetDrives = (IsNetDrive(nDrive) = 1)
        Exit Function
RestoreFail:
        RestoreNetDrives = True
    End Function

    Private Sub DoMapDrive(ByVal Drive As String, ByVal Map As String)
        Dim X As Object
        On Error GoTo MapFail
        Drive = UCase(Left(Drive, 1)) & ":"
        If Left(Map, 2) <> "\\" Then Exit Sub
        X = CreateObject("WScript.Network")
        X.MapNetworkDrive(Drive, Map)
MapFail:
        X = Nothing
    End Sub


End Module
