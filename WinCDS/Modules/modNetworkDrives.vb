Module modNetworkDrives
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


End Module
