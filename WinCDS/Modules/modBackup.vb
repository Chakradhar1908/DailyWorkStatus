Module modBackup
    Public Function ZipFiles(ByVal CompressPath As String, ByVal ZipDir As String, ByVal ZipFile As String, Optional ByVal Special as integer = 0) As Boolean
        Dim UnloadAfter As Boolean
        UnloadAfter = Not IsFormLoaded("frmBackupGeneric")
        ZipFiles = frmBackUpGeneric.ZipFiles(CompressPath, ZipDir, ZipFile, Special)
        If UnloadAfter Then
            'Unload frmBackUpGeneric
            frmBackUpGeneric.Close()
        End If
    End Function

End Module
