Public Class frmBackUpGeneric
    Public Function ZipFiles(ByVal CompressPath As String, ByVal ZipDir As String, ByVal ZipFile As String, Optional ByVal Special as integer = 0, Optional ByVal FileMask As String = "") As Boolean
        'Status = "Zipping " & ZipFile & "..."   ERROR
        'Select Case modBackup.ZipType
        'Case wzt7ZIP : ZipFiles = SevenZipZipFiles(CompressPath, ZipDir, ZipFile, Special, FileMask)
        '    Case wztINFO : ZipFiles = InfoZipZipFiles(CompressPath, ZipDir, ZipFile, Special)
        '    Case wztVJCZ : ZipFiles = VJCZipFiles(CompressPath, ZipDir, ZipFile, Special)
        'Case wztNone : Err.Raise -1, , "No Valid Zip component."   ERROR
        'Case Else : DevErr "frmBackupGeneric.ZipFiles - Unknown Zip Component [" & modBackup.ZipType & "]"  ERROR
        'End Select
    End Function

End Class