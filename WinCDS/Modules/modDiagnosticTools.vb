Module modDiagnosticTools
    Public Function DiagnosticErrorUpload(ByVal CrashDumpFile As String) As Boolean
        '::::DiagnosticErrorUpload
        ':::SUMMARY
        ':Uploads crash dump to secure FTP server.
        ':::DESCRIPTION
        ':  Uploads the file specified to the designated server using the store name as the directory.
        ':
        ':::RETURN
        ':  Boolean - Returns True
        ':::SEE ALSO
        ': - DiagnosticDataUpload, ReportError
        ': - cFTP, modFTP, modSetup
        DiagnosticErrorUpload = FTP_Put(CompanyURL_BARE, WEB_UPLOAD_USER, WEB_UPLOAD_PASS, Slug(StoreSettings(1).Name, 12), CrashDumpFile, True)
    End Function

End Module
