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

    Public Function DiagnosticDataUpload(Optional ByVal Logs As Boolean = False) As Boolean
        '::::DiagnosticDataUpload
        ':::SUMMARY
        ':Uploads data to secure FTP server (for diagnostic operations).
        ':::DESCRIPTION
        ':  Scans the output folder and deletes reports more than 1 month old.
        ':
        '::Performs the following steps:
        ': 1. Confirm the operation
        ': 1. Backup the data
        ': 1. Upload the data
        ': 1. Clean up and notify
        ':
        ':::RETURN
        ':  Boolean - Returns True
        ':::SEE ALSO
        ': - DiagnosticErrorUpload
        ': - cFTP, modFTP

        Dim M As String, T As String

        M = ""
        M = "This operation will compress your data and upload the it to a secure server for testing by " & CompanyName & "."
        M = M & vbCrLf & "This may take up to 15 minutes, depending on your connection speed."
        M = M & vbCrLf2 & "Continue?"

        '  If MsgBox(M, vbOKCancel + vbExclamation, "") = vbCancel Then Exit Function

        T = TempFolder()
        ProgressForm 0, 1, "Backing up...", , , , prgIndefinite
  modBackup.BackupTo T, IIf(Logs, bkLO + bkSS, bkSS + bkLO + bkPS + bkAP + bkGL + bkBK + bkPR)

  ProgressForm 0, 1, "Connecting...", , , , prgSpin
  FTP_PutDir CompanyURL_BARE, WEB_UPLOAD_USER, WEB_UPLOAD_PASS, Slug(StoreSettings(1).Name, 12) & "/" & DateStamp(), T, True, True

  RemoveFolder T
  ProgressForm()

        MsgBox "Complete!", vbInformation, "Data Upload Finished", , , 25
  DiagnosticDataUpload = True
    End Function

End Module
