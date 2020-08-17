Module modAccessDBs
    Public Sub CompactRepairAccessDB(ByVal sDBFILE As String, Optional ByVal sPassword As String = "", Optional ByVal ShowProgressForm As Boolean = True)
        On Error GoTo PreDAOError
        Dim sDBPATH As String, sDBNAME As String, sDB As String, sDBtmp As String
        sDBNAME = sDBFILE 'extrapulate the file name

        Do While InStr(1, sDBNAME, "\") <> 0
            sDBNAME = Right(sDBNAME, Len(sDBNAME) - InStr(1, sDBNAME, "\"))
        Loop

        'get the path name only
        sDBPATH = Left(sDBFILE, Len(sDBFILE) - Len(sDBNAME))

        sDB = sDBPATH & sDBNAME
        sDBtmp = sDBPATH & "tmp" & sDBNAME

        If ShowProgressForm Then ProgressForm(0, 1, "Working...")
        'Call the statement to execute compact and repair...

        Dim de As New DAO.DBEngine
        On Error GoTo NoDAO
        If sPassword <> "" Then
            'DAO._DBEngine.CompactDatabase(sDB, sDBtmp, DAO.LanguageConstants.dbLangGeneral, , ";pwd=" & sPassword)
            de.CompactDatabase(sDB, sDBtmp, DAO.LanguageConstants.dbLangGeneral, , ";pwd=" & sPassword)
        Else
            'DAO.DBEngine.CompactDatabase sDB, sDBtmp
            de.CompactDatabase(sDB, sDBtmp)
        End If
        On Error GoTo PostDAOError

        Application.DoEvents()            'wait for the app to finish
        If ShowProgressForm Then ProgressForm()

        Kill(sDB)            'remove the uncompressed original
        'Name sDBtmp As sDB  'rename the compressed file to the original to restore for other functions
        My.Computer.FileSystem.MoveFile(sDBtmp, sDB)
        Exit Sub
NoDAO:
        ProgressForm()
        MessageBox.Show("Call to DAO.DBEngine.CompactDatabase failed." & vbCrLf & "Either the Database is corrupted beyond repair, or" & vbCrLf & "you do not have DAO installed?" & vbCrLf & "[" & Err.Number & "]: " & Err.Description, "Could not Compact / Repair DB")
        Exit Sub
PreDAOError:
        ProgressForm()
        MessageBox.Show("DAO Compact Initialization failed." & vbCrLf & "[" & Err.Number & "]: " & Err.Description, "Could not Compact / Repair DB")
        Exit Sub
PostDAOError:
        ProgressForm()
        MessageBox.Show("DAO Compact Clean-up failed." & vbCrLf & "[" & Err.Number & "]: " & Err.Description, "Could not Compact / Repair DB")
        Exit Sub
    End Sub

End Module
