Module modJET
    'Public Function UpdateAllDBPasswords(ByVal NewPassword As String, Optional ByVal OldPassword As String = "") As Boolean
    '::::UpdateAllDBPasswords
    ':::SUMMARY
    ': Used to Update All DataBase Passwords.
    ':::DESCRIPTION
    ': This module is used to update all DataBase Passwords.
    ':::PARAMETERS
    ': - NewPassword - Indicates the New Password of DataBase.
    ': - OldPassword - Indicates the Old password of DataBase.
    ':::RETURN
    ': Boolean : Returns True.
    '    Dim I as integer
    '    SetDBPassword(GetDatabaseInventory, NewPassword, OldPassword)

    '    For I = 1 To ActiveNoOfLocations
    '        SetDBPassword(GetDatabaseAtLocation(I), NewPassword, OldPassword)
    '    Next

    '    UpdateAllDBPasswords = True
    'End Function
    'Public Function SetDBPassword(ByVal dB As String, ByVal NewPassword As String, Optional ByVal OldPassword As String = "") As Boolean
    '::::SetDBPassword
    ':::SUMMARY
    ': Used to Set DataBase password.
    ':::DESCRIPTION
    ': This function is used to set the DataBase Password.
    ':::PARAMETERS
    ': - dB - Indicates the DataBase.
    ': - NewPassword - Indicates the New Password of DataBase.
    ': - OldPassword - Indicates the Old Password oF DataBase.
    ':::RETURN
    ': Boolean - Returns True.
    ' the issue is this...
    ' We can only get the JET method to SET the password, which also access the database
    ' We can only get the DAO method to CLEAR the password, and JET errors on "installable ISAM".
    ' This we can get it to work by working with these limitation, we haven't solved the underlying issue behind either of these,
    ' but so long as this can get by, we do it.
    '    If OldPassword = NewPassword Then Exit Function

    '    If OldPassword <> "" Then
    '        SetDAOPassword(dB, "", OldPassword)
    '        OldPassword = ""
    '    End If

    '    'If NewPassword <> "" Then SetDBPassword = SetJetDBPassword(dB, NewPassword, OldPassword)
    'End Function
    'Public Function SetDAOPassword(ByVal strSourceDB As String, ByVal NewPassword As String, Optional ByVal OldPassword As String = "") As Boolean
    '::::SetDAOPassword
    ':::SUMMARY
    ': Set DSN password for DAO.
    ':::DESCRIPTION
    ': This function is used to set DAO Password.
    ':::PARAMETERS
    ': - strSourceDB - Indicates the Source DataBase.
    ': - NewPassword - Indicates the New Password of DataBase.
    ': - OldPassword - Indicates the Old Password of DataBase.
    ':::RETURN
    ': Boolean - Return True On Success

    '        SetDAOPassword = True
    '        On Error GoTo Fail
    '        'Dim D As dao.DataBase
    '        D = dao.OpenDatabase(strSourceDB, True, False, IIf(OldPassword = "", "", "MS Access;PWD=" & OldPassword))
    '        D.NewPassword(OldPassword, NewPassword)
    '        '  D.Properties("Track Name AutoCorrect Info") = False
    '        D.Close
    '        Exit Function
    'Fail:
    '        Debug.Print("Failed to set password: " & Err.Description)
    '        SetDAOPassword = False
    '        Err.Clear()
    '        Resume Next
    '    End Function

    Public Function CompactRepairJETAllDDBs(Optional ByVal ForceAll As Boolean = False) As Boolean
        ':::SUMMARY
        ':::DESCRIPTION
        ': Used to Compact and Repair JET DataBases.
        ':::PARAMETERS
        ': This function is used to compact and repair all JET DataBases .
        ': - ForceAll
        ':::RETURN
        ': Boolean : Returns True.
        Dim X() As Object, I As Integer, N As Integer
        Dim dB As String

        If Not ForceAll Then
            ArrAdd(X, "Enter Pathname")
            ArrAdd(X, "All WinCDS Databases")
            ArrAdd(X, "Inventory")
            For I = 1 To LicensedNoOfStores()
                ArrAdd(X, "Store #" & I)
            Next
            N = SelectOptionArray("Compact / Repair - JET Engine", frmSelectOption.ESelOpts.SelOpt_List, X, "Select Database")
        Else
            N = 2
        End If

        If N <= 0 Then Exit Function
        If N = 1 Then
            dB = InputBox("Path:", "Enter Database Name", InventFolder)
            If dB = "" Then Exit Function
            If Not FileExists(dB) Then
                MessageBox.Show("File does not exist:" & vbCrLf & dB)
                Exit Function
            End If
        ElseIf N = 2 Then
            For I = 1 To NoOfActiveLocations
                CompactRepairJETDatabase(GetDatabaseAtLocation(I))
            Next
            dB = GetDatabaseInventory()
        Else
            dB = IIf(N = 3, GetDatabaseInventory, GetDatabaseAtLocation(N - 3))
        End If
        '  If MsgBox("Compact and repair the following database?" & vbCrLf & DB, vbQuestion + vbOKCancel, "Confirm Compact and Repair") = vbCancel Then
        '    Exit Function
        '  End If

        CompactRepairJETDatabase(dB)

        MessageBox.Show("Complete.", "Compact And Repair - JET Engine")

        CompactRepairJETAllDDBs = True
    End Function

    Public Function CompactRepairJETDatabase(ByVal strSourceDB As String, Optional ByVal xPwd As String = "#") As Boolean
        '::::CompactRepairJETDatabase
        ':::SUMMARY
        ': Used to Compact and Repair the JET DataBase.
        ':::DESCRIPTION
        ': This function is used to Compact and Repair the JET DataBase.
        ':::PARAMETERS
        ': - strSourceDB - Indicates the Source DataBase.
        ': - xPwd
        ':::RETURN
        ': Boolean : Returns True.
        Dim strDestDB As String
        Dim JetEngine As jro.JetEngine
        Dim strSourceConnect As String
        Dim strDestConnect As String

        If xPwd = "#" Then xPwd = DatabasePassword

        ' Build connection strings for SourceConnection and
        strSourceConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strSourceDB & "'"
        If xPwd <> "" Then
            strSourceConnect = strSourceConnect & ";User ID=Admin;Jet OLDEDB:Database Password=" & xPwd & ";"
        End If

        strDestDB = TempFile(, , ".mdb", False)
        strDestConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strDestDB & "'"
        If xPwd <> "" Then
            strDestConnect = strDestConnect & ";User ID=Admin;Jet OLEDB:Encrypt Database=True;" & "Jet OLEDB:Database Password=" & xPwd & ";"
        End If
        ' DestConnection arguments.


        JetEngine = New jro.JetEngine

        ' Compact and encrypt the database specified by strSourceDB
        ' to the name and path specified by strDestDB.

        ProgressForm(0, 1, "JET - Compacting...")
        JetEngine.CompactDatabase(strSourceConnect, strDestConnect)
        ProgressForm()

        JetEngine = Nothing

        If Not FileExists(strDestDB) Then
            MessageBox.Show("Failed to create target DB." & vbCrLf & "  Src=" & strSourceDB & vbCrLf & "  Dst=" & strDestDB)
            Exit Function
        End If

        DeleteFileIfExists(strSourceDB)
        FileCopy(strDestDB, strSourceDB)
        DeleteFileIfExists(strDestDB)

        CompactRepairJETDatabase = True
    End Function

End Module
