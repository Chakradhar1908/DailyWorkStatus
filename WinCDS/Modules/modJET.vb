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
        Dim X() As Variant, I As Long, N As Long
        Dim dB As String

        If Not ForceAll Then
            ArrAdd X, "Enter Pathname"
    ArrAdd X, "All WinCDS Databases"
    ArrAdd X, "Inventory"
    For I = 1 To LicensedNoOfStores()
                ArrAdd X, "Store #" & I
    Next
            N = SelectOptionArray("Compact / Repair - JET Engine", SelOpt_List, X, "Select Database")
        Else
            N = 2
        End If

        If N <= 0 Then Exit Function
        If N = 1 Then
            dB = InputBox("Path:", "Enter Database Name", InventFolder)
            If dB = "" Then Exit Function
            If Not FileExists(dB) Then
                MsgBox "File does not exist:" & vbCrLf & dB
      Exit Function
            End If
        ElseIf N = 2 Then
            For I = 1 To NoOfActiveLocations
                CompactRepairJETDatabase GetDatabaseAtLocation(I)
    Next
            dB = GetDatabaseInventory()
        Else
            dB = IIf(N = 3, GetDatabaseInventory, GetDatabaseAtLocation(N - 3))
        End If
        '  If MsgBox("Compact and repair the following database?" & vbCrLf & DB, vbQuestion + vbOKCancel, "Confirm Compact and Repair") = vbCancel Then
        '    Exit Function
        '  End If

        CompactRepairJETDatabase dB

  MsgBox "Complete.", vbInformation, "Compact And Repair - JET Engine", , , 5


  CompactRepairJETAllDDBs = True
    End Function

End Module
