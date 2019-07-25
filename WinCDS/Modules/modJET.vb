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

End Module
