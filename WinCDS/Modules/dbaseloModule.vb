Module dbaseloModule
    '::::dbaseIoModule.bas
    ':::SUMMARY
    ': This module contains functions for the functioning of A/P and to perform multiple operations with its Database.
    ':::DESCRIPTION
    ': This module contains functions required to open / close the AP database.
    ': Also contains functions to update AP Transaction, Bank Account, Invoice Data, Factory email, & Vendor Name.
    Private dbGen As CDbAccessGeneral

    Public Function dbClose() As Boolean
        '::::dbClose
        ':::SUMMARY
        ': Close Accounting DB
        ':::DESCRIPTION
        ': This function is used to close the Database.
        ':
        ': Whichever connection was previously opened by the dbOpen is closed by this function.
        ':::PARAMETERS
        ':::SEE ALSO
        ': dbOpen
        ':::RETURN
        ': Returns true
        On Error Resume Next
        dbGen.dbClose
        dbGen = Nothing
        dbClose = True
    End Function

    Public Function GetVendorFactEmail(ByVal POName As String, ByRef completeName As String, ByRef FactEmail As String) As Boolean
        '::::GetVendorFactEmail
        ':::SUMMARY
        ': Used to get the Vendor Fact Email.
        ':::DESCRIPTION
        ': This function is  used to get the Vendor Fact Email after accessing data through sql statements using parameters.
        ':::PARAMETERS
        ': - POName - Indicate sthe PO Name String.
        ': - completeName - Indicates the Complete Name String.
        ': - FactEmail - Indicates the Fact Email String.
        ':::RETURN
        ': Boolean - Returns whether it it True or False.
        Dim SQL As String, PPO As String, N As Long, RS As Recordset
        If UseQB Then
            QBGetVendorName POName, completeName, , , , , , , , FactEmail
  Else
            PPO = ProtectSQL(UCase(Trim(POName)))
            N = Len(PPO)
            If PPO = "" Then Exit Function

            SQL = ""
            SQL = SQL & "SELECT * FROM tblAPVendors"
            SQL = SQL & " Where Left(UCase(fldVendorName), " & N & ") = '" & PPO & "'"
            SQL = SQL & " ORDER BY Left(fldVendorName,16)"

            OpenApDatabase
            dbGen.SQL = SQL
    Set RS = dbGen.getRecordset
    
    If Not RS.EOF Then
                GetVendorFactEmail = True

                completeName = IfNullThenNilString(RS("fldVendorName"))
                FactEmail = IfNullThenNilString(RS("fldFactEmail"))
            End If

            DisposeDA RS, dbGen
  End If
    End Function

End Module
