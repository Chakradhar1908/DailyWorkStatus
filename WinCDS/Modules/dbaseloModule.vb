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
        Dim SQL As String, PPO As String, N as integer, RS As ADODB.Recordset
        If UseQB Then
            QBGetVendorName(POName, completeName, , , , , , , , FactEmail)
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
            RS = dbGen.getRecordset

            If Not RS.EOF Then
                GetVendorFactEmail = True

                completeName = IfNullThenNilString(RS("fldVendorName").Value)
                FactEmail = IfNullThenNilString(RS("fldFactEmail").Value)
            End If

            DisposeDA(RS, dbGen)
        End If
    End Function

    Public Function GetVendorName(ByVal POName As String _
  , ByRef completeName As String _
  , Optional ByRef Address As String = "" _
  , Optional ByRef Address2 As String = "" _
  , Optional ByRef Address3 As String = "" _
  , Optional ByRef Zip As String = "" _
  , Optional ByRef Phone As String = "" _
  , Optional ByRef Fax As String = "" _
  , Optional ByRef CompleteCode As String = "" _
  , Optional ByRef EmailAddress As String = ""
  ) As Boolean
        '::::GetVendorName
        ':::SUMMARY
        ': Get a vendor name & information (from PO name)
        ':::DESCRIPTION
        ': Given a PO name, retrieves the Vendor Name, as well as other Vendor information.
        ':::PARAMETERS
        ': - completeName - Indicates the Complete Name String.
        ': - address - Indicates the Address String.
        ': - Address2 - Indicates the Address String.
        ': - Address3 - Indicates the Address String.
        ': - Zip - Indicates the Zip Code.
        ': - Phone - Indicates the Phone Number.
        ': - Fax - Indicates the Fax Number.
        ': - CompleteCode - Indicates the Complete Code String.
        ': - EmailAddress - Indicates the Email Address.
        ':::RETURN
        ': Boolean - Returns whether the vendor exists (was found).  False otherwise.

        Dim SQL As String

        POName = Trim(POName)
        If POName = "" Then Exit Function

        SQL = "SELECT left([tblAPVendors]![fldVendorName],16) AS Expr1" _
  & " , tblAPVendors.fldVendorName, tblAPVendors.fldVendorAddress1" _
  & " , tblAPVendors.fldVendorAddress2, tblAPVendors.fldVendorAddress3" _
  & " , tblAPVendors.fldVendorZip, tblAPVendors.fldVendorPhone" _
  & " , tblAPVendors.fldVendorFax, tblAPVendors.fldVendorCode, tblAPVendors.fldFactEMail" _
  & " From tblAPVendors" _
  & " Where ((left(UCASE([tblAPVendors]![fldVendorName]), Len(""" & ProtectSQL(UCase(POName)) & """)) = """ & ProtectSQL(UCase(POName)) & """))" _
  & " ORDER BY left([tblAPVendors]![fldVendorName],16)"
        '  Debug.Print sql
        On Error GoTo AnError
        If dbGen Is Nothing Then OpenApDatabase(1)
        dbGen.SQL = SQL
        Dim RS As ADODB.Recordset
        RS = dbGen.getRecordset   '(sql)
        If (RS.RecordCount = 0) Then
            GetVendorName = False
            Exit Function
        End If

        GetVendorName = True
        CompleteCode = IfNullThenNilString(RS("fldVendorCode").Value)
        'VendCode = completeCode
        completeName = IfNullThenNilString(RS("fldVendorName").Value)
        Address = IfNullThenNilString(RS("fldVendorAddress1").Value)
        Address2 = IfNullThenNilString(RS("fldVendorAddress2").Value)
        Address3 = IfNullThenNilString(RS("fldVendorAddress3").Value)
        Zip = IfNullThenNilString(RS("fldVendorZip").Value)
        Phone = IfNullThenNilString(RS("fldVendorPhone").Value)
        Fax = IfNullThenNilString(RS("fldVendorFax").Value)
        EmailAddress = IfNullThenNilString(RS("fldFactEMail").Value)
        RS.Close()
        RS = Nothing
        'GetVendorNameSucceeded = True
        Exit Function

AnError:

        ' MsgBox "No Accounts Payable database found"
        '  GetVendorNameSucceeded = False
        GetVendorName = False
        Exit Function
    End Function

    Public Sub OpenApDatabase(Optional ByVal Location as integer = 1, Optional ByRef PostingPO As Boolean = False)
        '::::OpenApDatabase
        ':::SUMMARY
        ': Open the AP Database.
        ':::DESCRIPTION
        ': This function is used to Open the AP Database
        ':::PARAMETERS
        ': - Location - Indicates the Location.
        ': - PostingPO - Indicates the Boolean value.
        Dim N As String
        ' if you need to post to separate (spelled like karate) A/P stores
        dbClose()
        If PostingPO And StoreSettings.bPostToLoc1 Then Location = 1     ' Only post to store 1.

        If Not InRange(1, Location, Setup_MaxStores) Then
            Err.Raise(-1500, , "Invalid store location: " & Location)
            Exit Sub
        End If

        N = GetDatabaseAP(Location)
        If FileExists(N) Then dbOpen(N)
    End Sub

    Public Function dbOpen(ByVal dbaseName As String) As Boolean
        '::::dbOpen
        ':::SUMMARY
        ': Open Accounting Database.
        ':::DESCRIPTION
        ': Opens the given account module database.
        ':::PARAMETERS
        ': - dbaseName - Filename to open.
        ':::RETURN
        ': Returns True
        On Error Resume Next
        dbClose()
        dbGen = DbAccessGeneral(dbaseName)
        dbOpen = True
    End Function

End Module
