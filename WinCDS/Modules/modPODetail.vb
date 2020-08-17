Module modPODetail
    Public Const PODetail_TABLE As String = "PO"
    Public Const PODetail_INDEX As String = "PoNo"

    Public Function GetPoNo() as integer
        Dim OldPoNo as integer
        '  This is how PO Numbers used to be generated...  Now we want to be able to specify them at will,
        '  but also have an incrementer, so it means we will need to not just look at Max() anymore.. (*sniff*)
        OldPoNo = GetTableRecordMax(File:=GetDatabaseInventory, Table:=PODetail_TABLE, Field:=PODetail_INDEX) + 1
        If OldPoNo < 2000 Then OldPoNo = 2000


        GetPoNo = GetConfigAutoNumber(PODetail_INDEX, 2000, OldPoNo - 1)

        ' BFH20050322
        '   Jerry wants to be able to manually enter a PO number...
        '   This meant no more looking at the TableRecordMax, and also having to keep incrementing
        '   if it already exists...
        '   This should be OK, but means we have to loop here in case the next PO is 2030 and someone
        '   has already created POs 2030, 2031, 2032, 2033, 2034, 2035, ...  etc..
        '   There shoulnd't be that many, as this feature will probably only be used once or twice in
        '   the lifetime of this program (why does it matter if you can specify your PO number), but
        '   it does mean we have to fireproof this routine a little bit...
        '   Also am implementing the [Config] table to make this work, which just has fieldname, value
        '   so it can be used for other config/autonumber settings...  See modConfigTable for the
        '   XXXConfigAutoNumber() functions..
        Do While PoNoExists(GetPoNo)
            GetPoNo = GetPoNo + 1
            SetConfigAutoNumber(PODetail_INDEX, GetPoNo)
        Loop
    End Function
    Public Function PoNoExists(ByVal PoNo As String) As Boolean
        Dim RS As ADODB.Recordset
        RS = GetRecordsetBySQL("SELECT * FROM PO WHERE PONO=" & Val(PoNo), , GetDatabaseInventory)
        PoNoExists = Not RS.EOF
        RS = Nothing
    End Function

End Module
