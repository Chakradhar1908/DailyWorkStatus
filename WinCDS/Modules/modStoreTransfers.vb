Module modStoreTransfers
    Public Function TransferViewRowColor(ByVal Status As String) As Color
        '::::TransferViewRowColor
        ':::SUMMARY
        ': Used to describe color to Transfer Row.
        ':::DESCRIPTION
        ': This function is used to describe color for Transfer Row based on status of Transfer.
        ':::PARAMETERS
        ': - Status - Indicates the Status of Transfer.
        ':::RETURN
        ': Long : Returns transfer row color as a long.
        Select Case LCase(Status)
            Case "open", "tp" : TransferViewRowColor = Color.Magenta
            Case "closed", "tr" : TransferViewRowColor = Color.Black
            Case Else : TransferViewRowColor = Color.Red  'void, etc
        End Select
    End Function

    Public Function DescribeTransferStatus(ByVal Trans As String) As String
        '::::DescribeTransferStatus
        ':::SUMMARY
        ': Used to Describe Transfer Status.
        ':::DESCRIPTION
        ': This function is used to describe Transfer Status.
        ':::PARAMETERS
        ': - Trans - Indicates the Transfer Status.
        ':::RETURN
        ': String : Returns Status as a String.
        Select Case Trans
            Case "TV" : DescribeTransferStatus = "Void"
            Case "TP" : DescribeTransferStatus = "Open"
            Case Else : DescribeTransferStatus = "Closed"
        End Select
    End Function

    Public Function GetTransferRSByTID(ByVal Tid As Integer, Optional ByVal ExtraSQL As String = "", Optional ByVal OrderBy As String = "ORDER BY DetailID") As ADODB.Recordset
        '::::GetTransferRSByTID
        ':::SUMMARY
        ': Used to get Transfer RecordSet using Transfer ID.
        ':::DESCRIPTION
        ': This function is used to get Transfer Recordset using Transfer ID through SQL Statement.
        ':::PARAMETERS
        ': - Tid - Indicates the Transfer Id.
        ': - ExtraSQL - Indicates the extra SQL statement to get Transfer Recordset.
        ': - OrderBy - Indicates the Order By condition to filter the records in RecordSet.
        ':::RETURN
        ': Recordset : Returns Transfer Recordset.

        '  Set GetTransferRSByTID = GetRecordsetBySQL("SELECT * FROM StoreTransfers WHERE TransferID=" & TID & " " & ExtraSQL & " " & OrderBy, , GetDatabaseInventory)
        GetTransferRSByTID = GetRecordsetBySQL("SELECT * FROM Detail WHERE Trans IN ('TP','TR','TV') AND DetailID=" & Tid & " " & ExtraSQL & " " & OrderBy, , GetDatabaseInventory)
    End Function
End Module
