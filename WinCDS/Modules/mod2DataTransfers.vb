Module mod2DataTransfers
    Public Function GetPendingTransfersFrom(ByVal Style As String, ByVal Loc As Integer) As Double
        '::::GetPendingTransfersFrom
        ':::SUMMARY
        ': Gets the pending transfers.
        ':::DESCRIPTION
        ': This function is used to get pending transfers from Detail table through SQL statement.
        ':::PARAMETERS
        ': - Style - Indicates the style.
        ': - Loc - Indicates the location.
        ':::RETURN
        ': Double - Returns the result as a Double.
        Dim R As ADODB.Recordset, X As Double
        R = GetRecordsetBySQL("SELECT * FROM [Detail] WHERE [Style]='" & Style & "' AND [Trans]='TP' AND [Loc" & Loc & "]<0", , GetDatabaseInventory)
        X = 0
        Do While Not R.EOF
            X = X - R("Loc" & Loc).Value  ' positive result
            R.MoveNext()
        Loop
        DisposeDA(R)

        GetPendingTransfersFrom = X
    End Function

End Module
