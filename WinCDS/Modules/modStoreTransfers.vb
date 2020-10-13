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

    Public Function PrintTransfer(ByVal TN As String, Optional ByVal wCost As Boolean = False, Optional ByVal Juice As String = "") As Boolean
        '::::PrintTransfer
        ':::SUMMARY
        ': Used to Print Transfers.
        ':::DESCRIPTION
        ': This function is used to Print Transfer by defining properities to it.
        ':::PARAMETERS
        ': - TN - Indicates the Transfer Number.
        ': - wCost
        ': - Juice - Indicates the amount of interest to be pay, if customer has not pay the loan with in time.
        ':::RETURN
        ': Boolean
        Dim X As ADODB.Recordset, P As Object, C As Integer, R As Integer, I As Integer, cInv As CInvRec, NA As Integer, F As Integer
        Dim TLanded As Decimal, TJuice As Decimal, TTotal As Decimal

        If Not TransferNoExists(TN) Then Exit Function
        X = GetTransferRSByTN(TN)
        NA = NoOfActiveLocations

        P = OutputObject
        R = Printer.ScaleWidth
        C = R / 2
        P.FontSize = 20
        P.FontBold = True
        PrintToPosition(P, "Store Transfer #" & TN, C, 5, True)
        P.FontSize = 16
        PrintToPosition(P, StoreSettings(1).Name, C, 5, True)
        P.FontSize = 14
        P.FontBold = False
        PrintToPosition(P, StoreSettings(1).Address, C, 5, False)
        '  P.FontSize = 10
        '  PrintToPosition P, "Created: " & X("Ddate1"), R, VBRUN.AlignConstants.vbAlignLeftvbAlignright, True

        PrintToPosition(P, "", , , True)
        PrintToPosition(P, "", , , True)

        P.FontSize = 10
        P.FontBold = True
        PrintToPosition(P, "Style", 100, VBRUN.AlignConstants.vbAlignLeft)
        PrintToPosition(P, "Scheduled", 1500, VBRUN.AlignConstants.vbAlignLeft)
        PrintToPosition(P, "Status", 4000, VBRUN.AlignConstants.vbAlignLeft)
        For I = 1 To NA
            PrintToPosition(P, "L" & I, 5500 + (I - 1) * 400, VBRUN.AlignConstants.vbAlignRight)
        Next

        F = 5500 + NA * 400 + 400
        If wCost Then
            PrintToPosition(P, "Cost", F, VBRUN.AlignConstants.vbAlignRight)
            PrintToPosition(P, "Juice", F + 900, VBRUN.AlignConstants.vbAlignRight)
            PrintToPosition(P, "Total", F + 1800, VBRUN.AlignConstants.vbAlignRight)
        End If


        Dim RRR As Integer
        PrintToPosition(P, "", , , True)
        RRR = OutputObject.CurrentY
        P.FontBold = False
        On Error Resume Next
        'P.Line(50, Printer.CurrentY)-(Printer.ScaleWidth - 50, Printer.CurrentY)
        P.Line(50, Printer.CurrentY, Printer.ScaleWidth - 50, Printer.CurrentY)
        On Error GoTo 0
        OutputObject.CurrentY = RRR

        Do While Not X.EOF
            PrintToPosition(P, IfNullThenNilString(X("Style")), 100, VBRUN.AlignConstants.vbAlignLeft)
            PrintToPosition(P, IfNullThenNilString(X("Ddate1")), 1500, VBRUN.AlignConstants.vbAlignLeft)
            PrintToPosition(P, DescribeTransferStatus(IfNullThenNilString(X("Trans"))), 4000, VBRUN.AlignConstants.vbAlignLeft)

            For I = 1 To NA
                PrintToPosition(P, "" & IfNullThenZero(X("Loc" & I)), 5500 + (I - 1) * 400, VBRUN.AlignConstants.vbAlignRight)
            Next

            If wCost Then
                Dim CC As CInvRec
                CC = New CInvRec
                If CC.Load(X("Style").Value, "Style") Then
                    PrintToPosition(P, CurrencyFormat(CC.Landed), F, VBRUN.AlignConstants.vbAlignRight)
                    PrintToPosition(P, CurrencyFormat(CC.Landed * GetPercent(Juice) * 0.01), F + 900, VBRUN.AlignConstants.vbAlignRight)
                    PrintToPosition(P, CurrencyFormat(CC.Landed + GetPercent(Juice) * 0.01), F + 1800, VBRUN.AlignConstants.vbAlignRight)
                    TLanded = TLanded + CC.Landed
                    TJuice = TJuice + CC.Landed * GetPercent(Juice) * 0.01
                    TTotal = TTotal + CC.Landed + CC.Landed * GetPercent(Juice) * 0.01
                End If
                DisposeDA(CC)
            End If

            PrintToPosition(P, "", , , True)
            cInv = New CInvRec
            If cInv.Load(X("Style").Value, "Style") Then
                PrintToPosition(P, cInv.Desc, 200, , True)
            End If
            DisposeDA(cInv)

            X.MoveNext()
        Loop

        P.FontBold = True
        PrintToPosition(P, CurrencyFormat(TLanded), F, VBRUN.AlignConstants.vbAlignRight)
        PrintToPosition(P, CurrencyFormat(TJuice), F + 900, VBRUN.AlignConstants.vbAlignRight)
        PrintToPosition(P, CurrencyFormat(TTotal), F + 1800, VBRUN.AlignConstants.vbAlignRight)
        P.FontBold = False

        If OutputToPrinter Then
            P.EndDoc
        Else
            'frmPrintPreviewDocument.MousePointer = 0
            frmPrintPreviewDocument.Cursor = Cursors.Default
            frmPrintPreviewDocument.DataEnd()
        End If
        PrintTransfer = True
    End Function

    Public Function TransferNoExists(ByVal TN As String) As Boolean
        '::::TransferViewRowColor
        ':::SUMMARY
        ': Used to check whether Transfer Number is exists or not.
        ':::DESCRIPTION
        ': This fucntion is used to check whether Transfer Number is exists or not from Recordset.
        ':::PARAMETERS
        ': - TN - Indicates the Transfer Number.

        Dim R As ADODB.Recordset
        If TN = "" Then Exit Function
        On Error GoTo None
        R = GetTransferRSByTN(TN)
        TransferNoExists = (R.RecordCount <> 0)
        R.Close
        R = Nothing
None:
    End Function

    Public Function GetTransferRSByTN(ByVal TN As String, Optional ByVal ExtraSQL As String = "", Optional ByVal OrderBy As String = "ORDER BY DetailID") As ADODB.Recordset
        '::::GetTransferRSByTN
        ':::SUMMARY
        ': Used to get Transfer RecordSet using Transfer Number.
        ':::DESCRIPTION
        ': This function is used to get Transfer Recordset using Transfer Number through SQL Statement.
        ':::PARAMETERS
        ': - TN - Indicates the Transfer Number.
        ': - ExtraSQL - Indicates the extra SQL statement to get Transfer Recordset.
        ': - OrderBy - Indicates the Order By condition to fileter the records in RecordSet.
        ':::RETURN
        ': Recordset : Returns Transfer Recordset.

        '  Set GetTransferRSByTN = GetRecordsetBySQL("SELECT * FROM StoreTransfers WHERE TransferNo='" & ProtectSQL(TN) & "' " & ExtraSQL & " " & OrderBy, , GetDatabaseInventory)
        GetTransferRSByTN = GetRecordsetBySQL("SELECT * FROM Detail WHERE Trans IN ('TP','TR','TV') AND [Misc]='" & ProtectSQL(TN) & "' " & ExtraSQL & " " & OrderBy, , GetDatabaseInventory)
    End Function
End Module
